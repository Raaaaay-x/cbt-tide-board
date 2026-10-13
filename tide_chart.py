#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
场上潮汐图 v1 —— 场地 vs 月台（dock）双资源瓶颈

v0（保留）:  W_f(t) = W0_f + [A_f(t) - A_f(t0)] - [L_f(t) - L_f(t0)]       已关包待发 gaylord
v1 新增:
  进港车次（DI / 入库）: 到达 → 卸货开始（场外排队）→ 卸货结束（占 dock）
  待分拣池（可选）:      P(t) = P0 + U(t) - S(t)    U=卸下件数, S=分拣件数（日报 / 手工）
  场地利用率:            u_Y(t) = [W(t) + α·P(t)/ppg_in] / K           α=1 当『进港货计入K=是』
  dock 占用:             进港卸货 + 出港现装 + 出港甩挂；van 优先上坡道，其余占 dock
  逐时段判定:            场地瓶颈 / Dock瓶颈 / 双瓶颈 / 宽松 → 每小时 可甩挂上限、进港可接收余量

详细说明见 潮汐图_模型说明.md

用法:
  python3 tide_chart.py --input 潮汐图_输入模板.xlsx --start "2026-09-01 14:00" --end "2026-09-02 14:00" \
      [--tms 出港TMS_清洗A.xlsx] [--tms-in 进港TMS.xlsx] [--wms 关包明细.xlsx] [--daily 日报收集结果.xlsx] [--out ./out]
依赖: pip3 install pandas openpyxl matplotlib
"""
import argparse
import math
import os
import re
import warnings

import numpy as np
import pandas as pd
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import matplotlib.dates as mdates
from matplotlib import font_manager
from matplotlib.patches import Patch

warnings.filterwarnings("ignore")

_FONTS = ["PingFang SC", "Heiti SC", "Arial Unicode MS", "Noto Sans CJK SC", "Noto Sans CJK JP"]
_avail = {f.name for f in font_manager.fontManager.ttflist}
plt.rcParams["font.sans-serif"] = [f for f in _FONTS if f in _avail] + ["DejaVu Sans"]
plt.rcParams["axes.unicode_minus"] = False

ISSUES = []
PURPOSES = ["出港现装", "进港卸货", "出港甩挂"]          # van 上坡道的优先顺序
P_COLORS = {"出港现装": "#4C78A8", "进港卸货": "#F58518", "出港甩挂": "#B279A2"}
CLASSES = ["宽松", "场地瓶颈", "Dock瓶颈", "双瓶颈", "场地未知"]
C_COLORS = {"宽松": "#FFFFFF", "场地瓶颈": "#9ECAE9", "Dock瓶颈": "#FFBF79", "双瓶颈": "#F28E8E", "场地未知": "#E0E0E0"}


def issue(kind, detail):
    ISSUES.append({"类型": kind, "说明": detail})


# ---------------- 工具 ----------------
def to_dt(s):
    if not isinstance(s, pd.Series):
        s = pd.Series(s)
    s = s.astype(str).str.replace(r"\s*PT\s*$", "", regex=True).replace({"--": None, "nan": None, "None": None, "NaT": None, "": None})
    return pd.to_datetime(s, errors="coerce")


def to_num(s):
    return pd.to_numeric(s, errors="coerce")


def nz(v, default):
    try:
        return default if v is None or (isinstance(v, float) and math.isnan(v)) or str(v).strip() == "" else v
    except Exception:
        return default


def drop_examples(df):
    if df is None:
        return None
    df = df.dropna(how="all")
    if "备注" in df.columns:
        return df[~df["备注"].astype(str).str.contains("示例", na=False)].copy()
    return df


def ramp(t, t0, t1, qty):
    """累计贡献：t0→t1 匀速；t1<=t0 时为 t0 处阶跃"""
    t = np.asarray(t, dtype="int64")
    a, b = t0.value, t1.value
    if b <= a:
        return qty * (t >= a).astype(float)
    return qty * np.clip((t - a) / (b - a), 0.0, 1.0)


def inc(g, a, b, q, t0):
    """窗口内相对 t0 的累计增量"""
    return ramp(g, a, b, q) - ramp(np.array([t0.value]), a, b, q)[0]


def default_flow(route):
    m = re.match(r"LAXSCBT\((\w+)\)", str(route))
    if m:
        return m.group(1)
    m = re.match(r"LAX\d*-(\w+)", str(route))
    if m:
        return "DI-" + m.group(1)
    return str(route)


def find_col(df, candidates, override=None, required=True, label=""):
    if override and override in df.columns:
        return override
    if override:
        issue("列映射", f"{label}：指定列名『{override}』不在文件中，改用自动识别")
    for c in candidates:
        if c in df.columns:
            return c
    if required:
        raise SystemExit(f"[错误] 找不到{label}列，候选: {candidates}；实际列: {list(df.columns)[:20]}…\n"
                         f"→ 在模板『列映射_进港』里填上你导出文件的实际列名")
    return None


def is_van(v):
    return str(v).strip().lower() in {"van", "厢式", "面包", "sprinter", "小车"}


def read_sheet(x, name):
    return drop_examples(x[name]) if name in x else None


# ---------------- 输入模板 ----------------
def load_inputs(path):
    x = pd.read_excel(path, sheet_name=None)
    p = drop_examples(x["参数"]).set_index("参数")["值"].to_dict()
    g = lambda k, d: nz(p.get(k), d)
    params = {
        "K": float(g("场地gaylord位上限K", np.nan)),
        "docks": float(g("月台数", 7)),
        "ramps": float(g("van坡道数", 1)),
        "bin_min": int(g("分桶分钟", 15)),
        "drop_h": float(g("甩挂判定小时", 3)),
        "ppg_local": float(g("件每gaylord_本地", 290)),
        "ppg_line": float(g("件每gaylord_干线", 430)),
        "ppg_in": float(g("件每进港容器", np.nan)),
        "load_min": float(g("默认装车时长分钟", 30)),
        "unload_min": float(g("默认卸货时长分钟", 45)),
        "unload_rate": float(g("卸货速率_容器每小时", np.nan)),
        "trailer_cap": float(g("挂车容量gaylord", np.nan)),
        "theta_Y": float(g("场地预警阈值", 0.85)),
        "reserve": float(g("dock预留余量", 1)),
        "wait_th": float(g("排队判定分钟", 15)),
        "ref_check": float(g("参考线_检查通道失守", np.nan)),
        "ref_phys": float(g("参考线_物理极限", np.nan)),
        "in_counts_K": str(g("进港货计入K", "否")).strip() in {"是", "Y", "y", "yes", "1", "True"},
    }
    rm = drop_examples(x["路线映射"])
    rmap = {}
    for _, r in rm.iterrows():
        if pd.isna(r.get("路线名称")):
            continue
        rmap[str(r["路线名称"]).strip()] = {
            "flow": str(nz(r.get("流向"), default_flow(r["路线名称"]))).strip(),
            "kind": str(nz(r.get("类型"), "本地")).strip(),
            "dir": str(nz(r.get("方向"), "出港")).strip(),
            "vtype": str(nz(r.get("车型"), "挂车")).strip(),
            "handling": str(nz(r.get("进港处理"), "待分拣")).strip(),
        }
    w0 = drop_examples(x["开班存量W0"])
    w0 = {str(r["流向"]).strip(): float(r["gaylord数"]) for _, r in w0.iterrows() if pd.notna(r["流向"]) and pd.notna(r["gaylord数"])}
    p0 = w0.pop("待分拣(件)", 0.0)
    cmap = read_sheet(x, "列映射_进港")
    colmap = {}
    if cmap is not None:
        colmap = {str(r["字段"]).strip(): str(r["导出列名"]).strip() for _, r in cmap.iterrows() if pd.notna(r.get("导出列名"))}
    return dict(params=params, rmap=rmap, w0=w0, p0=p0,
                a_man=read_sheet(x, "关包_手工"), trucks_man=read_sheet(x, "车次_手工"),
                in_man=read_sheet(x, "进港_手工"), sort_man=read_sheet(x, "分拣消耗_手工"),
                dock_sched=read_sheet(x, "月台时段"), spot=read_sheet(x, "实盘校准"), colmap=colmap)


def route_info(route, rmap, direction):
    r = str(route).strip()
    if r in rmap:
        return rmap[r]
    f = default_flow(r)
    return {"flow": f, "kind": "干线" if (f.startswith("DI-") or direction == "进港") else "本地",
            "dir": direction, "vtype": "挂车", "handling": "待分拣"}


def dedupe_manual(d, key, label):
    """同一任务编号同时出现在 TMS 和手工录入时，以手工（BOL/现场核对）为准"""
    if "src" not in d.columns:
        return d
    dup = d[key].duplicated(keep=False) & d[key].notna()
    if dup.any():
        over = d[dup & (d["src"] == "手工")][key].tolist()
        issue("手工覆盖TMS", f"{label} {len(over)} 个任务以手工录入为准：{', '.join(map(str, over[:8]))}{' …' if len(over) > 8 else ''}")
        d = d.assign(_p=(d["src"] != "手工").astype(int)).sort_values("_p").drop_duplicates(subset=[key], keep="first").drop(columns="_p")
    return d


# ---------------- TMS 原始导出（两行表头）→ 清洗A 列名 ----------------
def read_tms(path):
    """兼容两种格式：清洗A（单行表头）或 TMS 原始导出（第1行=站/环节，第2行=字段）"""
    raw = pd.read_excel(path, sheet_name=0, header=None)
    if "始发站_到达_实际抵达时间" in set(raw.iloc[0].astype(str)):
        return pd.read_excel(path, sheet_name=0)
    top, sub = raw.iloc[0], raw.iloc[1]
    if not any(str(v) == "始发站" for v in top):
        return pd.read_excel(path, sheet_name=0)
    cols, station, stage, seen = [], None, None, {}
    for t, s_ in zip(top, sub):
        t = None if pd.isna(t) else str(t).strip()
        s_ = None if pd.isna(s_) else str(s_).strip()
        if t and (t in ("始发站", "终到站") or t.startswith("途径站")):
            station, stage = t, None
            name = t
        elif station and s_:
            stage = t or stage
            name = f"{station}_{stage}_{s_}"
        else:
            name = t or s_ or f"col{len(cols)}"
        seen[name] = seen.get(name, 0) + 1
        cols.append(name if seen[name] == 1 else f"{name}#{seen[name]}")
    d = raw.iloc[2:].reset_index(drop=True)
    d.columns = cols
    issue("TMS格式", f"{os.path.basename(path)} 为原始两行表头导出，已自动展开为清洗A列名")
    return d


# ---------------- 出港车次 → L(t) ----------------
def load_outbound(tms_path, trucks_man, rmap, P):
    frames = []
    if tms_path:
        t = read_tms(tms_path)
        need = {"路线名称(代码)": "route", "状态": "status", "始发站_到达_实际抵达时间": "arr",
                "始发站_封车_封车时间": "seal", "始发站_发车_发车时间": "dep",
                "始发站_装车_容器数量": "n_cont", "始发站_装车_运单数量": "n_wb", "任务编号": "task"}
        miss = [k for k in need if k not in t.columns]
        if miss:
            raise SystemExit(f"[错误] 出港 TMS 文件缺列 {miss}。请使用『清洗A格式』（单行表头）。")
        vt = find_col(t, ["车型", "车辆类型", "发货车辆类型", "始发站_到达_车型"], required=False)
        t = t[list(need) + ([vt] if vt else [])].rename(columns={**need, **({vt: "vt_raw"} if vt else {})})
        n0 = len(t)
        t = t[t["status"].astype(str).str.strip() == "已完成"]
        issue("过滤", f"出港 TMS 共 {n0} 条，保留『已完成』{len(t)} 条")
        t["src"] = "TMS"
        frames.append(t)
    if trucks_man is not None and len(trucks_man):
        m = trucks_man.rename(columns={"路线名称": "route", "到达时间": "arr", "封车时间": "seal", "发车时间": "dep",
                                       "容器数量": "n_cont", "运单数量": "n_wb", "车型": "vt_raw"})
        for c in ["dep", "n_wb", "vt_raw"]:
            if c not in m.columns:
                m[c] = np.nan
        gen = pd.Series("出港手工-" + m.index.astype(str), index=m.index)
        m["task"] = m["任务编号"].astype(str).where(m["任务编号"].notna(), gen) if "任务编号" in m.columns else gen
        m["status"], m["src"] = "手工", "手工"
        frames.append(m[["route", "status", "arr", "seal", "dep", "n_cont", "n_wb", "task", "src", "vt_raw"]])
    if not frames:
        issue("缺出港", "没有出港车次（TMS 或 车次_手工），L(t)=0、出港不占 dock")
        return pd.DataFrame(columns=["task", "route", "flow", "kind", "arr", "out", "n_cont", "n_wb", "vtype", "drop", "purpose", "dock_s", "dock_e", "dir", "src"])
    d = pd.concat(frames, ignore_index=True)
    d = dedupe_manual(d, "task", "出港")
    if "vt_raw" not in d.columns:
        d["vt_raw"] = np.nan
    for c in ["arr", "seal", "dep"]:
        d[c] = to_dt(d[c])
    for c in ["n_cont", "n_wb"]:
        d[c] = to_num(d[c])
    info = d["route"].apply(lambda r: route_info(r, rmap, "出港"))
    d["flow"] = info.apply(lambda i: i["flow"])
    d["kind"] = info.apply(lambda i: i["kind"])
    d["vtype"] = [str(v).strip() if pd.notna(v) else i["vtype"] for v, i in zip(d["vt_raw"], info)]

    d["out"] = d["seal"].fillna(d["dep"])
    late = d["dep"].notna() & d["seal"].notna() & ((d["dep"] - d["seal"]).dt.total_seconds() > 2 * 3600)
    for _, r in d[late].iterrows():
        issue("发车补录", f"{r['task']} 封车 {r['seal']:%m-%d %H:%M} 发车 {r['dep']:%m-%d %H:%M}，按封车时间计")
    ppg = np.where(d["kind"] == "干线", P["ppg_line"], P["ppg_local"])
    est = d["n_cont"].isna() & d["n_wb"].notna()
    d.loc[est, "n_cont"] = (d.loc[est, "n_wb"] / ppg[est]).round()
    for _, r in d[est].iterrows():
        issue("容器数折算", f"{r['task']} 无容器数，按 {r['n_wb']:.0f} 件折算为 {r['n_cont']:.0f} gaylord")
    bad = d["out"].isna() | d["n_cont"].isna()
    for _, r in d[bad].iterrows():
        issue("剔除", f"出港 {r['task']} 缺封车/发车时间或数量，未计入")
    d = d[~bad].copy()

    noarr = d["arr"].isna() | (d["arr"] > d["out"])
    for _, r in d[noarr].iterrows():
        issue("到达时间异常", f"{r['task']} 无到达时间或晚于封车，按封车前 {P['load_min']:.0f} 分钟到达计")
    d.loc[noarr, "arr"] = d.loc[noarr, "out"] - pd.Timedelta(minutes=P["load_min"])
    d["dock_h"] = (d["out"] - d["arr"]).dt.total_seconds() / 3600
    d["drop"] = d["dock_h"] > P["drop_h"]
    d["purpose"] = np.where(d["drop"], "出港甩挂", "出港现装")
    d["dock_s"], d["dock_e"] = d["arr"], d["out"]
    d["dir"] = "出港"
    return d


# ---------------- 进港车次（v1 新增）----------------
IN_FIELDS = {
    "路线": ["路线名称(代码)", "路线名称", "线路名称", "路线"],
    "状态": ["状态", "任务状态"],
    "任务编号": ["任务编号", "任务单号"],
    "到达时间": ["终到站_到达_实际抵达时间", "目的站_到达_实际抵达时间", "实际抵达时间", "到达时间"],
    "卸货开始": ["终到站_卸车_卸车开始时间", "终到站_卸车_开始时间", "终到站_卸车_开始卸车时间", "卸车开始时间", "卸货开始时间", "卸货开始"],
    "卸货结束": ["终到站_卸车_卸车完成时间", "终到站_卸车_卸车结束时间", "终到站_卸车_结束时间", "终到站_卸车_完成时间",
             "卸车结束时间", "卸车完成时间", "卸货结束时间", "卸货结束"],
    "容器数量": ["终到站_卸车_容器数量", "终到站_到达_容器数量", "始发站_装车_容器数量", "容器数量"],
    "运单数量": ["终到站_卸车_运单数量", "终到站_到达_运单数量", "始发站_装车_运单数量", "运单数量"],
    "车型": ["车型", "车辆类型"],
    "进港类别": ["进港类别", "类别"],
}


def load_inbound(tms_in, in_man, rmap, colmap, P):
    frames = []
    if tms_in:
        t = read_tms(tms_in) if tms_in.endswith(("xlsx", "xls")) else pd.read_csv(tms_in)
        req = {"路线", "到达时间"}
        cols = {f: find_col(t, c, colmap.get(f), required=(f in req), label=f"进港-{f}") for f, c in IN_FIELDS.items()}
        u = pd.DataFrame({f: (t[c] if c else np.nan) for f, c in cols.items()})
        if cols["状态"]:
            n0 = len(u)
            u = u[u["状态"].astype(str).str.strip().isin(["已完成", "已到达", "已卸车"])]
            issue("过滤", f"进港 TMS 共 {n0} 条，保留 已完成/已到达/已卸车 {len(u)} 条")
        if not cols["卸货开始"] and not cols["卸货结束"]:
            issue("进港缺卸货时间", "进港 TMS 无卸货开始/结束列：dock 占用 = 到达 + 默认卸货时长，排队等待无法计算")
        u["src"] = "TMS"
        frames.append(u)
    if in_man is not None and len(in_man):
        m = in_man.rename(columns={"路线名称": "路线"})
        for f in IN_FIELDS:
            if f not in m.columns:
                m[f] = np.nan
        m["任务编号"] = m["任务编号"].fillna(pd.Series("进港手工-" + m.index.astype(str), index=m.index))
        m["src"] = "手工"
        if "进港处理" in m.columns:
            m["处理_raw"] = m["进港处理"]
        frames.append(m)
    if not frames:
        issue("缺进港", "没有进港车次（--tms-in 或 进港_手工）：dock 占用只含出港车，Dock瓶颈会被低估")
        return pd.DataFrame(columns=["task", "route", "flow", "arr", "us", "ue", "n_cont", "n_wb", "vtype", "handling",
                                     "cat", "wait_min", "unload_min", "purpose", "dock_s", "dock_e", "dir", "src", "drop"])
    u = pd.concat(frames, ignore_index=True)
    u["任务编号"] = u["任务编号"].astype(str)
    u = dedupe_manual(u, "任务编号", "进港")
    if "处理_raw" not in u.columns:
        u["处理_raw"] = np.nan
    d = pd.DataFrame({
        "task": u["任务编号"].astype(str), "route": u["路线"].astype(str).str.strip(),
        "arr": to_dt(u["到达时间"]), "us": to_dt(u["卸货开始"]), "ue": to_dt(u["卸货结束"]),
        "n_cont": to_num(u["容器数量"]), "n_wb": to_num(u["运单数量"]),
        "cat": u["进港类别"].fillna("DI").astype(str), "src": u["src"], "vt_raw": u["车型"], "h_raw": u["处理_raw"],
    })
    info = d["route"].apply(lambda r: route_info(r, rmap, "进港"))
    d["flow"] = info.apply(lambda i: i["flow"])
    d["vtype"] = [str(v).strip() if pd.notna(v) else i["vtype"] for v, i in zip(d["vt_raw"], info)]
    d["handling"] = [str(h).strip() if pd.notna(h) else i["handling"] for h, i in zip(d["h_raw"], info)]

    # 件 ↔ 容器 互补
    ppg_in = P["ppg_in"]
    both = d["n_cont"].gt(0) & d["n_wb"].notna()
    if math.isnan(ppg_in):
        ppg_in = float((d.loc[both, "n_wb"] / d.loc[both, "n_cont"]).median()) if both.sum() >= 3 else P["ppg_line"]
        issue("参数推断", f"件每进港容器 未填，按 {'进港数据中位数' if both.sum() >= 3 else '干线 ppg'} 取 {ppg_in:.0f}")
    P["ppg_in_used"] = ppg_in
    m1 = d["n_cont"].isna() & d["n_wb"].notna()
    d.loc[m1, "n_cont"] = (d.loc[m1, "n_wb"] / ppg_in).round()
    m2 = d["n_wb"].isna() & d["n_cont"].notna()
    d.loc[m2, "n_wb"] = d.loc[m2, "n_cont"] * ppg_in

    bad = d["arr"].isna() & d["us"].isna()
    for _, r in d[bad].iterrows():
        issue("剔除", f"进港 {r['task']} 无到达/卸货时间，未计入")
    d = d[~bad].copy()
    d["arr"] = d["arr"].fillna(d["us"])
    # 卸货开始缺失 → 视为到达即卸（无法看出排队）
    d["us_est"] = d["us"].isna()
    d["us"] = d["us"].fillna(d["arr"])
    early = d["us"] < d["arr"]
    for _, r in d[early].iterrows():
        issue("进港时间异常", f"{r['task']} 卸货开始早于到达，按到达即卸处理")
    d.loc[early, "us"] = d.loc[early, "arr"]
    # 卸货结束缺失 → 容器数/卸货速率，或默认时长
    miss_e = d["ue"].isna() | (d["ue"] <= d["us"])
    if miss_e.any():
        if not math.isnan(P["unload_rate"]):
            dur = (d.loc[miss_e, "n_cont"].fillna(0) / P["unload_rate"] * 60).clip(lower=10).fillna(P["unload_min"])
        else:
            dur = pd.Series(P["unload_min"], index=d.index[miss_e])
        d.loc[miss_e, "ue"] = d.loc[miss_e, "us"] + pd.to_timedelta(dur, unit="m")
        issue("卸货时长估算", f"{int(miss_e.sum())} 辆进港车无卸货结束时间，按"
              + (f"卸货速率 {P['unload_rate']:.0f} 容器/h" if not math.isnan(P['unload_rate']) else f"默认 {P['unload_min']:.0f} 分钟") + " 估算")
    d["wait_min"] = np.where(d["us_est"], np.nan, (d["us"] - d["arr"]).dt.total_seconds() / 60)
    d["unload_min"] = (d["ue"] - d["us"]).dt.total_seconds() / 60
    d["purpose"] = "进港卸货"
    d["dock_s"], d["dock_e"] = d["us"], d["ue"]
    d["dir"], d["drop"] = "进港", False
    return d


# ---------------- 关包 A(t) / 分拣 S(t) ----------------
TAG_CARRIER = {"FEDEX": "FDX", "ONTRAC": "OTC", "USPS": "USM", "UNIUNI": "UNI", "GOFO": "GOF",
               "UPS": "UPS", "VEHO": "VEHO", "SWIFTX": "SWX"}


def load_tags(paths, t1):
    """包牌导出 → 每个 gaylord 一条关包记录（时间=创建时间）。
    包号 TTLAX2{目的地}{YYMMDD}{流水}：目的地=LAX → 本地流向=尾程商代码（与 TMS 路线 LAXSCBT(XXX) 一致）；
    其他目的地（EWR/ORD/DFW/ATL）→ 流向=目的地"""
    if not paths:
        return [], None
    d = pd.concat([pd.read_excel(p) for p in paths], ignore_index=True)
    n0 = len(d)
    d = d.drop_duplicates(subset=["包号"])
    if len(d) < n0:
        issue("包牌重复", f"{n0 - len(d)} 个包号在多份导出中重复，只计一次")
    d["创建时间"] = to_dt(d["创建时间"])
    d["dest"] = d["包号"].astype(str).str[6:9]
    car = d["尾程商"].astype(str).str.strip().str.upper()
    d["flow"] = np.where(d["dest"] == "LAX", car.map(TAG_CARRIER).fillna(car), d["dest"])
    n_ts = d["创建时间"].nunique()
    issue("包牌口径", f"{len(d)} 个包牌只有 {n_ts} 个不同创建时间（批量打印）：创建时间≈gaylord 启用，真实关包更晚 → W(t) 偏早偏高")
    seq = d["包号"].astype(str).str[-5:].apply(pd.to_numeric, errors="coerce")
    for (dest, day), g in d.assign(day=d["包号"].astype(str).str[9:15], seq=seq).groupby(["dest", "day"]):
        if g["seq"].notna().any() and g["seq"].max() > len(g) * 1.2:
            issue("包牌不全", f"{dest} {day}：流水号到 {int(g['seq'].max())}，导出只有 {len(g)} 个 → 导出被筛选，请按创建日期导出全部状态")
    no_truck = sorted(set(d.loc[d["状态"].astype(str) == "已交接", "flow"]))
    recs = [(r["flow"], r["创建时间"], r["创建时间"], 1.0) for _, r in d.iterrows() if pd.notna(r["创建时间"])]
    return recs, (d, no_truck)


def load_closures(wms_path, a_man, rmap):
    recs = []
    if wms_path:
        w = pd.read_excel(wms_path, sheet_name=0) if wms_path.endswith(("xlsx", "xls")) else pd.read_csv(wms_path)
        ct = find_col(w, ["关包时间", "封包时间", "关箱时间", "close_time"], label="关包时间")
        cf = find_col(w, ["流向", "尾程商", "路线", "路线名称", "carrier"], label="流向")
        w[ct] = to_dt(w[ct])
        n0 = len(w)
        w = w.dropna(subset=[ct])
        if len(w) < n0:
            issue("关包缺时间", f"WMS 关包明细 {n0 - len(w)} 行无关包时间，已剔除")
        cid = find_col(w, ["容器号", "container_id", "gaylord号"], required=False)
        if cid:
            dup = w.duplicated(subset=[cid]).sum()
            if dup:
                issue("重复容器", f"WMS 关包明细 {dup} 个容器号重复，只计一次")
                w = w.drop_duplicates(subset=[cid])
        for _, r in w.iterrows():
            k = str(r[cf]).strip()
            recs.append((rmap[k]["flow"] if k in rmap else k, r[ct], r[ct], 1.0))
    if a_man is not None and len(a_man):
        for _, r in a_man.iterrows():
            a, b = to_dt(r["开始时间"])[0], to_dt(r["结束时间"])[0]
            if pd.isna(a) or pd.isna(r["关包gaylord数"]):
                continue
            recs.append((str(r["流向"]).strip(), a, b if pd.notna(b) else a, float(r["关包gaylord数"])))
    if not recs:
        issue("缺A(t)", "没有关包数据（WMS 或 关包_手工），A(t)=0")
    return recs


def load_sorting(daily_path, sort_man):
    """S(t)：分拣件数，段内匀速。日报列：分拣开始时间/分拣结束时间/实际CBT分拣数量"""
    recs = []
    if daily_path:
        d = pd.read_excel(daily_path) if daily_path.endswith(("xlsx", "xls")) else pd.read_csv(daily_path)
        for _, r in d.iterrows():
            a, b, q = to_dt(r.get("分拣开始时间"))[0], to_dt(r.get("分拣结束时间"))[0], to_num(pd.Series([r.get("实际CBT分拣数量")]))[0]
            if pd.notna(a) and pd.notna(b) and pd.notna(q):
                recs.append((a, b, float(q)))
    if sort_man is not None and len(sort_man):
        for _, r in sort_man.iterrows():
            a, b = to_dt(r["开始时间"])[0], to_dt(r["结束时间"])[0]
            if pd.notna(a) and pd.notna(r["分拣件数"]):
                recs.append((a, b if pd.notna(b) else a, float(r["分拣件数"])))
    return recs


# ---------------- 主计算 ----------------
def compute_yard(outb, inb, closures, sorts, w0, p0, t0, t1, P):
    grid = pd.date_range(t0, t1, freq=f"{P['bin_min']}min")
    g = grid.as_unit("ns").asi8
    xd_flows = set()
    for h in inb["handling"] if len(inb) else []:
        if str(h).startswith("直转"):
            xd_flows.add(str(h).split(":", 1)[-1].split("：", 1)[-1].strip())
    flows = sorted(set(outb["flow"]) | {c[0] for c in closures} | set(w0) | xd_flows)
    A = {f: np.zeros(len(g)) for f in flows}
    L = {f: np.zeros(len(g)) for f in flows}
    for f, a, b, q in closures:
        if b >= t0 and a <= t1:
            A[f] += inc(g, a, b, q, t0)
    ow = outb[(outb["out"] >= t0) & (outb["arr"] <= t1)]
    for _, r in ow.iterrows():
        L[r["flow"]] += inc(g, r["arr"], r["out"], r["n_cont"], t0)
    U = np.zeros(len(g))   # 进港卸下（件，待分拣）
    S = np.zeros(len(g))   # 分拣（件）
    iw = inb[(inb["ue"] >= t0) & (inb["arr"] <= t1)] if len(inb) else inb
    for _, r in iw.iterrows():
        h = str(r["handling"])
        if h.startswith("直转"):
            f = h.split(":", 1)[-1].split("：", 1)[-1].strip()
            A[f] += inc(g, r["us"], r["ue"], r["n_cont"], t0)
        elif h == "待分拣" and pd.notna(r["n_wb"]):
            U += inc(g, r["us"], r["ue"], r["n_wb"], t0)
    for a, b, q in sorts:
        if b >= t0 and a <= t1:
            S += inc(g, a, b, q, t0)
    W = {f: w0.get(f, 0.0) + A[f] - L[f] for f in flows}
    for f in flows:
        if W[f].min() < -0.5:
            issue("负存量", f"流向 {f} 最低 W={W[f].min():.0f}：W0 未填/偏小，或关包数据缺失")
    Ppool = p0 + U - S
    if P["in_counts_K"]:
        if S[-1] <= 0:
            issue("待分拣池", "进港货计入K=是，但窗口内没有分拣数据（--daily 或 分拣消耗_手工）：P(t) 只涨不落，场地利用率会偏高")
        if Ppool.min() < -1:
            issue("待分拣池", f"P(t) 最低 {Ppool.min():.0f} 件：开班待分拣存量未填，或进港车次不全（日报分拣量包含所有入库货）")
    return grid, flows, A, L, W, U, S, Ppool, ow, iw


def capacity_minutes(gm_idx, P, sched):
    D = np.full(len(gm_idx), P["docks"])
    R = np.full(len(gm_idx), P["ramps"])
    if sched is not None and len(sched):
        mins = gm_idx.hour * 60 + gm_idx.minute
        for _, r in sched.iterrows():
            try:
                a = pd.Timestamp(str(r["开始"])).hour * 60 + pd.Timestamp(str(r["开始"])).minute
                b = pd.Timestamp(str(r["结束"])).hour * 60 + pd.Timestamp(str(r["结束"])).minute
            except Exception:
                issue("月台时段", f"无法解析时段 {r['开始']}–{r['结束']}，格式应为 HH:MM")
                continue
            m = (mins >= a) & (mins < b) if a < b else ((mins >= a) | (mins < b))
            if pd.notna(r.get("月台数")):
                D[m] = float(r["月台数"])
            if pd.notna(r.get("van坡道数")):
                R[m] = float(r["van坡道数"])
    return D, R


def compute_docks(outb, inb, t0, t1, P, sched):
    """分钟级占用 → 按桶取峰值"""
    gm = pd.date_range(t0, t1, freq="1min")
    gi = gm.as_unit("ns").asi8
    nonvan = {p: np.zeros(len(gm)) for p in PURPOSES}
    van = {p: np.zeros(len(gm)) for p in PURPOSES}
    wait = np.zeros(len(gm))
    for df in [outb, inb]:
        for _, r in df.iterrows():
            s, e = np.searchsorted(gi, r["dock_s"].value), np.searchsorted(gi, r["dock_e"].value)
            (van if is_van(r["vtype"]) else nonvan)[r["purpose"]][s:e] += 1
    if len(inb):
        for _, r in inb.iterrows():
            if pd.notna(r["wait_min"]) and r["wait_min"] > P["wait_th"]:
                s, e = np.searchsorted(gi, r["arr"].value), np.searchsorted(gi, r["us"].value)
                wait[s:e] += 1
    D, R = capacity_minutes(gm, P, sched)
    ramp_left = R.copy()
    on_dock = {}
    for p in PURPOSES:                       # van 按优先顺序上坡道，剩下的占 dock
        use = np.minimum(van[p], ramp_left)
        ramp_left -= use
        on_dock[p] = nonvan[p] + van[p] - use
    used = sum(on_dock.values())
    ramp_used = R - ramp_left
    # 桶聚合
    b = P["bin_min"]
    starts = np.arange(0, len(gm), b)
    mx = lambda a: np.maximum.reduceat(a, starts)
    mn = lambda a: np.minimum.reduceat(a, starts)
    live = on_dock["出港现装"] + on_dock["进港卸货"]
    res = {
        "used": mx(used), "live": mx(live), "cap": mn(D), "ramp_used": mx(ramp_used), "ramp_cap": mn(R),
        "wait": mx(wait), "free": mn(D - used),
        **{f"dock_{p}": mx(on_dock[p]) for p in PURPOSES},
    }
    over = used - D
    if over.max() > 0:
        n_over = int((over > 0).sum())
        issue("dock超容", f"有 {n_over} 分钟同时占 dock 车数 > 月台数（最多超 {over.max():.0f} 辆）："
              "通常是到达时间被当作靠台时间（实际在场外等），或封车点晚——需要现场到发记录核对")
    # dock-小时（按用途）
    dh = {p: float(on_dock[p].sum() / 60) for p in PURPOSES}
    full_min = int(((D - used) <= P["reserve"]).sum())
    return res, dh, full_min


# ---------------- 瓶颈判定 + 时间窗规则 ----------------
def classify(grid, W, Ppool, docks, P, flows):
    Wtot = sum(W[f] for f in flows)
    Y = Wtot + (Ppool / P["ppg_in_used"] if P["in_counts_K"] else 0)
    K = P["K"]
    uY = Y / K if not math.isnan(K) else np.full(len(grid), np.nan)
    uD = docks["used"] / np.where(docks["cap"] > 0, docks["cap"], np.nan)
    y_t = uY >= P["theta_Y"]
    d_t = (docks["free"] <= P["reserve"]) | (docks["wait"] > 0)   # wait 只计排队 > 排队判定分钟 的车
    cls = np.where(np.isnan(uY), np.where(d_t, "Dock瓶颈", "场地未知"),
          np.where(y_t & d_t, "双瓶颈", np.where(y_t, "场地瓶颈", np.where(d_t, "Dock瓶颈", "宽松"))))
    if math.isnan(K):
        issue("缺K", "场地gaylord位上限K 未填：只能判断 Dock 瓶颈，场地侧标为『场地未知』")
    return Y, uY, uD, cls


def hourly_rules(grid, Y, uY, cls, docks, inb, P, t0):
    df = pd.DataFrame({"t": grid, "Y": Y, "uY": uY, "cls": cls, "used": docks["used"], "live": docks["live"],
                       "drop": docks["dock_出港甩挂"], "inb": docks["dock_进港卸货"], "cap": docks["cap"], "wait": docks["wait"]})
    df = df.iloc[:-1]                         # 末端点 = 窗口终点，不单独成桶
    df["day_hour"] = df["t"].dt.floor("h")
    per = df.groupby("day_hour").agg(Y=("Y", "max"), uY=("uY", "max"), used=("used", "max"), live=("live", "max"),
                                    drop=("drop", "max"), inb=("inb", "max"), cap=("cap", "min"), wait=("wait", "max")).reset_index()
    per["h"] = per["day_hour"].dt.hour
    q = lambda s: float(np.nanpercentile(s, 90)) if s.notna().any() else np.nan
    rows = []
    K, th, rv, cap_tr = P["K"], P["theta_Y"], P["reserve"], P["trailer_cap_used"]
    order = [(t0.hour + i) % 24 for i in range(24)]
    for h in order:
        s = per[per["h"] == h]
        if s.empty:
            continue
        Yp, live, used, cap = q(s["Y"]), q(s["live"]), q(s["used"]), float(s["cap"].min())
        drop_now = q(s["drop"])
        drop_cap = max(0, int(math.floor(cap - live - rv)))        # 该时段甩挂总上限
        drop_add = max(0, drop_cap - int(round(drop_now)))          # 还能新增
        drop_over = max(0, int(round(drop_now)) - drop_cap)         # 现有甩挂超上限
        in_room = max(0, int(math.floor(cap - used - rv)))
        excess = max(0.0, Yp - th * K) if not math.isnan(K) else np.nan
        need = int(math.ceil(excess / cap_tr)) if excess == excess and excess > 0 else 0
        w = inb[inb["arr"].dt.hour == h] if len(inb) else inb
        wm = float(w["wait_min"].mean()) if len(w) and w["wait_min"].notna().any() else np.nan
        uYp = Yp / K if not math.isnan(K) else np.nan
        yt = (uYp >= th) if uYp == uYp else None
        dt_ = (cap - used <= rv) or (q(s["wait"]) > 0)
        if yt is None:
            c = "Dock瓶颈" if dt_ else "场地未知"
        else:
            c = "双瓶颈" if (yt and dt_) else "场地瓶颈" if yt else "Dock瓶颈" if dt_ else "宽松"
        over_txt = f"；甩挂超上限 {drop_over} 辆→提前封车/拉走" if drop_over else ""
        if c == "场地瓶颈":
            rule = (f"新增甩挂 {need} 辆吸收超量 {excess:.0f}（可新增 ≤{drop_add}）" if need <= drop_add
                    else f"超量 {excess:.0f} 需 {need} 辆，只能新增 {drop_add} → 前移发车")
        elif c == "双瓶颈":
            rule = f"超量 {excess:.0f}，可新增甩挂 {drop_add} → 前移发车/进港错开" + over_txt
        elif c == "Dock瓶颈":
            rule = "进港避开" + (over_txt if drop_over else (f"；现有甩挂 {drop_now:.0f} 辆" if drop_now >= 0.5 else ""))
        elif c == "宽松":
            rule = f"可接收进港 {in_room} 辆；可新增甩挂 {drop_add}"
        else:
            rule = f"可新增甩挂 {drop_add}，进港可接收 {in_room}（场地未知）"
        rows.append({"时段": f"{h:02d}:00", "样本天数": len(s), "场地峰值P90(gaylord)": round(Yp, 1),
                     "场地利用率P90": round(uYp, 3) if uYp == uYp else None, "月台数": cap,
                     "dock占用P90": round(used, 1), "其中现装+卸货P90": round(live, 1), "其中甩挂P90": round(drop_now, 1),
                     "进港占dock峰值P90": round(q(s["inb"]), 1), "进港平均排队分钟": round(wm, 0) if wm == wm else None,
                     "甩挂上限(总)": drop_cap, "可新增甩挂": drop_add, "甩挂超上限": drop_over, "场地超量(gaylord)": round(excess, 1) if excess == excess else None,
                     "需甩挂辆": need, "进港可接收余量": in_room, "判定": c, "规则草案": rule})
    rules = pd.DataFrame(rows)
    # 连续同类时段合并成窗口
    wins = []
    for c, grp in rules.groupby((rules["判定"] != rules["判定"].shift()).cumsum()):
        h0, h1 = grp["时段"].iloc[0], grp["时段"].iloc[-1]
        h1n = f"{(int(h1[:2]) + 1) % 24:02d}:00"
        wins.append({"窗口": f"{h0}–{h1n}", "判定": grp["判定"].iloc[0], "小时数": len(grp),
                     "甩挂上限(最小)": int(grp["甩挂上限(总)"].min()), "可新增甩挂(最小)": int(grp["可新增甩挂"].min()),
                     "甩挂超上限(最大)": int(grp["甩挂超上限"].max()), "需甩挂(最大)": int(grp["需甩挂辆"].max()),
                     "进港可接收余量(最小)": int(grp["进港可接收余量"].min())})
    return rules, pd.DataFrame(wins)


# ---------------- 画图 ----------------
def fmt_ax(ax):
    ax.xaxis.set_major_formatter(mdates.DateFormatter("%m-%d\n%H:%M"))
    ax.grid(alpha=0.3)


def class_spans(ax, grid, cls, alpha=0.45):
    i = 0
    n = len(cls)
    while i < n - 1:
        j = i
        while j + 1 < n - 1 and cls[j + 1] == cls[i]:
            j += 1
        if cls[i] != "宽松":
            ax.axvspan(grid[i], grid[j + 1], color=C_COLORS[cls[i]], alpha=alpha, lw=0)
        i = j + 1


def plot_tide(grid, flows, W, Ppool, ow, P, spot, outdir, title, colors):
    Wtot = sum(W[f] for f in flows)
    fig, ax = plt.subplots(figsize=(14, 6))
    ax.stackplot(grid, [np.clip(W[f], 0, None) for f in flows], labels=flows, colors=[colors[f] for f in flows], alpha=0.85)
    ax.plot(grid, Wtot, color="black", lw=1.5, label="合计 W(t)")
    if P["in_counts_K"]:
        ax.plot(grid, Wtot + Ppool / P["ppg_in_used"], color="#F58518", lw=1.5, ls="-.", label="W + 待分拣占位")
    if not math.isnan(P["K"]):
        ax.axhline(P["K"], color="red", ls="--", lw=1.5, label=f"K={P['K']:.0f}")
        ax.axhline(P["theta_Y"] * P["K"], color="red", ls=":", lw=1, label=f"预警 {P['theta_Y']:.0%}·K")
    for key, lab, ls in [("ref_check", "检查通道失守", (0, (6, 3))), ("ref_phys", "物理极限", "-")]:
        if not math.isnan(P[key]):
            ax.axhline(P[key], color="#7F0000", ls=ls, lw=1.2, label=f"{lab} {P[key]:.0f}")
    if spot is not None and len(spot):
        ax.scatter(to_dt(spot["时间"]), spot["场上gaylord实数"], marker="X", s=90, color="red", zorder=5, label="实盘")
    top = ax.get_ylim()[1]
    for _, r in ow.iterrows():
        ax.axvline(r["out"], color="grey", lw=0.6, alpha=0.5)
        ax.text(r["out"], top * 0.98, r["flow"], rotation=90, fontsize=7, va="top", ha="right", color="grey")
    neg = any(W[f].min() < -0.5 for f in flows)
    ax.set_title(f"场上待发运 gaylord 潮汐图  {title}" + ("\n⚠ 部分流向负存量（堆叠按0画，黑线为真实合计），见『数据问题』" if neg else ""))
    ax.set_ylabel("场上 gaylord 数")
    ax.legend(loc="upper left", fontsize=8, ncol=2)
    fmt_ax(ax)
    fig.tight_layout()
    fig.savefig(os.path.join(outdir, "1_潮汐图.png"), dpi=150)
    plt.close(fig)


def plot_cum(grid, flows, A, L, W, outdir):
    Atot, Ltot = sum(A[f] for f in flows), sum(L[f] for f in flows)
    W0 = sum(W[f] for f in flows)[0]
    fig, ax = plt.subplots(figsize=(14, 5))
    ax.plot(grid, W0 + Atot, label="W0 + A(t) 累计关包", lw=2)
    ax.plot(grid, Ltot, label="L(t) 累计装车", lw=2)
    ax.fill_between(grid, Ltot, W0 + Atot, alpha=0.15, label="垂直距离 = 场上存量")
    ax.set_title("累计曲线（垂直距离=在库量，水平距离≈停留时长）")
    ax.set_ylabel("gaylord（累计）")
    ax.legend(fontsize=9)
    fmt_ax(ax)
    fig.tight_layout()
    fig.savefig(os.path.join(outdir, "2_累计曲线.png"), dpi=150)
    plt.close(fig)


def plot_gantt(grid, ow, iw, docks, P, outdir, colors):
    rows = []
    for _, r in ow.iterrows():
        rows.append(dict(s=r["dock_s"], e=r["dock_e"], ws=None, c=colors.get(r["flow"], "grey"), hatch="//" if r["drop"] else None,
                         lab=f" {r['flow']} {r['n_cont']:.0f}" + (" van" if is_van(r["vtype"]) else ""), sort=r["dock_s"]))
    for _, r in iw.iterrows():
        rows.append(dict(s=r["us"], e=r["ue"], ws=r["arr"] if r["us"] > r["arr"] else None, c=P_COLORS["进港卸货"], hatch="..",
                         lab=f" 进 {r['flow']} {r['n_cont']:.0f}" + (" van" if is_van(r["vtype"]) else ""), sort=r["arr"]))
    rows.sort(key=lambda d: d["sort"])
    h = max(4, 0.28 * len(rows) + 2)
    fig, (ax, ax2) = plt.subplots(2, 1, figsize=(14, h + 3), sharex=True, gridspec_kw={"height_ratios": [h, 3]})
    for i, d in enumerate(rows):
        if d["ws"] is not None:
            ax.barh(i, d["s"] - d["ws"], left=d["ws"], color="#D0D0D0", edgecolor="none", height=0.4)
        ax.barh(i, d["e"] - d["s"], left=d["s"], color=d["c"], hatch=d["hatch"], edgecolor="black", lw=0.4)
        ax.text(d["e"], i, d["lab"], va="center", fontsize=6.5)
    ax.set_yticks([])
    ax.invert_yaxis()
    ax.set_title("月台占用（出港：到达→封车，斜纹=甩挂；进港（标“进”）：卸货开始→结束，点纹，灰条=场外排队）")
    ax.legend(handles=[Patch(facecolor="white", edgecolor="black", hatch="//", label="出港甩挂"),
                       Patch(facecolor=P_COLORS["进港卸货"], edgecolor="black", hatch="..", label="进港卸货"),
                       Patch(facecolor="#D0D0D0", label="进港排队")], fontsize=8, loc="upper right")
    g = grid
    ax2.stackplot(g, [docks[f"dock_{p}"] for p in PURPOSES], labels=PURPOSES, colors=[P_COLORS[p] for p in PURPOSES], alpha=0.8, step="post")
    ax2.step(g, docks["cap"], where="post", color="red", ls="--", label="月台数")
    if docks["ramp_cap"].max() > 0:
        ax2.step(g, docks["ramp_used"], where="post", color="#54A24B", lw=1, label="坡道(van)在用")
    if docks["wait"].max() > 0:
        ax2.step(g, docks["wait"], where="post", color="black", lw=1, ls=":", label="进港排队车数")
    ax2.set_ylabel("车数（桶内峰值）")
    ax2.legend(fontsize=8, ncol=3, loc="upper left")
    fmt_ax(ax2)
    fig.tight_layout()
    fig.savefig(os.path.join(outdir, "3_月台甘特.png"), dpi=150)
    plt.close(fig)


def plot_bottleneck(grid, uY, uD, cls, docks, P, outdir, title):
    fig, (ax, ax2) = plt.subplots(2, 1, figsize=(14, 7.5), sharex=True, gridspec_kw={"height_ratios": [4, 2.6]})
    class_spans(ax, grid, cls)
    if not np.all(np.isnan(uY)):
        ax.plot(grid, uY, color="#1f4e9c", lw=2, label="场地利用率 W/K")
        ax.axhline(P["theta_Y"], color="#1f4e9c", ls=":", lw=1, label=f"场地预警 {P['theta_Y']:.0%}")
    ax.step(grid, uD, where="post", color="#D9480F", lw=2, label="dock 利用率 N/月台数")
    dline = (docks["cap"] - P["reserve"]) / np.where(docks["cap"] > 0, docks["cap"], np.nan)
    ax.step(grid, dline, where="post", color="#D9480F", ls=":", lw=1, label=f"dock 预警（余 {P['reserve']:.0f} 个）")
    ax.axhline(1, color="grey", lw=0.8)
    ax.set_ylabel("利用率")
    ax.set_ylim(0, max(1.15, np.nanmax(uD) * 1.05, np.nanmax(uY) * 1.05 if not np.all(np.isnan(uY)) else 0))
    h1 = [Patch(facecolor=C_COLORS[c], alpha=0.6, label=c) for c in ["场地瓶颈", "Dock瓶颈", "双瓶颈"]]
    hh, ll = ax.get_legend_handles_labels()
    ax.legend(handles=hh + h1, fontsize=8, ncol=4, loc="upper left")
    ax.set_title(f"场地 vs Dock 瓶颈逐时段判定  {title}")
    ax2.stackplot(grid, [docks[f"dock_{p}"] for p in PURPOSES], labels=PURPOSES, colors=[P_COLORS[p] for p in PURPOSES], alpha=0.85, step="post")
    ax2.step(grid, docks["cap"], where="post", color="red", ls="--", label="月台数")
    ax2.set_ylabel("dock 构成（辆）")
    ax2.legend(fontsize=8, ncol=4, loc="upper left")
    fmt_ax(ax)
    fmt_ax(ax2)
    fig.tight_layout()
    fig.savefig(os.path.join(outdir, "4_瓶颈对比.png"), dpi=150)
    plt.close(fig)


def plot_rules(rules, outdir, title):
    cols = ["场地利用率P90", "dock占用P90", "其中甩挂P90", "进港占dock峰值P90", "进港平均排队分钟", "甩挂上限(总)", "可新增甩挂", "需甩挂辆", "进港可接收余量"]
    show = [c for c in cols if c in rules.columns]
    fig, ax = plt.subplots(figsize=(max(10, 0.62 * len(rules) + 3), 0.36 * (len(show) + 1) + 1.4))
    ax.set_xlim(0, len(rules))
    ax.set_ylim(0, len(show) + 1)
    for j, (_, r) in enumerate(rules.iterrows()):
        ax.add_patch(plt.Rectangle((j, len(show)), 1, 1, color=C_COLORS[r["判定"]], ec="white"))
        ax.text(j + 0.5, len(show) + 0.5, r["判定"].replace("瓶颈", "\n瓶颈"), ha="center", va="center", fontsize=7)
        for i, c in enumerate(show):
            v = r[c]
            txt = "—" if v is None or (isinstance(v, float) and math.isnan(v)) else (f"{v:.0%}" if c == "场地利用率P90" else f"{v:g}")
            ax.add_patch(plt.Rectangle((j, len(show) - 1 - i), 1, 1, color="#F7F7F7" if i % 2 else "white", ec="#DDDDDD"))
            ax.text(j + 0.5, len(show) - 0.5 - i, txt, ha="center", va="center", fontsize=7.5)
    ax.set_xticks(np.arange(len(rules)) + 0.5)
    ax.set_xticklabels(rules["时段"], fontsize=8)
    ax.set_yticks(np.arange(len(show) + 1) + 0.5)
    ax.set_yticklabels(list(reversed(show)) + ["判定"], fontsize=8)
    ax.tick_params(length=0)
    for s in ax.spines.values():
        s.set_visible(False)
    ax.set_title(f"按小时的甩挂 / 进港时间窗（多天时取 P90）  {title}", fontsize=10)
    fig.tight_layout()
    fig.savefig(os.path.join(outdir, "5_时间窗规则.png"), dpi=150)
    plt.close(fig)


# ---------------- KPI ----------------
def kpis(grid, flows, A, L, W, Y, uY, cls, docks, dh, full_min, ow, iw, P, spot):
    Wtot = sum(W[f] for f in flows)
    i = int(np.argmax(Y))
    hours = (grid[-1] - grid[0]).total_seconds() / 3600
    shipped = sum(L[f][-1] for f in flows)
    lam = shipped / hours if hours else np.nan
    avgW = float(np.mean(Wtot))
    bh = P["bin_min"] / 60
    c = pd.Series(cls[:-1]).value_counts()
    rows = [
        ("峰值场地占用 Y*", round(float(Y[i]), 1), "gaylord" + ("（含待分拣占位）" if P["in_counts_K"] else "")),
        ("峰值时刻", grid[i].strftime("%m-%d %H:%M"), ""),
        ("场地利用率峰值 Y*/K", round(float(uY[i]), 3) if not math.isnan(P["K"]) else "未填K", ""),
        ("窗口内关包 ΣA", round(float(sum(A[f][-1] for f in flows)), 1), "gaylord"),
        ("窗口内装车 ΣL", round(float(shipped), 1), "gaylord"),
        ("平均待发在库 W̄", round(avgW, 1), "gaylord"),
        ("平均停留 W̄/λ (Little)", round(avgW / lam, 2) if lam else "—", "小时（稳态近似）"),
        ("出港车次 / 进港车次", f"{len(ow)} / {len(iw)}", "辆"),
        ("最大同时占 dock", int(docks["used"].max()), "辆（不含坡道）"),
        ("dock 满载时长（余量≤预留）", round(full_min / 60, 1), "小时"),
        ("dock-小时：出港现装 / 进港卸货 / 出港甩挂", " / ".join(f"{dh[p]:.1f}" for p in PURPOSES), "dock·h"),
        ("甩挂占 dock 比例", f"{dh['出港甩挂'] / max(sum(dh.values()), 1e-9):.0%}", "甩挂 dock·h / 全部 dock·h"),
    ]
    if len(iw):
        wv = iw["wait_min"].dropna()  # noqa
        rows += [("进港平均卸货时长", round(float(iw["unload_min"].mean()), 0), "分钟"),
                 ("进港平均排队 / 最长排队", f"{wv.mean():.0f} / {wv.max():.0f}" if len(wv) else "无卸货开始时间", "分钟"),
                 (f"排队>{P['wait_th']:.0f}分钟的进港车", int((wv > P["wait_th"]).sum()) if len(wv) else "—", "辆")]
    for k in ["场地瓶颈", "Dock瓶颈", "双瓶颈", "宽松", "场地未知"]:
        if k in c:
            rows.append((f"时长-{k}", round(c[k] * bh, 2), "小时"))
    rows.append(("挂车容量（规则换算用）", P["trailer_cap_used"], "gaylord/辆"))
    for key, lab in [("ref_check", "检查通道失守"), ("ref_phys", "物理极限")]:
        if not math.isnan(P[key]):
            over = float((Y[:-1] > P[key]).sum() * bh)
            rows.append((f"超过{lab}线时长（{P[key]:.0f}）", round(over, 2), "小时"))
    for f in flows:
        j = int(np.argmax(W[f]))
        rows.append((f"峰值在库-{f}", round(float(W[f][j]), 1), grid[j].strftime("%m-%d %H:%M")))
    if spot is not None and len(spot):
        g = grid.as_unit("ns").asi8
        for _, r in spot.iterrows():
            t = to_dt(r["时间"])[0]
            if pd.isna(t):
                continue
            mv, act = float(np.interp(t.value, g, Y)), float(r["场上gaylord实数"])
            rows.append((f"实盘校准 {t:%m-%d %H:%M}", f"模型 {mv:.0f} vs 实数 {act:.0f}", f"误差 {(mv - act) / act:+.0%}" if act else ""))
    return pd.DataFrame(rows, columns=["指标", "值", "单位/备注"])


def main():
    ap = argparse.ArgumentParser(description="场上潮汐图 v1：场地 vs dock 瓶颈")
    ap.add_argument("--input", required=True, help="潮汐图_输入模板.xlsx")
    ap.add_argument("--tms", help="出港 TMS 导出（清洗A格式）")
    ap.add_argument("--tms-in", dest="tms_in", help="进港 TMS 导出（到达 LAXSCBT 的任务）")
    ap.add_argument("--wms", help="WMS 关包明细（可选）")
    ap.add_argument("--tags", nargs="+", help="包牌导出 bag_tags_export_*.xlsx（可多个），作为关包 A(t)")
    ap.add_argument("--daily", help="日报收集结果（可选，用于待分拣池的分拣消耗 S(t)）")
    ap.add_argument("--start", required=True)
    ap.add_argument("--end", required=True)
    ap.add_argument("--out", default="./tide_out")
    ap.add_argument("--note", default="", help="标题附注，如 模拟数据")
    a = ap.parse_args()

    os.makedirs(a.out, exist_ok=True)
    t0, t1 = pd.Timestamp(a.start), pd.Timestamp(a.end)
    I = load_inputs(a.input)
    P = I["params"]
    if (t1 - t0).total_seconds() % (P["bin_min"] * 60):
        raise SystemExit("[错误] 窗口长度需是分桶分钟的整数倍")
    outb = load_outbound(a.tms, I["trucks_man"], I["rmap"], P)
    inb = load_inbound(a.tms_in, I["in_man"], I["rmap"], I["colmap"], P)
    P.setdefault("ppg_in_used", P["ppg_in"] if not math.isnan(P["ppg_in"]) else P["ppg_line"])
    if math.isnan(P["trailer_cap"]):
        dr = outb.loc[outb["drop"], "n_cont"] if len(outb) else pd.Series(dtype=float)
        P["trailer_cap_used"] = float(dr.median()) if len(dr) >= 3 else 26.0
        issue("参数推断", f"挂车容量gaylord 未填，按{'甩挂车容器数中位数' if len(dr) >= 3 else '日报出托常见值'} 取 {P['trailer_cap_used']:.0f}")
    else:
        P["trailer_cap_used"] = P["trailer_cap"]
    closures = load_closures(a.wms, I["a_man"], I["rmap"])
    tag_recs, tag_info = load_tags(a.tags, t1)
    closures += tag_recs
    if tag_recs:
        ISSUES[:] = [i for i in ISSUES if i["类型"] != "缺A(t)"]
    if tag_info is not None:
        tflows = set(outb["flow"]) if len(outb) else set()
        miss = [f for f in tag_info[1] if f not in tflows]
        if miss:
            issue("包牌无车次", f"流向 {miss} 有『已交接』包牌但 TMS 里没有对应车次（DI / SWX 不在 fromHere 导出）：这些流向的 W 只涨不落")
    sorts = load_sorting(a.daily, I["sort_man"])

    grid, flows, A, L, W, U, S, Ppool, ow, iw = compute_yard(outb, inb, closures, sorts, I["w0"], I["p0"], t0, t1, P)
    ow_d = outb[(outb["dock_e"] > t0) & (outb["dock_s"] < t1)]
    iw_d = inb[(inb["dock_e"] > t0) & (inb["arr"] < t1)] if len(inb) else inb
    docks, dh, full_min = compute_docks(ow_d, iw_d, t0, t1, P, I["dock_sched"])
    Y, uY, uD, cls = classify(grid, W, Ppool, docks, P, flows)
    rules, wins = hourly_rules(grid, Y, uY, cls, docks, iw_d, P, t0)

    title = f"{t0:%m-%d %H:%M} → {t1:%m-%d %H:%M}" + (f"（{a.note}）" if a.note else "")
    cmap = plt.get_cmap("tab20")
    colors = {f: cmap(i % 20) for i, f in enumerate(flows)}
    plot_tide(grid, flows, W, Ppool, ow, P, I["spot"], a.out, title, colors)
    plot_cum(grid, flows, A, L, W, a.out)
    plot_gantt(grid, ow_d, iw_d, docks, P, a.out, colors)
    plot_bottleneck(grid, uY, uD, cls, docks, P, a.out, title)
    plot_rules(rules, a.out, title)
    k = kpis(grid, flows, A, L, W, Y, uY, cls, docks, dh, full_min, ow, iw_d, P, I["spot"])

    ts = pd.DataFrame({"时间": grid})
    for f in flows:
        ts[f"A_{f}"], ts[f"L_{f}"], ts[f"W_{f}"] = A[f].round(2), L[f].round(2), W[f].round(2)
    ts["W_合计"] = sum(W[f] for f in flows).round(2)
    ts["待分拣P(件)"] = Ppool.round(0)
    ts["场地占用Y"] = np.round(Y, 2)
    ts["场地利用率"] = np.round(uY, 3)
    for p in PURPOSES:
        ts[f"dock_{p}"] = docks[f"dock_{p}"]
    ts["dock占用"], ts["月台数"], ts["dock利用率"] = docks["used"], docks["cap"], np.round(uD, 3)
    ts["坡道在用"], ts["进港排队车数"], ts["判定"] = docks["ramp_used"], docks["wait"], cls

    tv = ow_d[["task", "route", "flow", "kind", "vtype", "arr", "out", "dock_h", "drop", "n_cont", "n_wb", "src"]].rename(columns={
        "task": "任务", "route": "路线", "flow": "流向", "kind": "类型", "vtype": "车型", "arr": "到达", "out": "封车/离场",
        "dock_h": "停靠小时", "drop": "甩挂", "n_cont": "gaylord", "n_wb": "件数", "src": "来源"})
    ti = iw_d[["task", "route", "flow", "cat", "vtype", "handling", "arr", "us", "ue", "wait_min", "unload_min", "n_cont", "n_wb", "src"]].rename(columns={
        "task": "任务", "route": "路线", "flow": "来源站/流向", "cat": "进港类别", "vtype": "车型", "handling": "进港处理",
        "arr": "到达", "us": "卸货开始", "ue": "卸货结束", "wait_min": "排队分钟", "unload_min": "卸货分钟",
        "n_cont": "容器", "n_wb": "件数", "src": "来源"}) if len(iw_d) else pd.DataFrame()
    with pd.ExcelWriter(os.path.join(a.out, "潮汐图_结果.xlsx")) as xw:
        k.to_excel(xw, sheet_name="KPI", index=False)
        wins.to_excel(xw, sheet_name="时间窗汇总", index=False)
        rules.to_excel(xw, sheet_name="时间窗规则_按小时", index=False)
        ts.to_excel(xw, sheet_name="时序", index=False)
        tv.to_excel(xw, sheet_name="出港车次", index=False)
        ti.to_excel(xw, sheet_name="进港车次", index=False)
        pd.DataFrame(ISSUES or [{"类型": "—", "说明": "无"}]).to_excel(xw, sheet_name="数据问题", index=False)

    print(k.to_string(index=False))
    print("\n时间窗汇总:\n" + wins.to_string(index=False))
    print(f"\n数据问题 {len(ISSUES)} 条，详见 潮汐图_结果.xlsx / 数据问题")
    print(f"输出目录: {os.path.abspath(a.out)}")


if __name__ == "__main__":
    main()
