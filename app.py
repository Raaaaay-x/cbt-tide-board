#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""库容潮汐 Dashboard — 手动录入 + 直接出潮汐图 / 瓶颈判定

运行:   streamlit run app.py
依赖:   streamlit pandas openpyxl matplotlib (tide_chart.py 在同目录)

MVP 范围: 出港车次 + 实盘 + 开班存量 + 关包(可选) + K/月台数等参数
          → 复用 tide_chart.py 的计算/画图函数，页面上直接看
"""
import io
import json
import os
import sys
import tempfile
import traceback
from datetime import datetime
from pathlib import Path

import matplotlib.pyplot as plt
import numpy as np
import pandas as pd
import streamlit as st

APP_DIR = Path(__file__).parent.resolve()
sys.path.insert(0, str(APP_DIR))
import tide_chart as tc  # noqa: E402

DRAFT = APP_DIR / ".tide_draft.json"


# ---------- 时间工具 ----------
def hhmm_to_dt(s, shift_date, shift_start_h):
    """HH:MM + 班次日期 + 开始小时 → pd.Timestamp。小时 < 开始小时视为次日。"""
    if s is None or (isinstance(s, float) and np.isnan(s)) or str(s).strip() in {"", "nan", "None", "NaT"}:
        return pd.NaT
    text = str(s).strip()
    # 允许用户直接贴 "YYYY-MM-DD HH:MM"
    try:
        if " " in text or "-" in text:
            return pd.Timestamp(text)
    except Exception:
        pass
    try:
        h, m = text.split(":")[:2]
        h, m = int(h), int(m)
    except Exception:
        return pd.NaT
    base = pd.Timestamp(shift_date)
    offset = pd.Timedelta(hours=h, minutes=m)
    if h < int(shift_start_h):
        offset += pd.Timedelta(days=1)
    return base + offset


# ---------- 空表 schema ----------
OUTBOUND_COLS = ["路线名称", "车型", "到达时间", "封车时间", "容器数量", "运单数量", "BOL号", "甩挂", "备注"]
SPOT_COLS = ["时间", "场上gaylord实数", "备注"]
W0_COLS = ["流向", "gaylord数"]
CLOSURE_COLS = ["开始时间", "结束时间", "流向", "关包gaylord数", "备注"]


def empty(cols):
    return pd.DataFrame({c: pd.Series(dtype="object") for c in cols})


# ---------- 持久化 ----------
PARAM_KEYS = ["K", "docks", "ramps", "trailer_cap", "theta_Y", "reserve",
              "ref_check", "ref_phys", "bin_min", "drop_h"]


def save_draft():
    data = {
        "shift_date": str(st.session_state.shift_date),
        "shift_start_h": int(st.session_state.shift_start_h),
        "params": {k: (st.session_state[k] if k in st.session_state else None) for k in PARAM_KEYS},
        "outbound": st.session_state.outbound.to_dict("records"),
        "spot": st.session_state.spot.to_dict("records"),
        "w0": st.session_state.w0.to_dict("records"),
        "closures": st.session_state.closures.to_dict("records"),
    }
    try:
        DRAFT.write_text(json.dumps(data, ensure_ascii=False, default=str, indent=2))
    except Exception as e:
        st.sidebar.warning(f"自动保存失败: {e}")


def load_draft():
    if not DRAFT.exists():
        return None
    try:
        return json.loads(DRAFT.read_text())
    except Exception:
        return None


def init_state():
    if st.session_state.get("_init"):
        return
    d = load_draft() or {}
    sd = d.get("shift_date")
    st.session_state.shift_date = pd.Timestamp(sd).date() if sd else datetime.now().date()
    st.session_state.shift_start_h = int(d.get("shift_start_h", 14))
    defaults = {"K": 264.0, "docks": 7, "ramps": 1, "trailer_cap": 26.0,
                "theta_Y": 0.85, "reserve": 1.0,
                "ref_check": 400.0, "ref_phys": 520.0, "bin_min": 15, "drop_h": 3.0}
    p = d.get("params", {}) or {}
    for k, v in defaults.items():
        pv = p.get(k)
        st.session_state[k] = pv if pv is not None else v
    def _mk(recs, cols):
        if not recs:
            return empty(cols)
        df = pd.DataFrame(recs)
        for c in cols:
            if c not in df.columns:
                df[c] = None
        return df[cols]
    st.session_state.outbound = _mk(d.get("outbound"), OUTBOUND_COLS)
    st.session_state.spot = _mk(d.get("spot"), SPOT_COLS)
    st.session_state.w0 = _mk(d.get("w0"), W0_COLS)
    st.session_state.closures = _mk(d.get("closures"), CLOSURE_COLS)
    st.session_state._init = True
    # 首次启动没有 draft 文件时写一份，之后由各 widget 的 on_change 回调维护
    if not DRAFT.exists():
        save_draft()


# ---------- 分析 ----------
def _parse_times(df, cols, sd, sh):
    out = df.copy()
    for c in cols:
        if c in out.columns:
            out[c] = out[c].apply(lambda x: hhmm_to_dt(x, sd, sh))
    return out


def run_analysis():
    sd = st.session_state.shift_date
    sh = int(st.session_state.shift_start_h)
    t0 = pd.Timestamp(sd) + pd.Timedelta(hours=sh)
    t1 = t0 + pd.Timedelta(hours=24)

    # trucks_man 的 schema 与 tide_chart.load_outbound 读取一致（路线名称/到达时间/封车时间/容器数量/运单数量/车型）
    tm = st.session_state.outbound.copy()
    tm = tm[tm["路线名称"].notna() & (tm["路线名称"].astype(str).str.strip() != "")]
    if len(tm):
        tm = _parse_times(tm, ["到达时间", "封车时间"], sd, sh)
    else:
        tm = None

    # 实盘
    spot = st.session_state.spot.copy()
    spot = spot[spot["时间"].notna() & (spot["时间"].astype(str).str.strip() != "")]
    if len(spot):
        spot = _parse_times(spot, ["时间"], sd, sh)
        spot = spot.rename(columns={})  # keep columns
    else:
        spot = None

    # 开班存量 W0
    w0_df = st.session_state.w0.copy()
    w0_df = w0_df[w0_df["流向"].notna()]
    w0 = {str(r["流向"]).strip(): float(r["gaylord数"])
          for _, r in w0_df.iterrows() if pd.notna(r.get("gaylord数"))}

    # 关包手工
    a_man = st.session_state.closures.copy()
    a_man = a_man[a_man["流向"].notna() & a_man["开始时间"].notna()]
    if len(a_man):
        a_man = _parse_times(a_man, ["开始时间", "结束时间"], sd, sh)
    else:
        a_man = None

    K = float(st.session_state.K) if st.session_state.K else float("nan")
    P = {
        "K": K,
        "docks": float(st.session_state.docks),
        "ramps": float(st.session_state.ramps),
        "bin_min": int(st.session_state.bin_min),
        "drop_h": float(st.session_state.drop_h),
        "ppg_local": 290.0, "ppg_line": 430.0, "ppg_in": float("nan"),
        "load_min": 30.0, "unload_min": 45.0, "unload_rate": float("nan"),
        "trailer_cap": float(st.session_state.trailer_cap),
        "theta_Y": float(st.session_state.theta_Y),
        "reserve": float(st.session_state.reserve),
        "wait_th": 15.0,
        "ref_check": float(st.session_state.ref_check) if st.session_state.ref_check else float("nan"),
        "ref_phys": float(st.session_state.ref_phys) if st.session_state.ref_phys else float("nan"),
        "in_counts_K": False,
    }
    P["ppg_in_used"] = P["ppg_line"]
    P["trailer_cap_used"] = P["trailer_cap"]

    if (t1 - t0).total_seconds() % (P["bin_min"] * 60):
        raise ValueError(f"窗口 24h 需能被分桶分钟 {P['bin_min']} 整除")

    rmap = {}  # 空 → route_info 自动用 default_flow（LAXSCBT(USM) → USM 等）

    tc.ISSUES.clear()
    outb = tc.load_outbound(None, tm, rmap, P)
    inb = tc.load_inbound(None, None, rmap, {}, P)
    closures = tc.load_closures(None, a_man, rmap)

    grid, flows, A, L, W, U, S, Ppool, ow, iw = tc.compute_yard(
        outb, inb, closures, [], w0, 0.0, t0, t1, P)
    ow_d = outb[(outb["dock_e"] > t0) & (outb["dock_s"] < t1)] if len(outb) else outb
    iw_d = inb[(inb["dock_e"] > t0) & (inb["arr"] < t1)] if len(inb) else inb
    docks, dh, full_min = tc.compute_docks(ow_d, iw_d, t0, t1, P, None)
    Y, uY, uD, cls = tc.classify(grid, W, Ppool, docks, P, flows)
    rules, wins = tc.hourly_rules(grid, Y, uY, cls, docks, iw_d, P, t0)

    title = f"{t0:%m-%d %H:%M} → {t1:%m-%d %H:%M}"
    cmap = plt.get_cmap("tab20")
    colors = {f: cmap(i % 20) for i, f in enumerate(flows)}

    imgs = {}
    with tempfile.TemporaryDirectory() as td:
        if flows:
            tc.plot_tide(grid, flows, W, Ppool, ow, P, spot, td, title, colors)
            imgs["1_潮汐图"] = Path(td, "1_潮汐图.png").read_bytes()
        if len(ow_d) or len(iw_d):
            tc.plot_gantt(grid, ow_d, iw_d, docks, P, td, colors)
            imgs["3_月台甘特"] = Path(td, "3_月台甘特.png").read_bytes()
        tc.plot_bottleneck(grid, uY, uD, cls, docks, P, td, title)
        imgs["4_瓶颈对比"] = Path(td, "4_瓶颈对比.png").read_bytes()
        if len(rules):
            tc.plot_rules(rules, td, title)
            imgs["5_时间窗规则"] = Path(td, "5_时间窗规则.png").read_bytes()

    k = tc.kpis(grid, flows, A, L, W, Y, uY, cls, docks, dh, full_min,
                ow, iw_d, P, spot)
    return imgs, k, wins, rules, list(tc.ISSUES)


# ---------- 导出 xlsx（给 tide_chart.py CLI 用）----------
def export_xlsx():
    sd = st.session_state.shift_date
    sh = int(st.session_state.shift_start_h)
    tm = st.session_state.outbound.copy()
    tm = tm[tm["路线名称"].notna() & (tm["路线名称"].astype(str).str.strip() != "")]
    if len(tm):
        tm = _parse_times(tm, ["到达时间", "封车时间"], sd, sh)
        tm["数据来源"] = "手工"
    spot = st.session_state.spot.copy()
    spot = spot[spot["时间"].notna()]
    if len(spot):
        spot = _parse_times(spot, ["时间"], sd, sh)
    a_man = st.session_state.closures.copy()
    a_man = a_man[a_man["流向"].notna() & a_man["开始时间"].notna()]
    if len(a_man):
        a_man = _parse_times(a_man, ["开始时间", "结束时间"], sd, sh)

    params = pd.DataFrame([
        {"参数": "场地gaylord位上限K", "值": float(st.session_state.K)},
        {"参数": "月台数", "值": float(st.session_state.docks)},
        {"参数": "van坡道数", "值": float(st.session_state.ramps)},
        {"参数": "挂车容量gaylord", "值": float(st.session_state.trailer_cap)},
        {"参数": "场地预警阈值", "值": float(st.session_state.theta_Y)},
        {"参数": "dock预留余量", "值": float(st.session_state.reserve)},
        {"参数": "参考线_检查通道失守", "值": float(st.session_state.ref_check)},
        {"参数": "参考线_物理极限", "值": float(st.session_state.ref_phys)},
        {"参数": "分桶分钟", "值": float(st.session_state.bin_min)},
        {"参数": "甩挂判定小时", "值": float(st.session_state.drop_h)},
    ])
    rmap = pd.DataFrame(columns=["路线名称", "流向", "类型", "方向", "车型", "进港处理"])
    w0 = st.session_state.w0[st.session_state.w0["流向"].notna()]

    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as xw:
        params.to_excel(xw, sheet_name="参数", index=False)
        rmap.to_excel(xw, sheet_name="路线映射", index=False)
        w0.to_excel(xw, sheet_name="开班存量W0", index=False)
        (tm if tm is not None and len(tm) else empty(OUTBOUND_COLS + ["数据来源"])).to_excel(
            xw, sheet_name="车次_手工", index=False)
        (a_man if a_man is not None and len(a_man) else empty(CLOSURE_COLS)).to_excel(
            xw, sheet_name="关包_手工", index=False)
        (spot if spot is not None and len(spot) else empty(SPOT_COLS)).to_excel(
            xw, sheet_name="实盘校准", index=False)
    buf.seek(0)
    return buf.getvalue()


# ---------- UI ----------
def sidebar():
    with st.sidebar:
        c1, c2 = st.columns([2, 1])
        with c1:
            st.date_input("班次日期", key="shift_date", on_change=save_draft, label_visibility="visible")
        with c2:
            st.number_input("开始", 0, 23, step=1, key="shift_start_h", on_change=save_draft,
                            help="班次开始小时（默认 14）")
        c1, c2 = st.columns(2)
        with c1:
            st.number_input("K 场地上限", 0.0, 2000.0, step=10.0, key="K", on_change=save_draft)
        with c2:
            st.number_input("月台数", 1, 20, step=1, key="docks", on_change=save_draft)

        with st.expander("⚙ 高级参数", expanded=False):
            c1, c2 = st.columns(2)
            with c1:
                st.number_input("van 坡道", 0, 5, step=1, key="ramps", on_change=save_draft)
                st.number_input("场地预警", 0.0, 1.0, step=0.05, key="theta_Y", on_change=save_draft,
                                help="超过 K × 此比例 → 场地瓶颈")
                st.number_input("参考·失守", 0.0, 2000.0, step=10.0, key="ref_check", on_change=save_draft,
                                help="检查通道失守线")
                st.number_input("分桶分钟", 5, 60, step=5, key="bin_min", on_change=save_draft)
            with c2:
                st.number_input("挂车容量", 1.0, 60.0, step=1.0, key="trailer_cap", on_change=save_draft,
                                help="gaylord / 辆（规则换算用）")
                st.number_input("dock 预留", 0.0, 5.0, step=1.0, key="reserve", on_change=save_draft)
                st.number_input("参考·物理", 0.0, 2000.0, step=10.0, key="ref_phys", on_change=save_draft,
                                help="物理极限线")
                st.number_input("甩挂判定h", 0.5, 12.0, step=0.5, key="drop_h", on_change=save_draft,
                                help="靠台→封车超过此时长视为甩挂")

        st.divider()
        st.download_button("📥 导出 xlsx",
                           data=export_xlsx(),
                           file_name=f"潮汐图_输入_{st.session_state.shift_date:%Y%m%d}.xlsx",
                           use_container_width=True,
                           help="导出可供 tide_chart.py CLI 读取的 xlsx")
        if st.button("🗑 清空录入", use_container_width=True):
            st.session_state.outbound = empty(OUTBOUND_COLS)
            st.session_state.spot = empty(SPOT_COLS)
            st.session_state.w0 = empty(W0_COLS)
            st.session_state.closures = empty(CLOSURE_COLS)
            save_draft()
            st.rerun()
        st.caption(f"自动保存至 `{DRAFT.name}`")


def _persist(state_key, edited):
    """如果 data_editor 返回的 DataFrame 和原值不同，就写回 session_state 并保存"""
    cur = st.session_state[state_key]
    if not edited.equals(cur):
        st.session_state[state_key] = edited.reset_index(drop=True)
        save_draft()


def tab_outbound():
    st.caption("时间按班次规则：HH:MM。若 `小时 < 班次开始小时` 自动算次日。流向从路线名自动推断（LAXSCBT(USM) → 流向 USM）")
    edited = st.data_editor(
        st.session_state.outbound,
        num_rows="dynamic", use_container_width=True, key="ed_outbound",
        column_config={
            "路线名称": st.column_config.TextColumn("路线名称*", help="如 LAXSCBT(USM) / LAX3-ATL"),
            "车型": st.column_config.SelectboxColumn(options=["挂车", "van"], default="挂车"),
            "到达时间": st.column_config.TextColumn("靠台 HH:MM*", help="车实际靠上 dock 的时间"),
            "封车时间": st.column_config.TextColumn("封车 HH:MM*", help="BOL 签字时间 = 离场时间"),
            "容器数量": st.column_config.NumberColumn("gaylord", min_value=0, step=1),
            "运单数量": st.column_config.NumberColumn("件数", min_value=0, step=1),
            "BOL号": st.column_config.TextColumn(),
            "甩挂": st.column_config.CheckboxColumn(default=False, help="仅记录；模型按『靠台→封车 > 3h』自动判定"),
            "备注": st.column_config.TextColumn(),
        },
    )
    _persist("outbound", edited)


def tab_spot():
    st.caption("时间 HH:MM。可在收班后一次填入几个关键时刻的实盘数，用于校准模型")
    edited = st.data_editor(
        st.session_state.spot,
        num_rows="dynamic", use_container_width=True, key="ed_spot",
        column_config={
            "时间": st.column_config.TextColumn("时间 HH:MM*"),
            "场上gaylord实数": st.column_config.NumberColumn(min_value=0, step=1),
            "备注": st.column_config.TextColumn(),
        },
    )
    _persist("spot", edited)


def tab_w0():
    st.caption("班次开始时各流向的在库 gaylord。流向名需和出港车次的流向一致（LAXSCBT(USM) → USM）")
    edited = st.data_editor(
        st.session_state.w0,
        num_rows="dynamic", use_container_width=True, key="ed_w0",
        column_config={
            "流向": st.column_config.TextColumn("流向*"),
            "gaylord数": st.column_config.NumberColumn(min_value=0, step=1),
        },
    )
    _persist("w0", edited)


def tab_closures():
    st.caption("按整小时录入各流向关包数（可选）。没有 WMS 关包数据时，W(t) 只减不加，这里能让曲线上翘")
    edited = st.data_editor(
        st.session_state.closures,
        num_rows="dynamic", use_container_width=True, key="ed_closures",
        column_config={
            "开始时间": st.column_config.TextColumn("开始 HH:MM*"),
            "结束时间": st.column_config.TextColumn("结束 HH:MM", help="空则视为开始时刻的阶跃"),
            "流向": st.column_config.TextColumn("流向*"),
            "关包gaylord数": st.column_config.NumberColumn(min_value=0, step=1),
            "备注": st.column_config.TextColumn(),
        },
    )
    _persist("closures", edited)


def _kpi(k, name, default="—"):
    row = k[k["指标"] == name]
    return row["值"].iloc[0] if len(row) else default


def _render_tldr(wins):
    """从时间窗汇总生成一句话判定。严重级别 双瓶颈 > 场地瓶颈 > Dock瓶颈 > 宽松。"""
    if wins is None or not len(wins):
        st.success("✓ 24h 全时段宽松")
        return
    sev = {"双瓶颈": 3, "场地瓶颈": 2, "Dock瓶颈": 2, "宽松": 0, "场地未知": 1}
    rows = wins.assign(_s=wins["判定"].map(sev).fillna(0)).sort_values("_s", ascending=False)
    top = rows.iloc[0]
    cls, win, hrs = top["判定"], top["窗口"], int(top["小时数"])
    extra = []
    if top.get("需甩挂(最大)", 0) > 0:
        extra.append(f"需甩挂 **{int(top['需甩挂(最大)'])}** 辆")
    if top.get("甩挂超上限(最大)", 0) > 0:
        extra.append(f"甩挂超上限 **{int(top['甩挂超上限(最大)'])}** 辆")
    if top.get("进港可接收余量(最小)", 99) == 0:
        extra.append("进港无余量")
    tail = ("；" + "、".join(extra)) if extra else ""
    msg = f"**{cls}** · {win}（{hrs}h）{tail}"
    (st.error if cls == "双瓶颈" else st.warning if cls in ("场地瓶颈", "Dock瓶颈") else st.info)(msg)


def tab_analysis():
    top_l, top_r = st.columns([1, 5])
    with top_l:
        run = st.button("🔄 生成", type="primary", use_container_width=True)
    with top_r:
        st.caption("按当前所有录入即时重算；图都是现场出的（不落盘）")

    if run:
        try:
            with st.spinner("计算中…"):
                imgs, k, wins, rules, issues = run_analysis()
            st.session_state.analysis = dict(imgs=imgs, kpi=k, wins=wins, rules=rules, issues=issues)
        except Exception as e:
            st.session_state.analysis = None
            st.error(f"分析失败: {e}")
            st.code(traceback.format_exc())

    if "analysis" not in st.session_state or st.session_state.analysis is None:
        st.info("把出港车次、K 填好，点上方 **生成**")
        return
    r = st.session_state.analysis
    k = r["kpi"]

    # —— TL;DR 瓶颈判定条
    _render_tldr(r["wins"])

    # —— 4 个关键 metric
    peak = _kpi(k, "峰值场地占用 Y*")
    ratio = _kpi(k, "场地利用率峰值 Y*/K")
    peak_time = _kpi(k, "峰值时刻", "")
    max_dock = _kpi(k, "最大同时占 dock", 0)
    drop_pct = _kpi(k, "甩挂占 dock 比例", "0%")
    full_h = _kpi(k, "dock 满载时长（余量≤预留）", 0)
    docks_cap = int(st.session_state.docks)

    try:
        ratio_disp = f"{float(ratio):.0%}" if isinstance(ratio, (int, float)) else str(ratio)
    except Exception:
        ratio_disp = str(ratio)

    m1, m2, m3, m4 = st.columns(4)
    m1.metric("场地利用率 Y*/K", ratio_disp, help=f"峰值 {peak_time}")
    m2.metric("场地峰值 Y*", f"{peak}" if peak != "—" else "—", "gaylord")
    m3.metric("最大 dock 占用", f"{max_dock} / {docks_cap}", f"满载 {full_h}h")
    m4.metric("甩挂 dock·h 占比", str(drop_pct), help="甩挂 dock·h / 全部 dock·h")

    st.divider()

    # —— 主图：瓶颈对比
    if "4_瓶颈对比" in r["imgs"]:
        st.markdown("##### 瓶颈对比（主图）")
        st.image(r["imgs"]["4_瓶颈对比"], use_column_width=True)

    # —— 2 列：潮汐图 | 月台甘特
    c1, c2 = st.columns(2)
    with c1:
        if "1_潮汐图" in r["imgs"]:
            st.markdown("##### 场地潮汐")
            st.image(r["imgs"]["1_潮汐图"], use_column_width=True)
    with c2:
        if "3_月台甘特" in r["imgs"]:
            st.markdown("##### 月台甘特")
            st.image(r["imgs"]["3_月台甘特"], use_column_width=True)

    # —— 时间窗规则（宽图）
    if "5_时间窗规则" in r["imgs"]:
        st.markdown("##### 按小时时间窗")
        st.image(r["imgs"]["5_时间窗规则"], use_column_width=True)

    st.divider()

    # —— 折叠：时间窗汇总 / KPI 全表 / 规则 / 数据问题
    cols = st.columns(4)
    with cols[0]:
        with st.expander(f"📅 时间窗汇总 ({len(r['wins']) if r['wins'] is not None else 0})"):
            if r["wins"] is not None and len(r["wins"]):
                st.dataframe(r["wins"], use_container_width=True, hide_index=True)
            else:
                st.caption("—")
    with cols[1]:
        with st.expander(f"📊 KPI 全表 ({len(k)})"):
            st.dataframe(k, use_container_width=True, hide_index=True)
    with cols[2]:
        with st.expander(f"📋 小时规则 ({len(r['rules']) if r['rules'] is not None else 0})"):
            if r["rules"] is not None and len(r["rules"]):
                st.dataframe(r["rules"], use_container_width=True, hide_index=True)
            else:
                st.caption("—")
    with cols[3]:
        n_iss = len(r["issues"])
        label = f"⚠ 数据问题 ({n_iss})" if n_iss else "✓ 数据问题 (0)"
        with st.expander(label):
            if n_iss:
                st.dataframe(pd.DataFrame(r["issues"]), use_container_width=True, hide_index=True)
            else:
                st.caption("—")


def main():
    st.set_page_config("库容潮汐 Dashboard", layout="wide")
    init_state()
    sidebar()
    sd = st.session_state.shift_date
    sh = int(st.session_state.shift_start_h)
    t1 = pd.Timestamp(sd) + pd.Timedelta(hours=sh + 24)
    st.markdown(
        f"### 库容潮汐 · {sd:%m-%d}  "
        f"<span style='color:#888;font-size:0.75em;font-weight:normal'>"
        f"{sh:02d}:00 → {t1:%m-%d %H:%M} · K {st.session_state.K:.0f} · 月台 {st.session_state.docks}</span>",
        unsafe_allow_html=True,
    )
    tabs = st.tabs(["📊 分析", "🚚 出港", "🎯 实盘", "📦 W0", "📥 关包"])
    with tabs[0]:
        tab_analysis()
    with tabs[1]:
        tab_outbound()
    with tabs[2]:
        tab_spot()
    with tabs[3]:
        tab_w0()
    with tabs[4]:
        tab_closures()


if __name__ == "__main__":
    main()
