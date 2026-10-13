# LAX CBT 库容潮汐板

单文件 html 看板，Ray 后台更新数据 → GitHub Pages 托管 → TT 现场查看。

## 两个 URL

| 角色 | 链接 | 说明 |
|---|---|---|
| **Ray 后台** | `.../dock_board.html?admin=1` | 关包可编辑；点「🚀 发布」把数据 PUT 到本 repo 的 `data.json` |
| **TT 查看** | `.../dock_board.html` | 关包只读；启动自动拉 `data.json`；点「🔄 刷新」拉最新；Dock 可编辑并本地持久化 |

## Ray 第一次用（配置 PAT）

1. 打开 GitHub → Settings → Developer settings → **Fine-grained personal access tokens** → New
2. **Resource owner**：你自己；**Repository access**：Only select repositories → 选这个 repo
3. **Permissions / Repository permissions**：**Contents = Read and write**
4. Generate 后复制 token
5. 打开 `.../dock_board.html?admin=1`，点「🔑 Token」→ 粘贴 token
6. Token 存在你浏览器 localStorage，不会上传任何地方（但别在别人电脑上粘贴）

## 日常流程

- Ray 打开后台 URL → 填关包清单 + 时刻 + K → 点「🚀 发布」
- TT 打开 TT URL → 看到最新数据 + 自己改 Dock（本地）
- Dock 状态 **不同步**（TT 自己维护，Ray 看不到；他们截图发回即可）

## 本地开发

```bash
python3 -m http.server 8512
# http://127.0.0.1:8512/dock_board.html?admin=1
```

## 配套文件

- `tide_chart.py` — Ray 自己分析用的脚本（CLI 出 KPI/图）
- `app.py` — Streamlit 一体化 dashboard（Ray 自己分析用）
- `潮汐图_模型说明.md` — 完整模型文档
- `潮汐图_输入模板.xlsx` — tide_chart.py 的输入模板
- `dock_board.html` — **交付品**，发给 TT 的看板
- `data.json` — Ray 发布的最新数据
