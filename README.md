# 汇金 ETF 公示份额跟踪

> 每天自动抓取上交所/深交所公示的 ETF 总份额，写入追踪表并完成计算与总结。

## 文件说明

| 文件 | 用途 | 维护方式 |
|---|---|---|
| `etf清单.xlsx` | 基金基础数据（名称/代码/来源/网址） | **手动编辑** |
| `etf跟踪.xlsx` | 跟踪表，每基金一个 sheet + 总结 sheet | 脚本自动写入 |
| `etf_fetch.py` | 抓取层：读清单 → 调交易所接口 → 输出 JSON | 改抓取策略时改它 |
| `etf_update.py` | 写入层：读 JSON → 写明细 sheet（含公式）→ 刷新总结 | 改计算/总结时改它 |
| `run_etf.py` | 入口调度：先 fetch 后 update | 定时任务入口 |
| `export_dashboard.py` | 展示层：读 Excel → 生成 `dashboard/report.html`（数据内嵌） | 改页面样式/结构时改它 |
| `dashboard/report.html` | 自包含展示页：汇总表 + 明细筛选分页 | 双击浏览器打开；数据需重新导出 |
| `update_etf.py` | 旧版单文件脚本（已拆分，保留参考） | 不再使用 |
| `data/` | 中间数据目录，存放 `etf_data.json` | 脚本自动生成 |
| `etf跟踪_backup_*.xlsx` | 回填前的自动备份 | 确认无误后可删除 |
| `etf-share-tracker/` | 可分发 Skill 包（SKILL.md + scripts/references/assets） | 供其他 agent 复用 |
| `etf-share-tracker.zip` | Skill 打包文件 | 分发用，可删除 |
| `requirements.txt` | CI 依赖声明（requests / openpyxl） | 新增依赖时更新 |
| `.gitignore` | git 忽略规则（缓存/中间产物/备份） | 改忽略策略时更新 |
| `.github/workflows/update-pages.yml` | GitHub Actions：工作日收盘后自动更新数据并部署 GitHub Pages | 改定时/部署策略时更新 |

## 快速开始

### 日常更新（定时任务执行）

```bash
python run_etf.py
```

周末自动跳过；加 `--force` 可强制运行。

### 历史回填

```bash
# 全部基金回填7月数据
python run_etf.py --backfill 2026-07-01 2026-07-31

# 仅回填单只基金
python run_etf.py --backfill 2026-07-01 2026-07-31 --only 510300
```

### 分层单独运行

```bash
# 只抓数据（不入 Excel）
python etf_fetch.py                           # 最新交易日
python etf_fetch.py --backfill 2026-07-01 2026-07-31 --only 510300  # 历史区间

# 只更新 Excel（读已有 JSON）
python etf_update.py                          # day 模式
python etf_update.py --backfill 2026-07-01 2026-07-31  # 回填模式
```

### 生成展示页面

```bash
# 读 etf跟踪.xlsx 生成 dashboard/report.html（数据内嵌，双击即可打开）
python export_dashboard.py
```

> 展示页特性：上部汇总表（全基金最新状态）、下部明细表（按基金筛选 + 每页 10 条分页）、数值千分位、百分比保留 2 位、涨红跌绿。数据更新后重新运行本脚本即可刷新页面。

## 增删基金

编辑 `etf清单.xlsx` 的 `ETF清单` sheet，按表头添加/删除行即可：

| 名称 | 代码 | 来源 | 网址 |
|---|---|---|---|
| 上证50华夏 | 510050 | 上交所 | https://www.sse.com.cn/market/funddata/volumn/etfvolumn/ |
| ... | ... | ... | ... |

- **来源** 支持 `上交所` 或 `深交所`，其他值会被跳过
- 新增基金后执行一次 `python run_etf.py --backfill 起 止` 回填历史数据

## 数据流

```
etf清单.xlsx ──→ etf_fetch.py ──→ data/etf_data.json ──→ etf_update.py ──→ etf跟踪.xlsx
                   (网络请求)        (中间数据)              (Excel写入+计算)
                                                                   │
                                                                   ▼
                                                   export_dashboard.py → dashboard/report.html
```

## 环境变量

| 变量 | 默认值 | 说明 |
|---|---|---|
| `ETF_LIST_PATH` | `./etf清单.xlsx` | 基金清单路径 |
| `ETF_TRACK_PATH` | `./etf跟踪.xlsx` | 跟踪表路径 |
| `ETF_DATA_PATH` | `./data/etf_data.json` | 中间 JSON 路径 |

## 发布为公开网页（GitHub Pages）

项目已内置 GitHub Actions 工作流 `.github/workflows/update-pages.yml`：
**工作日（周一~周五）北京时间 09:30** 自动执行 `run_etf.py` → `export_dashboard.py`，
把最新数据回写仓库并把 `dashboard/` 部署为公开网页。节假日抓不到新数据时脚本幂等跳过，不会出错。

> 时间说明：上交所/深交所每天早上 7 点才确保更新前一交易日的公示份额，
> 因此定时任务安排在 09:30（7 点后留 2.5 小时缓冲），抓取的是前一交易日的完整数据。
> 选择 09:30 而非更早时间，是为了避开 GitHub 定时任务高峰时段（UTC 0 点前后），
> 降低 schedule 延迟/跳过的概率。
>
> ⚠️ GitHub 的 schedule 定时触发**实测不可靠**（可能延迟数小时甚至不触发）。
> 若某天页面数据没更新，到 Actions 页手动点一次「Run workflow」即可补上。
> 长期可靠的自动化方案见 [`docs/外部定时触发配置指南.md`](docs/外部定时触发配置指南.md)。

### 上线步骤（一次性）

1. 注册 GitHub 账号（https://github.com/signup），准备一个邮箱
2. 新建公开仓库（Public），例如 `Etf-tracker`，**不要**勾选自动生成 README（GitHub Pages 的 Free 计划要求仓库为 Public）
3. 推送项目到仓库：

```bash
git remote add origin https://github.com/<你的用户名>/Etf-tracker.git
git branch -M main
git push -u origin main
```

4. 打开仓库 **Settings → Pages**，在「Build and deployment」的 Source 下拉框选择 **GitHub Actions**
5. 回到仓库 **Actions** 标签页，选中 `ETF 每日更新并发布`，点 **Run workflow** 手动触发一次
6. 首次运行成功后，页面地址为 `https://<你的用户名>.github.io/Etf-tracker/`

> 提示：GitHub Pages 的 Free 计划要求仓库为 **Public**（即公开仓库源码）。若不想公开代码，需升级 Pro/Team，或改用其他静态托管（如 Cloudflare Pages、Netlify）。
> 公开仓库的 Actions 免费且无限量，工作日每日一次约 22 次，绰绰有余。

### 日常维护

- 修改 `etf清单.xlsx` 增删基金后，把文件 push 到 GitHub，下次定时任务即生效（历史数据需 `--backfill` 回填）
- 想立即刷新页面：Actions 页手动 Run workflow
- 网页展示的是 `dashboard/report.html`，即 `export_dashboard.py` 的产物，本地预览仍是原方式

## 数据源

| 交易所 | 页面 | 接口 |
|---|---|---|
| 上交所 | sse.com.cn/market/funddata/volumn/etfvolumn/ | `POST query.sse.com.cn/commonQuery.do` |
| 深交所 | szse.cn/market/fund/volume/etf/index.html | `GET szse.cn/api/report/ShowReport/data` |
