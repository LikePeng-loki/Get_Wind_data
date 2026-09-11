# Wind 取数操作手册（Mac 终端版）

> 适用：红利股跟踪模型「基础数据录入区」取数，以及临时单项取数
> 更新：2026-09-11

---

## 一、文件都在哪

| 路径 | 是什么 | 平时要不要动 |
| --- | --- | --- |
| `~/.agents/skills/wind-mcp-skill/` | Wind 取数技能本体（`scripts/cli.mjs` 负责调用接口，`references/` 是各工具参数说明） | 不动 |
| `~/.agents/skills/wind-find-finance-skill/` | Wind 技能发现/安装路由 | 不动 |
| `~/.claude/skills/` | 指向上面两个技能的快捷方式，给 Claude Code 用 | 不动 |
| `~/.wind-aifinmarket/config` | API Key（一行 `WIND_API_KEY=...`） | 换 Key 时才改 |
| `~/wind-data/` | **工作文件夹**，所有取数都在这里运行 | 常用 |
| `~/wind-data/fetch_dividend_data.py` | 红利股模板一键取数脚本 | 常用 |
| `~/wind-data/wind2excel.py` | 单项取数并导出 Excel 的脚本 | 常用 |
| `~/wind-data/<代码>/` | 每家公司的取数结果，如 `1816HK/` | 查看结果 |

`~` 就是你的用户目录 `/Users/likepeng`。

---

## 二、一次性准备（已完成；换电脑时再做一遍）

| 步骤 | 命令 | 检查是否完成 |
| --- | --- | --- |
| 1. 装 Node.js | 从 https://nodejs.org 下载 LTS 版 .pkg 安装，装完重开终端 | `node -v` 显示版本号 |
| 2. 装 Wind 技能 | `npx --yes --registry=https://registry.npmjs.org skills add https://github.com/Wind-Alice/AliceMarket.git --skill wind-find-finance-skill --skill wind-mcp-skill -g -y` | `ls ~/.agents/skills` 能看到两个文件夹 |
| 3. 给 Claude Code 链接 | `mkdir -p ~/.claude/skills && ln -s ~/.agents/skills/wind-mcp-skill ~/.agents/skills/wind-find-finance-skill ~/.claude/skills/` | `ls ~/.claude/skills` |
| 4. 配 API Key | `mkdir -p ~/.wind-aifinmarket && nano ~/.wind-aifinmarket/config`，写入 `WIND_API_KEY=你的Key`，`Ctrl+O` 回车保存，`Ctrl+X` 退出；再 `chmod 600 ~/.wind-aifinmarket/config` | `grep -c '^WIND_API_KEY=' ~/.wind-aifinmarket/config` 显示 `1` |
| 5. 装 Python 库 | `python -m pip install pandas openpyxl -i https://pypi.tuna.tsinghua.edu.cn/simple` | `python -c "import pandas, openpyxl; print('ok')"` 显示 `ok` |
| 6. 放脚本 | 把 `fetch_dividend_data.py`、`wind2excel.py` 放进 `~/wind-data/` | `ls ~/wind-data` |

安装前 GitHub 连不上时，第 2 步可换国内源：把地址换成 `https://gitee.com/WindAlice/AliceMarket.git`，`--registry` 换成 `https://registry.npmmirror.com`。

---

## 三、每次取数：从打开终端开始

### 第 1 步：打开终端

按 `Command + 空格`，输入“终端”（或 Terminal），回车。

### 第 2 步：进入工作文件夹

```bash
cd ~/wind-data
```

### 第 3 步：运行一键取数

以中广核电力为例：

```bash
python fetch_dividend_data.py --code 1816.HK --name 中广核电力 --latest 2026-06-30 --a-code 003816.SZ
```

屏幕会依次显示 `[1/7] 01_ROE扣非_毛利率 … 成功` 等进度。每一项都会调用一次 Wind、消耗积分；第 1 项失败时脚本会自动停止。

### 第 4 步：看结果

```bash
open 1816HK/1816HK_原始数据.xlsx
```

结果文件夹里：

- `01_…json` 到 `06_…json`：每项查询的 Wind 原始返回（留档、核对用）
- `1816HK_原始数据.xlsx`：所有表格汇总，每项一个工作表，最后一张“说明”表记录每项成功/失败及查询原文

### 第 5 步（可选）：让 Claude 整理进模板

在 Claude 里说：“读取 ~/wind-data/1816HK 的结果，按红利股模板 A 区、B 区整理，TTM、有息负债、近 5 年平均、近 10 年波动率帮我算好。”（新会话里需要先允许 Claude 访问 `wind-data` 文件夹。）

---

## 四、参数怎么填

| 参数 | 必填 | 怎么填 | 例子 |
| --- | --- | --- | --- |
| `--code` | 是 | Wind 代码。A 股：6 位 + `.SH`/`.SZ`/`.BJ`；港股：4 位（不足补 0）+ `.HK` | `600900.SH`、`1816.HK`、`0002.HK` |
| `--name` | 是 | 公司简称（只用于生成查询语句，不影响代码识别） | `长江电力` |
| `--latest` | 是 | 最新**已披露**的报告期，见下表 | `2026-06-30` |
| `--a-code` | 否 | 仅 A+H 两地上市的港股填，取 A 股口径扣非 ROE 作候选 | `003816.SZ` |
| `--only` | 否 | 只补取某几项，填编号，逗号分隔 | `02,04b` |

### `--latest` 怎么选

按“今天之前已经过了哪个披露截止日”选最近的一期：

| 报告期 | 填 | A 股披露截止 | 港股披露截止 |
| --- | --- | --- | --- |
| 一季报 | `YYYY-03-31` | 4 月 30 日 | 多数港股无季报 |
| 中报 | `YYYY-06-30` | 8 月 31 日 | 8 月 31 日 |
| 三季报 | `YYYY-09-30` | 10 月 31 日 | 多数港股无季报 |
| 年报 | `YYYY-12-31` | 次年 4 月 30 日 | 次年 3 月 31 日 |

- 填中报/季报：脚本自动取“本期累计、上年年报、上年同期”三期，用于算 TTM（TTM = 本期累计 + 上年年报 − 上年同期）。
- 填年报（12-31）：TTM 就是年报本身，只取一期。
- A 区的“近 10 年”自动以最近一个完整年报年度为终点倒推（如 `--latest 2026-06-30` → 2016–2025）。

### 常用例子

```bash
# 港股，同时有 A 股（A+H）
python fetch_dividend_data.py --code 1816.HK --name 中广核电力 --latest 2026-06-30 --a-code 003816.SZ
python fetch_dividend_data.py --code 1088.HK --name 中国神华 --latest 2026-06-30 --a-code 601088.SH

# 纯 A 股（不需要 --a-code）
python fetch_dividend_data.py --code 600900.SH --name 长江电力 --latest 2026-06-30

# 只补取第 02 项和 04b 项（其他项不重新调用，汇总 Excel 会自动合并所有已有结果）
python fetch_dividend_data.py --code 1816.HK --name 中广核电力 --latest 2026-06-30 --only 02,04b
```

---

## 五、每项查的是什么

| 编号 | 对应模板 | 查询内容 |
| --- | --- | --- |
| 01 | A 区 ROE（扣非）、毛利率 | 近 10 年每年年报的扣非 ROE、销售毛利率 |
| 02 | A 区 股息率(TTM) | 近 5 年每年 12-31（非交易日取前一交易日）的股息率（近 12 个月） |
| 03 | A 区 分红率 | 近 5 年现金分红总额、归母净利润、现金分红比例 |
| 04 | B 区 营收/净利润/经营现金流/财务费用 | 算 TTM 所需的各期累计值 |
| 04b | B 区 财务费用（仅港股自动加） | 财务成本（利息支出），港股报表常无“财务费用”科目 |
| 05 | B 区 资产负债表项目 | 总资产、总负债、有息负债各分项、预收款项、合同负债 |
| 06 | A 区 ROE（扣非）候选（仅填了 `--a-code` 时） | A 股口径扣非 ROE |

有息负债 = 短期借款 + 一年内到期的非流动负债 + 长期借款 + 应付债券 + 租赁负债（缺项按 0）。

---

## 六、已知口径问题（1816.HK 实跑发现，取其他公司时也要核对）

1. **现金分红比例 Wind 可能返回空值**：用 `现金分红总额 ÷ 归母净利润` 自己算（03 项两个数都有）。
2. **股息率年份要核对**：旧版脚本问“每年最后一个交易日”时，Wind 返回的日期整体错了一年（2020–2024）。现版本已改为明确日期，但取完仍要看表里的“时间”列是否是对应年份年末。1816HK 现有的 02 结果是旧版的，需要用 `--only 02` 补取。
3. **港股“财务费用”返回空值**：港股报表科目不同，看 04b 的“财务成本”；口径与 A 股“财务费用”（含利息收入抵减）不完全相同，使用时注明。
4. **扣非 ROE 有两个口径**：01 项（1816.HK）与 06 项（003816.SZ，A 股“扣除/加权”）数值略有差异（如 2016 年 12.89% vs 12.57%）。模板统一用哪个要先定，全表保持一致。
5. **单位不统一**：同一张表里有的列是“亿元”，有的是“元”（如财务费用、预收款项），以 Excel 列名括号里的单位为准，填模板前统一换算。
6. **币种**：港股公司财务数据一般为人民币（看“记账本位币”列），股价为港元；股息率用 Wind 现成值，不要自己用港元股价 ÷ 人民币股利。
7. **空值 ≠ 0**：Wind 返回空值表示缺失或不适用，除有息负债分项外不要按 0 计算。

---

## 七、单项临时取数（wind2excel.py）

格式：`python wind2excel.py <数据类别> <工具名> '<参数JSON>' -o 文件名.xlsx`

```bash
cd ~/wind-data

# 日 K 线（默认前复权；不复权加 "aftype":"2"；周K "period":"1w"，月K "1mo"）
python wind2excel.py stock_data get_stock_kline '{"windcode":"1816.HK","begin_date":"2026-01-01","end_date":"2026-09-11"}' -o 中广核电力日K.xlsx

# 最新行情快照（可多只，逗号分隔，单次最多 50 只）
python wind2excel.py stock_data get_stock_price_indicators '{"windcode":"1816.HK,600900.SH"}' -o 行情快照.xlsx

# 财务/估值（自然语言提问）
python wind2excel.py stock_data get_stock_fundamentals '{"question":"查询长江电力（600900.SH）2025-12-31的ROE、营业收入和净利润"}' -o 长江电力财务.xlsx

# 分红派息历史
python wind2excel.py stock_data get_stock_events '{"question":"查询中广核电力（1816.HK）的分红派息历史"}' -o 中广核分红.xlsx

# 打开导出的文件
open 中广核电力日K.xlsx
```

其他数据类别：`fund_data`（基金）、`index_data`（指数）、`bond_data`（债券）、`economic_data`（宏观）、`financial_docs`（公告/新闻）。各工具参数写法见 `~/.agents/skills/wind-mcp-skill/references/` 里对应的 `.md` 文件，不确定时让 Claude 写好命令再复制运行。

---

## 八、常见报错

| 现象 | 原因 | 处理 |
| --- | --- | --- |
| `command not found: npx` / `node` | 没装 Node.js | 见第二部分第 1 步 |
| `No module named 'pandas'` | 没装 Python 库 | `python -m pip install pandas openpyxl` |
| `No such file or directory` | 路径不对或不在 `~/wind-data` | 先 `cd ~/wind-data`，再 `ls` 看文件在不在 |
| 命令长时间没反应 | 网络连不上 Wind（VPN/代理/公司内网） | 等 1–2 分钟无果按 `Ctrl+C`；测网络：`curl -sS -m 10 -o /dev/null -w "%{http_code}\n" https://mcp.wind.com.cn/vserver_stock_data/mcp/`，返回三位数字（如 401/405）说明网络通，`000` 说明不通 |
| `AUTH_ERROR` | Key 没配或写错 | 检查 `~/.wind-aifinmarket/config` |
| `RATE_LIMIT_ERROR` / 额度不足 | 调用太频繁或积分用完 | 稍后重试 / 到开发者中心查额度 |
| `PARAM_VALIDATION_ERROR` | 参数名或格式不对 | 对照 `references/*.md` 修改，或发给 Claude |
| nano 退出时问 `Save modified buffer?` | 还没保存 | 按 `Y`，再按回车 |
| 终端里上一条命令还在跑时又输入了新命令 | 新命令会排队等前一条结束 | 等出现 `%` 提示符再输入下一条 |

---

## 附：注意事项

- 每次调用都消耗 Wind 积分；只缺某几项时用 `--only` 补取，不要整套重跑。
- `wind-mcp-skill` 每天会在后台自动检查并更新自身（从 GitHub/Gitee 重装），代码可能在不知情时变化。
- 数据来源于万得 Wind 金融数据服务。
