#!/usr/bin/env python3
"""按「红利股跟踪模型 · 基础数据录入区」取数（A 历史年度序列 + B 最新报告期财务数据）。

用法（在 ~/wind-data 里运行）：
  python fetch_dividend_data.py --code 1816.HK --name 中广核电力 --latest 2026-06-30 --a-code 003816.SZ
  python fetch_dividend_data.py --code 600900.SH --name 长江电力 --latest 2026-06-30

结果：
  <代码>/  下每个问题一个原始 JSON（01_*.json …），以及汇总 Excel <代码>_原始数据.xlsx
"""
import argparse
import glob
import json
import os
import subprocess
import sys
from datetime import datetime

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, SCRIPT_DIR)
import pandas as pd  # noqa: E402
from wind2excel import extract_tables, safe_sheet_name, to_number_if_possible  # noqa: E402

DEFAULT_CLI = os.path.expanduser("~/.agents/skills/wind-mcp-skill/scripts/cli.mjs")


def build_questions(code, name, latest, a_code):
    y = int(latest[:4])
    is_annual = latest[5:] == "12-31"
    last_annual = latest if is_annual else f"{y - 1}-12-31"
    first_year, last_year = int(last_annual[:4]) - 9, int(last_annual[:4])
    div_from = last_year - 4
    year_ends = "、".join(f"{yr}-12-31" for yr in range(div_from, last_year + 1))
    # TTM 组件：年报期只要当期；非年报期要 本期累计、上年年报、上年同期
    periods = [latest] if is_annual else [latest, last_annual, f"{y - 1}{latest[4:]}"]
    periods_txt = "、".join(periods)
    ent = f"{name}（{code}）"
    qs = [
        ("01_ROE扣非_毛利率",
         f"查询{ent}{first_year}年至{last_year}年每年年报（12月31日）的净资产收益率ROE（扣除非经常性损益）和销售毛利率"),
        ("02_股息率TTM",
         f"查询{ent}在{year_ends}这几个日期的股息率（近12个月），非交易日取前一交易日"),
        ("03_分红率",
         f"查询{ent}{div_from}年至{last_year}年每个会计年度的现金分红总额、归属母公司股东的净利润和现金分红比例"),
        ("04_TTM组件_利润表现金流",
         f"查询{ent}{periods_txt}报告期的营业收入、净利润、经营活动产生的现金流量净额、财务费用（合并报表，累计值）"),
        ("05_资产负债表",
         f"查询{ent}{latest}报告期末的总资产、总负债、短期借款、一年内到期的非流动负债、长期借款、应付债券、租赁负债、预收款项、合同负债（合并报表）"),
    ]
    if code.upper().endswith(".HK"):
        # 港股报表常无"财务费用"科目，补取"财务成本"作候选
        qs.append(("04b_候选_港股财务成本",
                   f"查询{ent}{periods_txt}报告期的财务成本（利息支出）（合并报表，累计值）"))
    if a_code:
        qs.append(("06_候选_A股ROE扣非",
                   f"查询{a_code}{first_year}年至{last_year}年每年年报的净资产收益率ROE（扣除非经常性损益）"))
    return qs


def call(cli, question):
    params = json.dumps({"question": question}, ensure_ascii=False)
    proc = subprocess.run(
        ["node", cli, "call", "stock_data", "get_stock_fundamentals", params],
        cwd=os.path.dirname(os.path.dirname(cli)),
        capture_output=True, text=True, encoding="utf-8",
    )
    out = proc.stdout.strip()
    if not out:
        return {"ok": False, "code": "NO_OUTPUT", "message": proc.stderr.strip()[:500]}
    try:
        return json.loads(out)
    except json.JSONDecodeError:
        return {"ok": False, "code": "BAD_JSON", "message": out[:500]}


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--code", required=True, help="Wind代码，如 1816.HK / 600900.SH")
    ap.add_argument("--name", required=True, help="公司简称，如 中广核电力")
    ap.add_argument("--latest", required=True, help="最新报告期，如 2026-06-30")
    ap.add_argument("--a-code", default="", help="可选：港股公司对应的A股代码，用于取A股口径扣非ROE作候选")
    ap.add_argument("--only", default="", help="只补取某几项，填编号前缀，逗号分隔，如 02,04b")
    ap.add_argument("--cli", default=DEFAULT_CLI)
    a = ap.parse_args()

    out_dir = os.path.join(SCRIPT_DIR, a.code.replace(".", ""))
    os.makedirs(out_dir, exist_ok=True)
    questions = build_questions(a.code, a.name, a.latest, a.a_code)

    if a.only:
        wanted = [x.strip() for x in a.only.split(",") if x.strip()]
        questions = [(k, q) for k, q in questions if any(k.startswith(w + "_") for w in wanted)]
        if not questions:
            sys.exit(f"--only {a.only} 没有匹配的项目")

    for i, (key, q) in enumerate(questions, start=1):
        print(f"[{i}/{len(questions)}] {key} …", flush=True)
        result = call(a.cli, q)
        with open(os.path.join(out_dir, f"{key}.json"), "w", encoding="utf-8") as f:
            json.dump({"question": q, "result": result}, f, ensure_ascii=False, indent=2)
        if result.get("ok") is False:
            print(f"    失败：{result.get('code')} {result.get('message')}")
            if i == 1 and len(questions) > 1:  # 第一条当探针，失败就停，避免浪费积分
                sys.exit("探针调用失败，已停止。请把上面的报错发给 Claude。")
        else:
            print("    成功")

    # 汇总表：用该公司文件夹里全部 JSON 重建（补取后也是完整的）
    xlsx_tables, log = [], []
    for path in sorted(glob.glob(os.path.join(out_dir, "*.json"))):
        key = os.path.splitext(os.path.basename(path))[0]
        with open(path, encoding="utf-8") as f:
            saved = json.load(f)
        result = saved.get("result", {})
        if result.get("ok") is False:
            log.append((key, f"失败 {result.get('code')}: {result.get('message')}", saved.get("question")))
            continue
        n = 0
        for item in result.get("content", []):
            text = item.get("text") if isinstance(item, dict) else None
            if not isinstance(text, str):
                continue
            try:
                payload = json.loads(text)
            except json.JSONDecodeError:
                xlsx_tables.append((f"{key}_文本", pd.DataFrame({"内容": [text]})))
                n += 1
                continue
            for _, df in extract_tables(payload, "data"):
                xlsx_tables.append((key, df))
                n += 1
        log.append((key, f"成功，{n} 张表", saved.get("question")))

    xlsx = os.path.join(out_dir, f"{a.code.replace('.', '')}_原始数据.xlsx")
    used = set()
    with pd.ExcelWriter(xlsx, engine="openpyxl") as xw:
        for name, df in xlsx_tables:
            to_number_if_possible(df).to_excel(xw, sheet_name=safe_sheet_name(name, used), index=False)
        pd.DataFrame(
            log + [("数据来源", "万得 Wind 金融数据服务", f"导出于 {datetime.now():%Y-%m-%d %H:%M}")],
            columns=["步骤", "状态", "问题"],
        ).to_excel(xw, sheet_name=safe_sheet_name("说明", used), index=False)

    print(f"\n完成：{out_dir}")
    print(f"汇总表：{xlsx}")


if __name__ == "__main__":
    main()
