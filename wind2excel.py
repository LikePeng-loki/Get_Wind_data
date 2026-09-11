#!/usr/bin/env python3
"""把 Wind 取数结果（wind-mcp-skill 的 JSON 输出）转成 Excel。

用法一：直接取数并导出
  python wind2excel.py stock_data get_stock_kline '{"windcode":"600519.SH","begin_date":"2026-01-01","end_date":"2026-09-11"}' -o 茅台日K.xlsx

用法二：把已保存的 JSON 文件转成 Excel
  python wind2excel.py --json 结果.json -o 结果.xlsx

依赖：pandas、openpyxl（Anaconda 自带）。
"""
import argparse
import json
import os
import re
import subprocess
import sys
from datetime import datetime

import pandas as pd

DEFAULT_CLI = os.path.expanduser("~/.agents/skills/wind-mcp-skill/scripts/cli.mjs")
# 代码/名称列保留文本，避免丢失前导 0；日期列单独处理（注意"今日开盘价"不是日期列）
KEEP_TEXT = re.compile(r"代码|简称|名称|code|name", re.I)
DATE_COL = re.compile(r"日期|时间|交易日|报告期|日$|date|time", re.I)
DATE_ONLY = re.compile(r"^\d{8}$|^\d{4}-\d{2}-\d{2}$")
DATE_TIME = re.compile(r"^\d{4}-\d{2}-\d{2}[T ]\d{2}:\d{2}(:\d{2})?")


def run_cli(cli, server_type, tool_name, params):
    cli_dir = os.path.dirname(os.path.dirname(cli))  # skill 根目录
    proc = subprocess.run(
        ["node", cli, "call", server_type, tool_name, params],
        cwd=cli_dir, capture_output=True, text=True, encoding="utf-8",
    )
    if not proc.stdout.strip():
        sys.exit(f"CLI 没有输出。\n{proc.stderr}")
    return proc.stdout


def to_date_if_possible(series):
    values = series.dropna().astype(str)
    if values.empty:
        return series
    if values.str.match(DATE_ONLY).all():
        return pd.to_datetime(series.astype("string").str.replace("-", ""), format="%Y%m%d", errors="coerce")
    if values.str.match(DATE_TIME).all():
        # 去掉时区后缀，按北京时间原样保留
        cleaned = series.astype("string").str.slice(0, 19).str.replace("T", " ")
        return pd.to_datetime(cleaned, errors="coerce")
    return series


def to_number_if_possible(df):
    df = df.replace({"INVALID": None})
    for col in df.columns:
        if DATE_COL.search(str(col)) and not re.search(r"代码|code", str(col), re.I):
            df[col] = to_date_if_possible(df[col])
            continue
        if KEEP_TEXT.search(str(col)):
            continue
        converted = pd.to_numeric(df[col], errors="coerce")
        # 只有整列（非空值）都能转成数字时才转换
        if converted.notna().sum() == df[col].notna().sum() and df[col].notna().any():
            df[col] = converted
    return df


def table_from_rows(node):
    cols = node.get("columns") or []
    names = [c.get("name") if isinstance(c, dict) else str(c) for c in cols]
    rows = node["rows"]
    width = max((len(r) for r in rows if isinstance(r, list)), default=len(names))
    if len(names) < width:
        names += [f"列{i + 1}" for i in range(len(names), width)]
    units = node.get("unit") or {}
    if isinstance(units, dict):
        names = [f"{n}({units[n]})" if n in units and units[n] else n for n in names]
    if rows and isinstance(rows[0], dict):
        return pd.DataFrame(rows)
    return pd.DataFrame(rows, columns=names[:width])


def extract_tables(obj, path="data"):
    """在任意 JSON 结构里找出能变成表格的部分。返回 [(表名, DataFrame)]。"""
    found = []
    if isinstance(obj, dict):
        if isinstance(obj.get("rows"), list):
            found.append(("数据", table_from_rows(obj)))
            return found
        # EDB 宏观指标：{meta, date[], value[]}
        if isinstance(obj.get("date"), list) and isinstance(obj.get("value"), list):
            meta = obj.get("meta") or {}
            name = meta.get("name") or path
            unit = "".join(str(meta.get(k) or "") for k in ("magnitude", "unit"))
            label = f"{name}({unit})" if unit else str(name)
            found.append((str(name), pd.DataFrame({"日期": obj["date"], label: obj["value"]})))
            return found
        for k, v in obj.items():
            found += extract_tables(v, f"{path}.{k}" if path else k)
    elif isinstance(obj, list) and obj:
        if all(isinstance(x, dict) for x in obj) and not any(
            isinstance(x.get("rows"), list) or isinstance(x.get("date"), list) for x in obj
        ):
            found.append((path.split(".")[-1] if path != "data" else "数据", pd.json_normalize(obj)))
        else:
            for i, v in enumerate(obj):
                found += extract_tables(v, f"{path}[{i}]")
    return found


def safe_sheet_name(name, used):
    name = re.sub(r"[\[\]:*?/\\]", "_", str(name))[:28] or "Sheet"
    base, n = name, 1
    while name in used:
        n += 1
        name = f"{base[:25]}_{n}"
    used.add(name)
    return name


def main():
    ap = argparse.ArgumentParser(description="Wind 取数结果转 Excel")
    ap.add_argument("server_type", nargs="?")
    ap.add_argument("tool_name", nargs="?")
    ap.add_argument("params", nargs="?", default="{}")
    ap.add_argument("--json", help="已保存的 Wind JSON 输出文件")
    ap.add_argument("-o", "--output", help="输出 Excel 文件名")
    ap.add_argument("--cli", default=DEFAULT_CLI, help="cli.mjs 路径")
    a = ap.parse_args()

    if a.json:
        with open(a.json, encoding="utf-8") as f:
            raw = f.read()
        source = a.json
    elif a.server_type and a.tool_name:
        raw = run_cli(a.cli, a.server_type, a.tool_name, a.params)
        source = f"{a.server_type}.{a.tool_name} {a.params}"
    else:
        ap.print_help()
        sys.exit(1)

    result = json.loads(raw)
    if result.get("ok") is False:
        sys.exit(f"Wind 返回错误：{result.get('code')} - {result.get('message')}")

    tables = []
    for item in result.get("content", []):
        text = item.get("text") if isinstance(item, dict) else None
        if not isinstance(text, str):
            continue
        try:
            payload = json.loads(text)
        except json.JSONDecodeError:
            tables.append(("文本", pd.DataFrame({"内容": [text]})))
            continue
        if isinstance(payload, dict) and payload.get("error"):
            print(f"注意：后端 error 字段 = {payload['error']}", file=sys.stderr)
        tables += extract_tables(payload, "data")

    if not tables:
        sys.exit("没找到表格数据，原始输出：\n" + raw[:2000])

    out = a.output or f"wind_{datetime.now():%Y%m%d_%H%M%S}.xlsx"
    if not out.lower().endswith(".xlsx"):
        out += ".xlsx"

    meta = result.get("cli_meta", {})
    info = pd.DataFrame(
        [
            ("数据来源", "万得 Wind 金融数据服务"),
            ("调用", source),
            ("导出时间", f"{datetime.now():%Y-%m-%d %H:%M:%S}"),
            ("提示", json.dumps(meta.get("warnings", []), ensure_ascii=False)),
        ],
        columns=["项目", "内容"],
    )

    used = set()
    with pd.ExcelWriter(out, engine="openpyxl") as xw:
        for name, df in tables:
            df = to_number_if_possible(df)
            sheet = safe_sheet_name(name, used)
            df.to_excel(xw, sheet_name=sheet, index=False)
            ws = xw.sheets[sheet]
            ws.freeze_panes = "A2"
            for i, col in enumerate(df.columns, start=1):
                if pd.api.types.is_datetime64_any_dtype(df[col]):
                    has_time = (df[col].dropna().dt.floor("D") != df[col].dropna()).any()
                    fmt = "yyyy-mm-dd hh:mm:ss" if has_time else "yyyy-mm-dd"
                    for row in ws.iter_rows(min_row=2, min_col=i, max_col=i):
                        row[0].number_format = fmt
            for i, col in enumerate(df.columns, start=1):
                width = max([len(str(col))] + [len(str(v)) for v in df[col].head(200)])
                ws.column_dimensions[ws.cell(1, i).column_letter].width = min(max(width * 1.6, 8), 50)
            print(f"表 {sheet}: {len(df)} 行 × {len(df.columns)} 列")
        info.to_excel(xw, sheet_name=safe_sheet_name("说明", used), index=False)

    print(f"已保存：{os.path.abspath(out)}")


if __name__ == "__main__":
    main()
