# -*- coding: utf-8 -*-
"""
update_html.py - 将 菜品数据.xlsx 的最新菜品数据同步到 index.html 的内嵌 DISHES 数组。
用法: python update_html.py [--no-pause]
- 自动比对 Excel 与 HTML 内嵌数据，有变化才写回
- Excel 被占用（Excel 打开中）会给出明确提示
"""
import argparse
import json
import os
import re
import sys

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
EXCEL_FILE = os.path.join(BASE_DIR, "菜品数据.xlsx")
HTML_FILE = os.path.join(BASE_DIR, "index.html")
FIELDS = ["编号", "菜品名称", "餐次", "分类", "菜谱步骤", "抖音链接", "食材清单", "食材（原料）"]


def load_dishes_from_excel():
    try:
        import openpyxl
    except ImportError:
        print("[错误] 缺少 openpyxl，请先执行: pip install openpyxl")
        return None
    try:
        wb = openpyxl.load_workbook(EXCEL_FILE, data_only=True)
    except PermissionError:
        print("[错误] 菜品数据.xlsx 被占用（可能正用 Excel 打开着），请先关闭 Excel 再运行。")
        return None
    ws = wb["菜品数据"]
    rows = list(ws.iter_rows(values_only=True))
    header = [str(h or "").strip() for h in rows[0]]
    dishes = []
    for r in rows[1:]:
        if r[0] is None and r[1] is None:
            continue
        d = {}
        for i, col in enumerate(header):
            v = r[i] if i < len(r) else None
            d[col] = "" if v is None else str(v).strip()
        dishes.append(d)
    return dishes


def build_js_array(dishes):
    parts = []
    for d in dishes:
        fields = []
        for f in FIELDS:
            fields.append('"%s": %s' % (f, json.dumps(d.get(f, ""), ensure_ascii=False)))
        parts.append("{" + ", ".join(fields) + "}")
    return "const DISHES = [" + ",\n".join(parts) + "];"


def main():
    ap = argparse.ArgumentParser(description="同步 Excel 菜品数据到 HTML")
    ap.add_argument("--no-pause", action="store_true", help="运行结束后不等待回车")
    args = ap.parse_args()

    dishes = load_dishes_from_excel()
    if dishes is None:
        if not args.no_pause:
            input("按回车键退出...")
        return 1

    try:
        with open(HTML_FILE, encoding="utf-8") as f:
            content = f.read()
    except OSError:
        print("[错误] 未找到 index.html")
        if not args.no_pause:
            input("按回车键退出...")
        return 1

    new_js = build_js_array(dishes)
    pattern = re.compile(r"const DISHES = \[.*?\];", re.S)
    m = pattern.search(content)
    if not m:
        print("[错误] 未在 HTML 中找到 DISHES 数组，请确认 index.html 未损坏。")
        if not args.no_pause:
            input("按回车键退出...")
        return 1

    if m.group(0) == new_js:
        print("[信息] 菜品数据无变化（共 %d 道菜），HTML 已是最新。" % len(dishes))
    else:
        content = pattern.sub(lambda _: new_js, content, count=1)
        with open(HTML_FILE, "w", encoding="utf-8") as f:
            f.write(content)
        print("[完成] 已更新 HTML 内嵌数据：共 %d 道菜。" % len(dishes))

    if not args.no_pause:
        input("按回车键退出...")
    return 0


if __name__ == "__main__":
    sys.exit(main())
