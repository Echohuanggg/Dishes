# -*- coding: utf-8 -*-
"""
菜品系统本地服务：HTTP 中转，实现「HTML 加菜/记录 -> 自动写入 Excel」全自动流程。
- 零依赖（仅 Python 标准库 http.server + openpyxl）
- 端口默认 8765，被占用时自动递增找空闲端口
- GET  /              -> 返回同目录 菜品选择.html
- POST /api/add_dish  -> 新增菜品写入 菜品数据.xlsx（按名称去重、编号顺延、重建统计报表）
- POST /api/sync_records -> 选菜记录合并去重写入 Excel
"""
import json
import os
import socket
import sys
import webbrowser
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from urllib.parse import urlparse

try:
    import openpyxl
except ImportError:
    openpyxl = None

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
HTML_FILE = os.path.join(BASE_DIR, "菜品选择.html")
EXCEL_FILE = os.path.join(BASE_DIR, "菜品数据.xlsx")

DATA_HEADERS = ["编号", "菜品名称", "餐次", "分类", "食材（原料）", "菜谱步骤", "抖音链接", "食材清单"]
RECORD_HEADERS = ["日期", "餐次", "菜品编号", "菜品名称"]
STAT_MEAL_TITLE = "各餐次菜品数量"
STAT_CAT_TITLE = "各分类菜品数量"
STAT_PICK_TITLE = "菜品被选次数"


def ensure_openpyxl():
    if openpyxl is None:
        raise RuntimeError("缺少 openpyxl，请先执行: pip install openpyxl")


def find_free_port(start):
    port = start
    while port < start + 200:
        with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
            try:
                s.bind(("127.0.0.1", port))
                return port
            except OSError:
                port += 1
    raise RuntimeError("未找到可用端口")


def load_excel():
    ensure_openpyxl()
    wb = openpyxl.load_workbook(EXCEL_FILE)
    return wb


def get_max_dish_id(ws):
    max_id = 0
    for r in range(2, ws.max_row + 1):
        v = ws.cell(r, 1).value
        try:
            max_id = max(max_id, int(v))
        except (TypeError, ValueError):
            continue
    return max_id


def rebuild_stats_sheet(wb, dish_rows):
    """重建统计报表：动态餐次/分类计数 + 菜品被选次数 COUNTIF 公式"""
    if "统计报表" in wb.sheetnames:
        del wb["统计报表"]
    ws = wb.create_sheet("统计报表")

    ws.cell(1, 1, "各餐次菜品数量")
    ws.cell(1, 4, "各分类菜品数量")
    ws.cell(1, 7, "菜品被选次数")
    ws.cell(2, 1, "餐次")
    ws.cell(2, 2, "菜品数量")
    ws.cell(2, 4, "分类")
    ws.cell(2, 5, "菜品数量")
    ws.cell(2, 7, "菜品名称")
    ws.cell(2, 8, "被选次数")

    # 动态统计餐次（固定顺序：早餐/正餐/甜品 + 自定义按出现顺序）
    meals = []
    cats = []
    for row in dish_rows:
        meal = str(row.get("餐次") or "").strip()
        cat = str(row.get("分类") or "").strip()
        if meal and meal not in meals:
            meals.append(meal)
        if cat and cat not in cats:
            cats.append(cat)
    meal_prio = {"早餐": 0, "正餐": 1, "甜品": 2}
    meals = sorted(meals, key=lambda m: meal_prio.get(m, 99))
    cat_prio = {"素菜": 0, "主食": 1, "荤菜": 2, "甜品": 3, "汤": 4, "酱料": 5}
    cats = sorted(cats, key=lambda c: cat_prio.get(c, 99))

    r = 3
    for meal in meals:
        cnt = sum(1 for row in dish_rows if str(row.get("餐次") or "").strip() == meal)
        ws.cell(r, 1, meal)
        ws.cell(r, 2, cnt)
        r += 1
    ws.cell(r, 1, "合计")
    ws.cell(r, 2, len(dish_rows))

    r = 3
    for cat in cats:
        cnt = sum(1 for row in dish_rows if str(row.get("分类") or "").strip() == cat)
        ws.cell(r, 4, cat)
        ws.cell(r, 5, cnt)
        r += 1
    ws.cell(r, 4, "合计")
    ws.cell(r, 5, len(dish_rows))

    # 菜品被选次数（COUNTIF 引用选菜记录!$D:$D）
    r = 3
    for row in dish_rows:
        name = row.get("菜品名称")
        if name is None:
            continue
        ws.cell(r, 7, name)
        ws.cell(r, 8, "=COUNTIF(选菜记录!$D:$D,G{0})".format(r))
        r += 1

    # 简单样式：标题加粗
    for c in range(1, 9):
        ws.cell(1, c).font = openpyxl.styles.Font(bold=True)
    for c in (1, 2, 4, 5, 7, 8):
        ws.cell(2, c).font = openpyxl.styles.Font(bold=True)
    return ws


def api_add_dish(body):
    ensure_openpyxl()
    name = str(body.get("菜品名称") or "").strip()
    if not name:
        return {"ok": False, "提示": "菜品名称不能为空"}
    # 兼容前端两种字段名：食材 / 食材（原料）
    ingr = str(body.get("食材") or body.get("食材（原料）") or "").strip()
    meal = str(body.get("餐次") or "").strip()
    cat = str(body.get("分类") or "").strip()
    steps = str(body.get("菜谱步骤") or "").strip()
    douyin = str(body.get("抖音链接") or "").strip()
    shop = str(body.get("食材清单") or "").strip()

    wb = load_excel()
    ws = wb["菜品数据"]
    # 按菜品名称去重
    for r in range(2, ws.max_row + 1):
        if str(ws.cell(r, 2).value or "").strip() == name:
            return {"ok": False, "提示": "已存在同名菜品《" + name + "》，未重复添加"}

    new_id = get_max_dish_id(ws) + 1
    row = [new_id, name, meal, cat, ingr, steps, douyin, shop]
    ws.append(row)
    # 表头样式保持一致（若表头无加粗则补上）
    for c in range(1, len(DATA_HEADERS) + 1):
        cell = ws.cell(1, c)
        if cell.value is None:
            cell.value = DATA_HEADERS[c - 1]
        cell.font = openpyxl.styles.Font(bold=True)

    # 重建统计报表
    dish_rows = []
    for r in range(2, ws.max_row + 1):
        vals = [ws.cell(r, c).value for c in range(1, 9)]
        if vals[0] is None and vals[1] is None:
            continue
        dish_rows.append({
            "编号": vals[0], "菜品名称": vals[1], "餐次": vals[2], "分类": vals[3],
            "食材（原料）": vals[4], "菜谱步骤": vals[5], "抖音链接": vals[6], "食材清单": vals[7],
        })
    rebuild_stats_sheet(wb, dish_rows)

    wb.save(EXCEL_FILE)
    # 保存后重开验证
    wb2 = openpyxl.load_workbook(EXCEL_FILE)
    ws2 = wb2["菜品数据"]
    found = None
    for r in range(2, ws2.max_row + 1):
        if str(ws2.cell(r, 2).value or "").strip() == name:
            found = (ws2.cell(r, 1).value, r)
            break
    if found is None:
        return {"ok": False, "提示": "写入验证失败：未在 Excel 中找到新菜品"}
    return {"ok": True, "编号": found[0], "提示": "已写入Excel（编号" + str(found[0]) + "）"}


def api_sync_records(body):
    ensure_openpyxl()
    if not isinstance(body, list):
        return {"ok": False, "提示": "请求体应为记录数组"}

    wb = load_excel()
    ws = wb["选菜记录"]
    existing = set()
    for r in range(2, ws.max_row + 1):
        d = ws.cell(r, 1).value
        m = ws.cell(r, 2).value
        i = ws.cell(r, 3).value
        n = ws.cell(r, 4).value
        if d is None and m is None and i is None and n is None:
            continue
        existing.add((str(d or ""), str(m or ""), str(i or ""), str(n or "")))

    new_count = 0
    pending = []
    for rec in body:
        d = str(rec.get("日期") or "").strip()
        m = str(rec.get("餐次") or "").strip()
        i = str(rec.get("菜品编号") or rec.get("id") or "").strip()
        n = str(rec.get("菜品名称") or rec.get("name") or "").strip()
        if not (d or m or i or n):
            continue
        key = (d, m, i, n)
        if key in existing:
            continue
        existing.add(key)
        pending.append([d, m, i, n])

    if pending:
        # 表头保留
        for c, h in enumerate(RECORD_HEADERS, start=1):
            cell = ws.cell(1, c)
            if cell.value is None:
                cell.value = h
            cell.font = openpyxl.styles.Font(bold=True)
        start_row = ws.max_row + 1
        for row in pending:
            for c, v in enumerate(row, start=1):
                ws.cell(start_row, c, v)
            start_row += 1
        wb.save(EXCEL_FILE)
    total = len(existing)
    return {"ok": True, "新增条数": len(pending), "总条数": total}


class Handler(BaseHTTPRequestHandler):
    server_version = "CaiPinService/1.0"

    def _send_cors(self):
        self.send_header("Access-Control-Allow-Origin", "*")
        self.send_header("Access-Control-Allow-Methods", "GET, POST, OPTIONS")
        self.send_header("Access-Control-Allow-Headers", "Content-Type")
        self.send_header("Access-Control-Max-Age", "3600")

    def _send_json(self, obj, status=200):
        data = json.dumps(obj, ensure_ascii=False).encode("utf-8")
        self.send_response(status)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Content-Length", str(len(data)))
        self._send_cors()
        self.end_headers()
        self.wfile.write(data)

    def _send_html(self, status=200):
        try:
            with open(HTML_FILE, "rb") as f:
                data = f.read()
        except OSError:
            self._send_json({"ok": False, "提示": "未找到 菜品选择.html"}, status=500)
            return
        self.send_response(status)
        self.send_header("Content-Type", "text/html; charset=utf-8")
        self.send_header("Content-Length", str(len(data)))
        self._send_cors()
        self.end_headers()
        self.wfile.write(data)

    def log_message(self, fmt, *args):
        sys.stdout.write("[%s] %s\n" % (self.log_date_time_string(), fmt % args))
        sys.stdout.flush()

    def do_OPTIONS(self):
        self.send_response(204)
        self._send_cors()
        self.end_headers()

    def do_GET(self):
        parsed = urlparse(self.path)
        if parsed.path == "/":
            self._send_html()
        elif parsed.path == "/api/ping":
            self._send_json({"ok": True, "service": "cai-pin"})
        else:
            self._send_json({"ok": False, "提示": "Not Found"}, status=404)

    def do_POST(self):
        parsed = urlparse(self.path)
        length = int(self.headers.get("Content-Length") or 0)
        raw = self.rfile.read(length) if length else b""
        try:
            body = json.loads(raw.decode("utf-8")) if raw else None
        except (ValueError, UnicodeDecodeError):
            self._send_json({"ok": False, "提示": "JSON 解析失败"}, status=400)
            return

        try:
            if parsed.path == "/api/add_dish":
                if not isinstance(body, dict):
                    self._send_json({"ok": False, "提示": "请求体应为 JSON 对象"}, status=400)
                    return
                result = api_add_dish(body)
                self._send_json(result)
            elif parsed.path == "/api/sync_records":
                result = api_sync_records(body)
                self._send_json(result)
            else:
                self._send_json({"ok": False, "提示": "Not Found"}, status=404)
        except Exception as e:  # noqa: BLE001
            self._send_json({"ok": False, "提示": "服务端错误: " + str(e)}, status=500)


def main():
    import argparse
    ap = argparse.ArgumentParser(description="菜品系统本地服务")
    ap.add_argument("--no-browser", action="store_true", help="启动后不自动打开浏览器（由 bat 统一打开页面）")
    args = ap.parse_args()

    if openpyxl is None:
        print("[错误] 未检测到 openpyxl，请先执行: pip install openpyxl")
        input("按回车键退出...")
        return 1
    port = find_free_port(8765)
    url = "http://127.0.0.1:%d/" % port
    server = ThreadingHTTPServer(("127.0.0.1", port), Handler)
    print("=" * 56)
    print("  慧的菜品选择系统 - 本地服务已启动")
    print("  访问地址: %s" % url)
    print("  接口说明:")
    print("    POST /api/add_dish      新增菜品并写入 Excel")
    print("    POST /api/sync_records  选菜记录写入 Excel")
    print("  停止服务: 关闭本窗口或按 Ctrl+C")
    print("=" * 56)
    if not args.no_browser:
        webbrowser.open(url)
    try:
        server.serve_forever()
    except KeyboardInterrupt:
        print("\n服务已停止")
    return 0


if __name__ == "__main__":
    sys.exit(main())
