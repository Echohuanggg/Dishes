# -*- coding: utf-8 -*-
"""
选菜记录 + 新菜品自动同步脚本
功能：
  1. 自动扫描 选菜记录_*.json 文件，合并去重后写入 菜品数据.xlsx 的「选菜记录」工作表
  2. 自动扫描 新菜品_*.json 文件，按菜名去重合并后写入「菜品数据」工作表，
     新菜品编号自动分配（现有最大编号+1 依次递增），并重建「统计报表」工作表
用法：双击「同步选菜记录.bat」或命令行执行 python sync_records.py
扫描位置：本脚本所在目录（慧做菜文件夹）+ 系统下载目录（Downloads）
"""

import json
import os
import sys
import glob
import shutil
import subprocess

# ---------- 依赖检查：缺少 openpyxl 自动安装 ----------
try:
    import openpyxl
    from openpyxl.styles import Font
except ImportError:
    print("缺少 openpyxl 库，正在自动安装（仅首次需要，请稍候）...")
    try:
        subprocess.check_call([sys.executable, "-m", "pip", "install", "openpyxl", "-q"])
        import openpyxl
        from openpyxl.styles import Font
    except Exception as e:
        print(f"自动安装 openpyxl 失败：{e}")
        print("请手动执行：pip install openpyxl")
        input("按回车键退出...")
        sys.exit(1)

# ---------- 路径配置 ----------
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
XLSX_PATH = os.path.join(BASE_DIR, "菜品数据.xlsx")
IMPORTED_DIR = os.path.join(BASE_DIR, "已导入")
SCAN_DIRS = [BASE_DIR, os.path.join(os.path.expanduser("~"), "Downloads")]


def find_json_files(prefix):
    """扫描所有指定前缀的待导入 json 文件"""
    files = []
    for d in SCAN_DIRS:
        if os.path.isdir(d):
            files.extend(glob.glob(os.path.join(d, prefix + "_*.json")))
    return files


def load_records(filepath):
    """从选菜记录 json 文件解析记录列表，兼容多种字段命名"""
    try:
        with open(filepath, "r", encoding="utf-8-sig") as f:
            data = json.load(f)
    except Exception as e:
        print(f"读取文件失败（已跳过）：{filepath}，原因：{e}")
        return []
    if isinstance(data, dict):
        data = data.get("records", data.get("data", []))
    if not isinstance(data, list):
        return []
    records = []
    for item in data:
        if not isinstance(item, dict):
            continue
        rec = {
            "日期": str(item.get("日期", item.get("date", ""))).strip(),
            "餐次": str(item.get("餐次", item.get("meal", ""))).strip(),
            "菜品编号": item.get("菜品编号", item.get("id", "")),
            "菜品名称": str(item.get("菜品名称", item.get("name", ""))).strip(),
        }
        if rec["日期"] and rec["餐次"] and rec["菜品名称"]:
            records.append(rec)
    return records


def merge_records(existing, new_records):
    """按 日期+餐次+菜品编号+菜名 去重合并"""
    seen = set()
    merged = []
    for rec in existing:
        key = (rec.get("日期"), rec.get("餐次"), str(rec.get("菜品编号")), rec.get("菜品名称"))
        if key not in seen:
            seen.add(key)
            merged.append(rec)
    for rec in new_records:
        key = (rec["日期"], rec["餐次"], str(rec["菜品编号"]), rec["菜品名称"])
        if key not in seen:
            seen.add(key)
            merged.append(rec)
    # 按日期排序，同日按早/中/晚顺序
    meal_order = {"早餐": 0, "中餐": 1, "晚餐": 2, "正餐": 3, "甜品": 4, "夜宵": 5}
    merged.sort(key=lambda r: (str(r.get("日期", "")), meal_order.get(str(r.get("餐次", "")), 9)))
    return merged


def load_new_dishes(filepath):
    """从新菜品 json 文件解析菜品列表，兼容多种字段命名"""
    try:
        with open(filepath, "r", encoding="utf-8-sig") as f:
            data = json.load(f)
    except Exception as e:
        print(f"读取文件失败（已跳过）：{filepath}，原因：{e}")
        return []
    if isinstance(data, dict):
        data = data.get("dishes", data.get("items", data.get("data", [])))
    if not isinstance(data, list):
        return []
    dishes = []
    for item in data:
        if not isinstance(item, dict):
            continue
        d = {
            "菜品名称": str(item.get("菜品名称", item.get("name", ""))).strip(),
            "餐次": str(item.get("餐次", item.get("meal", ""))).strip(),
            "分类": str(item.get("分类", item.get("category", ""))).strip(),
            "食材（原料）": str(item.get("食材（原料）", item.get("ingredient", ""))).strip(),
            "菜谱步骤": str(item.get("菜谱步骤", item.get("steps", ""))).strip(),
            "抖音链接": str(item.get("抖音链接", item.get("douyin", item.get("video", "")))).strip(),
            "食材清单": str(item.get("食材清单", item.get("shopping", item.get("list", "")))).strip(),
        }
        if d["菜品名称"] and d["餐次"]:
            dishes.append(d)
    return dishes


def merge_dishes(existing_rows, new_dishes):
    """按菜品名称去重合并，新菜品编号自动分配（现有最大编号+1 依次递增）
    existing_rows: 现有数据行列表（tuple，顺序同表头）
    new_dishes: 新菜品 dict 列表
    返回 (merged_rows, added_count)
    """
    merged = list(existing_rows)
    names = set()
    for r in existing_rows:
        if r and r[1]:
            names.add(str(r[1]).strip())
    next_id = 1
    for r in existing_rows:
        if r and r[0] is not None and str(r[0]).strip().isdigit():
            next_id = max(next_id, int(str(r[0]).strip()) + 1)
    added = 0
    for d in new_dishes:
        name = d["菜品名称"]
        if name in names:
            print(f"  同名跳过（已存在）：{name}")
            continue
        merged.append((
            next_id,
            name,
            d["餐次"],
            d["分类"],
            d["食材（原料）"],
            d["菜谱步骤"],
            d["抖音链接"],
            d["食材清单"],
        ))
        names.add(name)
        next_id += 1
        added += 1
    return merged, added


def write_dish_sheet(ws, rows):
    """重写「菜品数据」工作表：表头加粗、列宽合适、开启筛选 A1:H{末行}"""
    ws.delete_rows(1, ws.max_row)
    headers = ["编号", "菜品名称", "餐次", "分类", "食材（原料）", "菜谱步骤", "抖音链接", "食材清单"]
    ws.append(headers)
    for r in rows:
        ws.append(list(r))
    for cell in ws[1]:
        cell.font = Font(bold=True)
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = "A1:H{}".format(ws.max_row)
    for col, w in {"A": 6, "B": 16, "C": 10, "D": 14, "E": 22, "F": 60, "G": 12, "H": 42}.items():
        ws.column_dimensions[col].width = w


def rebuild_stats_sheet(wb):
    """重建「统计报表」工作表：各餐次/分类菜品数量（动态）+ 菜品被选次数（COUNTIF 公式）"""
    idx = 2
    if "统计报表" in wb.sheetnames:
        idx = wb.sheetnames.index("统计报表")
        wb.remove(wb["统计报表"])
    ws = wb.create_sheet("统计报表", index=idx)

    ws_dish = wb["菜品数据"]
    dish_rows = []
    for row in ws_dish.iter_rows(min_row=2, values_only=True):
        if row and row[0] is not None and str(row[0]).strip() != "":
            dish_rows.append(row)

    # 动态统计餐次（固定顺序：早餐/正餐/甜品 + 自定义按出现顺序）
    meal_counts = {}
    for r in dish_rows:
        m = str(r[2]).strip() if r[2] is not None else ""
        if m:
            meal_counts[m] = meal_counts.get(m, 0) + 1
    meal_order = {"早餐": 0, "正餐": 1, "甜品": 2}
    meals = sorted(meal_counts.keys(), key=lambda m: (meal_order.get(m, 99), m))

    # 动态统计分类（按数量降序，数量相同按出现顺序）
    cat_counts = {}
    for r in dish_rows:
        c = str(r[3]).strip() if r[3] is not None else ""
        if c:
            cat_counts[c] = cat_counts.get(c, 0) + 1
    cats = sorted(cat_counts.keys(), key=lambda c: (-cat_counts[c], list(cat_counts.keys()).index(c)))

    # 全部菜品名称（用于被选次数表）
    names = []
    for r in dish_rows:
        n = str(r[1]).strip() if r[1] is not None else ""
        if n:
            names.append(n)

    ws["A1"] = "各餐次菜品数量"
    ws["D1"] = "各分类菜品数量"
    ws["G1"] = "菜品被选次数"
    ws["A2"] = "餐次"
    ws["B2"] = "菜品数量"
    ws["D2"] = "分类"
    ws["E2"] = "菜品数量"
    ws["G2"] = "菜品名称"
    ws["H2"] = "被选次数"

    r = 3
    for m in meals:
        ws.cell(row=r, column=1, value=m)
        ws.cell(row=r, column=2, value=meal_counts[m])
        r += 1
    ws.cell(row=r, column=1, value="合计")
    ws.cell(row=r, column=2, value=len(dish_rows))

    r = 3
    for c in cats:
        ws.cell(row=r, column=4, value=c)
        ws.cell(row=r, column=5, value=cat_counts[c])
        r += 1
    ws.cell(row=r, column=4, value="合计")
    ws.cell(row=r, column=5, value=len(dish_rows))

    r = 3
    for n in names:
        ws.cell(row=r, column=7, value=n)
        ws.cell(row=r, column=8, value="=COUNTIF(选菜记录!$D:$D,G{})".format(r))
        r += 1

    # 表头加粗
    for row in (1, 2):
        for col in range(1, 9):
            ws.cell(row=row, column=col).font = Font(bold=True)

    # 列宽
    for col, w in {"A": 14, "B": 12, "C": 6, "D": 14, "E": 12, "F": 6, "G": 22, "H": 10}.items():
        ws.column_dimensions[col].width = w


def archive_files(files):
    """把已处理的 json 文件移到「已导入」文件夹"""
    os.makedirs(IMPORTED_DIR, exist_ok=True)
    for fp in files:
        try:
            shutil.move(fp, os.path.join(IMPORTED_DIR, os.path.basename(fp)))
            print(f"  已归档：{os.path.basename(fp)} -> 已导入/")
        except Exception as e:
            print(f"  归档失败（不影响导入）：{fp}，原因：{e}")


def sync():
    if not os.path.exists(XLSX_PATH):
        print(f"未找到 Excel 文件：{XLSX_PATH}")
        print("请确认脚本放在「慧做菜」文件夹中且存在 菜品数据.xlsx")
        return False

    rec_files = find_json_files("选菜记录")
    dish_files = find_json_files("新菜品")
    if not rec_files and not dish_files:
        print("未找到待导入的 json 文件。")
        print("请先在网页导出 选菜记录_*.json 或 新菜品_*.json，")
        print("放到慧做菜文件夹或下载目录，再运行本脚本。")
        return False

    wb = openpyxl.load_workbook(XLSX_PATH)
    changed = False

    # ===== 1. 同步选菜记录 =====
    rec_new_count = 0
    rec_total = 0
    rec_archives = []
    if rec_files:
        if "选菜记录" not in wb.sheetnames:
            ws = wb.create_sheet("选菜记录")
            ws.append(["日期", "餐次", "菜品编号", "菜品名称"])
        ws = wb["选菜记录"]
        existing = []
        for row in ws.iter_rows(min_row=2, values_only=True):
            if row and row[0]:
                existing.append({
                    "日期": str(row[0]), "餐次": str(row[1]),
                    "菜品编号": row[2], "菜品名称": str(row[3])
                })
        new_records = []
        valid_rec_files = []
        for fp in rec_files:
            recs = load_records(fp)
            if recs:
                print(f"读取 {os.path.basename(fp)}：{len(recs)} 条记录")
                new_records.extend(recs)
                valid_rec_files.append(fp)
        if new_records:
            merged = merge_records(existing, new_records)
            ws.delete_rows(2, ws.max_row - 1)
            for rec in merged:
                ws.append([rec["日期"], rec["餐次"], rec["菜品编号"], rec["菜品名称"]])
            for col, w in {"A": 14, "B": 10, "C": 12, "D": 30}.items():
                ws.column_dimensions[col].width = w
            rec_new_count = len(new_records)
            rec_total = len(merged)
            rec_archives = valid_rec_files
            changed = True
        else:
            print("选菜记录 json 中没有有效记录（已跳过）。")

    # ===== 2. 同步新菜品 =====
    dish_added = 0
    dish_total = 0
    dish_archives = []
    if dish_files:
        new_dishes = []
        valid_dish_files = []
        for fp in dish_files:
            ds = load_new_dishes(fp)
            if ds:
                print(f"读取 {os.path.basename(fp)}：{len(ds)} 道新菜品")
                new_dishes.extend(ds)
                valid_dish_files.append(fp)
        if new_dishes:
            ws_dish = wb["菜品数据"]
            existing_rows = []
            for row in ws_dish.iter_rows(min_row=2, values_only=True):
                if row and row[0] is not None and str(row[0]).strip() != "":
                    existing_rows.append(row)
            merged_rows, added = merge_dishes(existing_rows, new_dishes)
            dish_added = added
            dish_total = len(merged_rows)
            if added > 0:
                write_dish_sheet(ws_dish, merged_rows)
                rebuild_stats_sheet(wb)
                changed = True
            else:
                print("新菜品 json 中所有菜品均已存在（未新增）。")
            dish_archives = valid_dish_files
        else:
            print("新菜品 json 中没有有效菜品（已跳过）。")

    # ===== 3. 保存与归档 =====
    if changed:
        wb.save(XLSX_PATH)
        print("=" * 46)
        print("  同步完成！")
        if rec_new_count:
            print(f"  本次新导入选菜记录：{rec_new_count} 条")
            print(f"  选菜记录总计：{rec_total} 条")
        if dish_files:
            print(f"  本次新增 {dish_added} 道新菜品、总计 {dish_total} 道菜")
        print(f"  Excel：{XLSX_PATH}")
        print("  （打开 Excel 统计报表可查看每道菜被选次数）")
        print("=" * 46)
    else:
        print("没有可写入的新数据，Excel 未修改。")

    # 归档有效 json
    archive_files(rec_archives)
    archive_files(dish_archives)
    return changed


if __name__ == "__main__":
    no_pause = "--no-pause" in sys.argv
    print("=" * 46)
    print("  选菜记录 / 新菜品 自动同步工具")
    print("=" * 46)
    try:
        sync()
    except Exception as e:
        print(f"同步过程出错：{e}")
        import traceback
        traceback.print_exc()
    print()
    if not no_pause:
        input("按回车键退出...")
