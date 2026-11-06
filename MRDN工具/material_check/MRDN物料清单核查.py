# -*- coding: utf-8 -*-
"""
MRDN 物料清单核查工具（material_check）
=================================

功能：
    把「最新物料设计表」（宽表：DU × itemcode 数量矩阵）按 ML_transfor 逻辑
    转置成新物料清单（长表），再和「物料制作清单」（已转置好的旧长表）对比，
    按「DUID + itemcode」维度找出差异，标记 Date ready for configuration，
    输出一份「对比差异表」。

对比规则：
    1. 旧有、新无  → 删减标签 = "需剔除"
    2. 新有、旧无  → 删减标签 = "需新增"
    3. 都有但数量不同 → 删减标签 = "数量变化"
    4. 完全相同（DUID+itemcode 数量一致）→ 不计入差异表

依赖：openpyxl
    pip install openpyxl
"""


# ============================================================================
#  ⚙ MRDN 系统集成参数对照表（配置系统「工具输入关联」时，照此填「参数名」）
# ============================================================================
#   参数名（填到系统）       对应文件 / 含义
#   --material_list          物料制作清单（旧长表，含 DUID/itemcode/qty/type 列）
#   --design_table           最新物料设计表（新宽表，第 2 行 itemcode、第 3 行描述、数据行起）
#   --type_table             物料类型表（itemcode → type 映射）
#   --output_template        物料转置输出模板（转置结果表头，含 Date ready 列）
#   --du_list                DUID 清单（批次动态：只核查清单内的 DUID；Excel 或 txt）
#   （--out_dir 为系统自动传入的输出目录，无需在系统里配置）
# ============================================================================

import os
import re
import sys
import math
import argparse
from datetime import datetime

import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side


# ============================================================
# 基础路径
# ============================================================
if getattr(sys, 'frozen', False):
    BASE_DIR = os.path.dirname(os.path.abspath(sys.executable))
else:
    BASE_DIR = os.path.dirname(os.path.abspath(__file__))

OUTPUT_FILENAME = "对比差异表.xlsx"


# ============================================================
# 常量（转置宽表时的行号约定，同 ML_transfor）
# ============================================================
ITEMCODE_ROW = 2        # itemcode 所在行
DESC_ROW = 3            # 描述所在行
DATA_START_ROW = 4      # 数据起始行


# ============================================================
# 转置后替换规则（itemcode 替换 + 数量比例折算，同 ML_transfor）
# ============================================================
CODE_REPLACE_RULES = [
    ("30070011", "02237428", 1, "1:1"),
    ("28010521", "02237428", 2, "2:1"),
    ("04070012", "04023275", 1, "1:1"),
]


# ============================================================
# 工具函数
# ============================================================
def to_number(v):
    if isinstance(v, bool):
        return None
    if isinstance(v, (int, float)):
        return v
    if isinstance(v, str):
        s = v.strip()
        if s == "":
            return None
        try:
            return float(s)
        except ValueError:
            return None
    return None


def _norm(s):
    return re.sub(r"[\s\-_]+", "", str(s)).lower()


def _norm_duid(s):
    return re.sub(r"\s+", "", str(s)).lower()


def _norm_code(s):
    s = str(s).strip()
    if s.isdigit():
        s = s.lstrip("0") or "0"
    return s


def has_any_keyword(text, keywords):
    up = str(text).upper()
    return any(k.upper() in up for k in keywords)


def _parse_date(v):
    if v is None:
        return None
    if isinstance(v, datetime):
        return v
    if isinstance(v, str):
        s = v.strip()
        if s == "":
            return None
        for fmt in ("%d-%b", "%d-%b-%Y", "%d-%b-%y", "%Y-%m-%d", "%Y/%m/%d"):
            try:
                return datetime.strptime(s, fmt)
            except ValueError:
                continue
    return None


def _format_date(v):
    if v is None:
        return ""
    if isinstance(v, datetime):
        return str(v.day) + "-" + v.strftime("%b")
    return str(v).strip()


def _s(v):
    if v is None:
        return ""
    return str(v).strip()


# ============================================================
# 读取物料类型表：ITEM编码 -> 设备类型
# ============================================================
def read_type_map(path):
    wb = openpyxl.load_workbook(path, data_only=True)
    ws = wb.active
    header = [_s(c.value) for c in ws[1]]
    col_code = col_type = None
    for i, h in enumerate(header):
        nh = _norm(h)
        if nh in ("item编码", "itemcode", "编码", "物料编码"):
            col_code = i
        elif nh in ("设备类型", "type", "类型"):
            col_type = i
    if col_code is None or col_type is None:
        col_code, col_type = 0, 2
    type_map = {}
    for r in range(2, ws.max_row + 1):
        code = ws.cell(row=r, column=col_code + 1).value
        t = ws.cell(row=r, column=col_type + 1).value
        if code is None:
            continue
        code = str(code).strip()
        if code == "":
            continue
        type_map[code] = _s(t)
    return type_map


def lookup_type(type_map, code):
    if code in type_map:
        return type_map[code]
    if code.isdigit() and len(code) < 8:
        padded = code.zfill(8)
        if padded in type_map:
            return type_map[padded]
    return ""


# ============================================================
# 读取 DUID 清单（批次动态筛选）
# ============================================================
def read_du_list(path):
    """读取「DUID 清单」——含 DUID 列的前置条件清单（批次）。

    自动识别 DUID 列：表头列名归一化后为 duid / du（兼容 DU ID、DU_ID、du id、
    DUID、DU 等写法），只取该列的非空值作为目标 DUID。

    支持两种格式：
      - Excel（.xlsx/.xlsm/.xls）：找含 DUID 列名的表头行，取其下一行起该列的值。
      - 文本（.txt/.csv）：优先识别表头（含 DUID 列名）；否则每行一个 DUID。

    返回 dict：{归一化 DUID: 原始 DUID}（去重，保留第一个出现的原始值）。
    找不到 DUID 列 / 读取失败时返回空 dict（视为未提供清单，核查全部 DU）。
    """
    du_map = {}
    if not path or not os.path.exists(path):
        return du_map
    ext = os.path.splitext(path)[1].lower()
    try:
        if ext in (".xlsx", ".xlsm", ".xls"):
            wb = openpyxl.load_workbook(path, data_only=True, read_only=True)
            ws = wb.active
            rows = list(ws.iter_rows(values_only=True))
            wb.close()

            du_col = None
            header_idx = None
            for ri, row in enumerate(rows):
                for ci, v in enumerate(row):
                    if v is not None and _norm(v) in ("duid", "du", "duid2"):
                        du_col = ci
                        header_idx = ri
                        break
                if du_col is not None:
                    break
            if du_col is None:
                print("[提示] DUID 清单中未找到 DUID 列（期望 DU ID / DU_ID / du id / DUID 等），将核查全部 DU")
                return {}
            for row in rows[header_idx + 1:]:
                v = row[du_col] if du_col < len(row) else None
                if v is None:
                    continue
                s = str(v).strip()
                if s:
                    du_map.setdefault(_norm_duid(s), s)
        else:
            with open(path, "r", encoding="utf-8-sig", errors="replace") as f:
                lines = [ln.strip() for ln in f if ln.strip()]
            if not lines:
                return {}
            first_parts = re.split(r"[\t,，;；]", lines[0])
            du_col = next((i for i, p in enumerate(first_parts)
                           if p and _norm(p) in ("duid", "du", "duid2")), None)
            if du_col is not None:
                for line in lines[1:]:
                    parts = re.split(r"[\t,，;；]", line)
                    v = parts[du_col] if du_col < len(parts) else ""
                    s = v.strip()
                    if s:
                        du_map.setdefault(_norm_duid(s), s)
            else:
                for line in lines:
                    for p in re.split(r"[\t,，;；]", line):
                        s = p.strip()
                        if s:
                            du_map.setdefault(_norm_duid(s), s)
                            break
        return du_map
    except Exception as e:
        print("[警告] 读取 DUID 清单失败（将核查全部 DU）：%s" % e)
        return {}


# ============================================================
# 读取宽表（最新物料设计表）
# ============================================================
def read_wide(path):
    """读宽表，返回 (data_rows, du_col, date_col, code_columns)。"""
    wb = openpyxl.load_workbook(path, data_only=True, read_only=True)
    ws = wb.active
    rows = list(ws.iter_rows(values_only=True))
    wb.close()

    if len(rows) < DESC_ROW:
        return [], None, None, []

    itemcode_row = rows[ITEMCODE_ROW - 1]
    desc_row = rows[DESC_ROW - 1]

    du_col = None
    date_col = None
    for i, v in enumerate(desc_row):
        s = _norm(str(v)) if v is not None else ""
        if du_col is None and s in ("duid", "du", "duid2"):
            du_col = i
        if date_col is None and "date" in s and "ready" in s:
            date_col = i
    if du_col is None:
        du_col = 1

    code_columns = []
    for i, code in enumerate(itemcode_row):
        if code is None:
            continue
        code = str(code).strip()
        if code == "":
            continue
        desc = desc_row[i] if i < len(desc_row) else None
        desc = _s(desc)
        code_columns.append((i, code, desc))

    data_rows = rows[DATA_START_ROW - 1:]
    return data_rows, du_col, date_col, code_columns


def collect_wide(data_rows, du_col, date_col, code_columns, du_set=None):
    """把宽表转置成 [(duid, date, [(code, desc, qty)...])]，重复 DU 取日期最新。

    返回 (result, seen)：seen 是模板里实际出现的 DUID（归一化）集合。
    du_set 非空时只转置清单内的 DU。"""
    best = {}
    seen = set()
    for row in data_rows:
        if du_col is None or du_col >= len(row):
            continue
        duid = row[du_col]
        if duid is None or _s(duid) == "":
            continue
        duid = _s(duid)
        if duid.lower() == "remark":
            continue
        norm = _norm_duid(duid)
        seen.add(norm)
        if du_set and norm not in du_set:
            continue
        raw_date = row[date_col] if date_col is not None and date_col < len(row) else None
        date_key = _parse_date(raw_date)
        if norm not in best:
            best[norm] = (duid, raw_date, row)
        else:
            prev_key = _parse_date(best[norm][1])
            if date_key is not None and (prev_key is None or date_key > prev_key):
                best[norm] = (duid, raw_date, row)

    result = []
    for duid, raw_date, row in best.values():
        agg = {}
        order = []
        for c, code, desc in code_columns:
            qty = to_number(row[c] if c < len(row) else None)
            if qty is None or qty <= 0:
                continue
            if code not in agg:
                agg[code] = {"qty": 0.0, "desc": ""}
                order.append(code)
            agg[code]["qty"] += qty
            if not agg[code]["desc"] and desc:
                agg[code]["desc"] = desc
        items = []
        for code in order:
            q = agg[code]["qty"]
            q = int(q) if float(q).is_integer() else q
            items.append((code, agg[code]["desc"], q))
        result.append((duid, raw_date, items))
    return result, seen


def apply_replacements(data):
    """itemcode 替换 + 数量比例折算（同 ML_transfor）。"""
    result = []
    for duid, date, items in data:
        new_items = []
        for code, desc, qty in items:
            target, divisor, ratio_str = code, 1, ""
            for src, dst, div, ratio in CODE_REPLACE_RULES:
                if _norm_code(code) == _norm_code(src):
                    target, divisor, ratio_str = dst, div, ratio
                    break
            if target != code:
                new_qty = qty if divisor == 1 else math.ceil(qty / divisor)
                new_items.append((target, desc, new_qty, "是", code, qty, ratio_str))
            else:
                new_items.append((code, desc, qty, "", "", "", ""))
        result.append((duid, date, new_items))
    return result


def resolve_type(code, desc, type_map):
    orig = lookup_type(type_map, code)
    final = orig
    if code[:2] == "34":
        final = "Type 3"
    elif code[:1] == "5":
        final = "Type 1"
    elif has_any_keyword(desc, ["breaker", "AMB"]):
        final = "Type 1"
    return final


# ============================================================
# 读取物料制作清单（旧长表）
# ============================================================
def _find_col(headers, *names):
    targets = {_norm(n) for n in names}
    for i, h in enumerate(headers):
        if _norm(h) in targets:
            return i
    return None


def read_material_list(path, du_set=None):
    """读旧长表，返回 {(norm_duid, norm_code): {duid, code, qty, desc, type, date}}。
    du_set 非空时只保留清单内的 DU。"""
    wb = openpyxl.load_workbook(path, data_only=True)
    ws = wb.active
    headers = [_s(c.value) for c in ws[1]]

    ci_du = _find_col(headers, "DUID", "DU ID", "DU", "DU_ID", "duid")
    ci_code = _find_col(headers, "itemcode", "item code", "item", "物料编码", "material code")
    ci_qty = _find_col(headers, "qty", "数量", "quantity", "Qty")
    ci_desc = _find_col(headers, "描述", "desc", "description", "item desc")
    ci_type = _find_col(headers, "type", "设备类型", "Type")
    ci_date = _find_col(headers, "Date ready for configuration", "date ready")

    if ci_du is None or ci_code is None or ci_qty is None:
        raise RuntimeError(
            "物料制作清单缺少必需列：DUID / itemcode / qty（识别到表头：%s）" % headers)

    result = {}
    for r in range(2, ws.max_row + 1):
        duid = _s(ws.cell(row=r, column=ci_du + 1).value)
        code = _s(ws.cell(row=r, column=ci_code + 1).value)
        if duid == "" or code == "":
            continue
        if du_set and _norm_duid(duid) not in du_set:
            continue
        qty = to_number(ws.cell(row=r, column=ci_qty + 1).value)
        if qty is None:
            qty = 0.0
        desc = _s(ws.cell(row=r, column=ci_desc + 1).value) if ci_desc is not None else ""
        typ = _s(ws.cell(row=r, column=ci_type + 1).value) if ci_type is not None else ""
        date = ws.cell(row=r, column=ci_date + 1).value if ci_date is not None else None

        key = (_norm_duid(duid), _norm_code(code))
        if key in result:
            result[key]["qty"] += qty
        else:
            result[key] = {
                "duid": duid, "code": code, "qty": qty,
                "desc": desc, "type": typ, "date": date,
            }
    return result


# ============================================================
# 对比逻辑
# ============================================================
def build_new_map(wide_data, type_map):
    """把转置结果转成 {(norm_duid, norm_code): {duid, code, qty, desc, type, date}}。"""
    new_map = {}
    for duid, date, items in wide_data:
        for code, desc, qty, is_rcode, src_code, src_qty, ratio in items:
            final = resolve_type(code, desc, type_map)
            key = (_norm_duid(duid), _norm_code(code))
            if key in new_map:
                new_map[key]["qty"] += qty
            else:
                new_map[key] = {
                    "duid": duid, "code": code, "qty": qty,
                    "desc": desc, "type": final, "date": date,
                }
    return new_map


def compare(old_map, new_map):
    """对比新旧，返回差异行列表，每行是 dict。"""
    rows = []

    # 需剔除：旧有、新无
    for key, old in old_map.items():
        if key not in new_map:
            rows.append({
                "duid": old["duid"], "code": old["code"],
                "desc": old["desc"], "type": old["type"],
                "old_qty": old["qty"], "new_qty": "",
                "date": old["date"], "label": "需剔除",
            })

    # 需新增：新有、旧无
    for key, new in new_map.items():
        if key not in old_map:
            rows.append({
                "duid": new["duid"], "code": new["code"],
                "desc": new["desc"], "type": new["type"],
                "old_qty": "", "new_qty": new["qty"],
                "date": new["date"], "label": "需新增",
            })

    # 数量变化：都有、数量不同
    for key in old_map.keys() & new_map.keys():
        old = old_map[key]
        new = new_map[key]
        if _num(old["qty"]) != _num(new["qty"]):
            rows.append({
                "duid": new["duid"], "code": new["code"],
                "desc": new["desc"] or old["desc"],
                "type": new["type"] or old["type"],
                "old_qty": old["qty"], "new_qty": new["qty"],
                "date": new["date"] or old["date"], "label": "数量变化",
            })

    return rows


def _num(v):
    try:
        return float(v)
    except (TypeError, ValueError):
        return 0.0


# ============================================================
# 输出差异表
# ============================================================
HEADER_FILL = PatternFill("solid", fgColor="C00000")
HEADER_FONT = Font(bold=True, color="FFFFFF", name="微软雅黑", size=11)
DATA_FONT = Font(name="微软雅黑", size=11)
CENTER = Alignment(horizontal="center", vertical="center", wrap_text=True)
LEFT = Alignment(horizontal="left", vertical="center", wrap_text=True)
THIN = Side(style="thin", color="BFBFBF")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)

DIFF_HEADERS = ["DUID", "itemcode", "描述", "type",
                "旧数量", "新数量", "Date ready for configuration", "删减标签"]


def build_diff_output(rows, out_path):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "对比差异表"

    for i, h in enumerate(DIFF_HEADERS, start=1):
        c = ws.cell(row=1, column=i, value=h)
        c.fill = HEADER_FILL
        c.font = HEADER_FONT
        c.alignment = CENTER
        c.border = BORDER
    ws.freeze_panes = "A2"

    # 排序：DUID 相同挨在一起；标签按 需剔除→需新增→数量变化；再按 itemcode
    label_order = {"需剔除": 0, "需新增": 1, "数量变化": 2}
    rows.sort(key=lambda r: (
        _norm_duid(r["duid"]),
        label_order.get(r["label"], 9),
        _norm_code(r["code"]),
    ))

    for ri, r in enumerate(rows, start=2):
        vals = [
            r["duid"], r["code"], r["desc"], r["type"],
            r["old_qty"], r["new_qty"], _format_date(r["date"]), r["label"],
        ]
        for ci, v in enumerate(vals, start=1):
            c = ws.cell(row=ri, column=ci, value=v)
            c.font = DATA_FONT
            c.border = BORDER
            c.alignment = LEFT if ci == 3 else CENTER

    widths = [26, 18, 55, 10, 10, 10, 22, 12]
    for i, w in enumerate(widths, start=1):
        ws.column_dimensions[openpyxl.utils.get_column_letter(i)].width = w

    wb.save(out_path)
    return len(rows)


# ============================================================
# CLI
# ============================================================
def _configure_stdout():
    if hasattr(sys.stdout, "reconfigure"):
        try:
            sys.stdout.reconfigure(encoding="utf-8", errors="replace")
            sys.stderr.reconfigure(encoding="utf-8", errors="replace")
        except Exception:
            pass


def parse_args():
    p = argparse.ArgumentParser(description="MRDN 物料清单核查工具")
    p.add_argument("--material_list", default=None, help="物料制作清单（旧长表）路径")
    p.add_argument("--design_table", default=None, help="最新物料设计表（新宽表）路径")
    p.add_argument("--type_table", default=None, help="物料类型表路径")
    p.add_argument("--output_template", default=None, help="物料转置输出模板路径")
    p.add_argument("--du_list", default=None, help="DUID 清单（批次）路径")
    p.add_argument("--out_dir", default=None, help="输出目录（系统自动传入）")
    return p.parse_args()


def main():
    _configure_stdout()
    args = parse_args()

    out_dir = args.out_dir or BASE_DIR
    os.makedirs(out_dir, exist_ok=True)
    out_file = os.path.join(out_dir, OUTPUT_FILENAME)

    for p, name in ((args.material_list, "物料制作清单"),
                    (args.design_table, "最新物料设计表"),
                    (args.type_table, "物料类型表"),
                    (args.output_template, "物料转置输出模板")):
        if not p or not os.path.exists(p):
            print("[错误] 找不到%s：%s" % (name, p))
            sys.exit(1)

    # 1. 物料类型表
    type_map = read_type_map(args.type_table)
    print("物料类型表编码数：", len(type_map))

    # 2. DUID 清单（批次动态筛选）
    du_map = read_du_list(args.du_list)
    du_set = set(du_map.keys())
    if du_set:
        print("DUID 清单（批次）DUID 数：", len(du_set))
    else:
        print("未提供 DUID 清单（或为空），将核查全部 DU")

    # 3. 转置「最新物料设计表」→ 新长表（只转置清单内 DU）
    data_rows, du_col, date_col, code_columns = read_wide(args.design_table)
    print("最新设计表 itemcode 列数：", len(code_columns))
    wide_data, template_duids = collect_wide(data_rows, du_col, date_col, code_columns, du_set)

    # 校验：DUID 清单里有、但最新设计表里找不到的 DUID → 报错退出
    if du_set:
        missing = sorted(du_map[n] for n in du_set if n not in template_duids)
        if missing:
            print("[错误] DUID 清单中以下 DUID 在最新设计表中找不到，请核对清单：")
            for m in missing:
                print("   -", m)
            sys.exit(1)

    wide_data = apply_replacements(wide_data)
    new_map = build_new_map(wide_data, type_map)
    print("新物料清单（转置后）记录数：", len(new_map))

    # 4. 读「物料制作清单」→ 旧长表（只保留清单内 DU）
    old_map = read_material_list(args.material_list, du_set)
    print("物料制作清单（旧）记录数：", len(old_map))

    # 5. 对比
    diffs = compare(old_map, new_map)
    print("差异条数：", len(diffs))

    # 6. 输出
    n = build_diff_output(diffs, out_file)
    print("\n已生成对比差异表：", out_file)
    print("差异明细：需剔除 %d 条 / 需新增 %d 条 / 数量变化 %d 条" % (
        sum(1 for d in diffs if d["label"] == "需剔除"),
        sum(1 for d in diffs if d["label"] == "需新增"),
        sum(1 for d in diffs if d["label"] == "数量变化"),
    ))


if __name__ == "__main__":
    try:
        main()
    except SystemExit:
        raise
    except Exception as e:
        _configure_stdout()
        print("[错误] 运行失败：%s" % e)
        sys.exit(1)
