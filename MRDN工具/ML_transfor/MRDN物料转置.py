# -*- coding: utf-8 -*-
"""
MRDN 物料转置工具（ML_transfor）
================================

功能：
    把「输入模板.xlsx」转置成「输出模板.xlsx」的样子：
    统计每个 DU ID 下各种物料（itemcode）的数量，通过「物料类型表.xlsx」
    映射 type（设备类型），并应用若干替换/发包规则。

    新增：提供「DU清单」（含 DUID 列的前置条件清单，列名可写作 DU ID / DU_ID /
    du id / DUID 等）时，用清单中的 DUID 筛选输入模板 B 列的 DU，只对清单内
    的 DU 做转置；不提供清单则转置全部 DU（向后兼容）。

转换规则：
    1. itemcode 取自输入模板第 2 行；凡第 2 行有编码值的列都算候选列。
       数量 = 该列数据行（第 4 行起）下面的值，且必须是「纯数字」才统计。
    2. 只保留数量 > 0 的物料。
    3. 同一 itemcode 出现在多个列时，去重后数量相加。
    4. 「描述」列 = itemcode 下一行（第 3 行）的原文。
    5. type 映射：
       - 输入表 itemcode == 物料类型表 ITEM编码 -> 设备类型；
       - 纯数字编码查不到时，尝试补前导 0 到 8 位再查一次（容错）。
    6. 替换规则（type 与「是否替换类型 / 替换前类型」联动）：
       a. code 前 2 位为 "34"  -> type 改为 Type 3；
       b. code 前 1 位为 "5"   -> type 改为 Type 1；
       c. 描述含关键字 "breaker" 或 "AMB" -> type 改为 Type 1；
       （触发替换时：「是否替换类型」填 "是"，「替换前类型」填替换前的 type）
    7. 发包规则：最终 type 为 Type 1 且描述含 "RRU"/"AAU"/"BBU" 关键字时，
       「是否发包」列填 "发包"，否则为空。
    8. 转置后替换处理（itemcode 替换 + 数量按比例折算，联动
       「是否替换code / 替换前code / 替换前数量 / 替换比例」四列）：
       a. 30070011 -> 02237428，比例 1:1（数量不变）；
       b. 28010521 -> 02237428，比例 2:1（2 个 28010521 = 1 个 02237428，
          数量向上取整，保证物料整数够用）；
       c. 04070012 -> 04023275，比例 1:1（数量不变）。

    9. Date ready for configuration 读取与输出：
       读取输入表第 3 行表头含「Date ready for configuration」的列，输出时追加
       「Date ready for configuration」列（值格式化为「5-Jan」简写）；同一 DU 出现
       多行时取日期最新的那行。

性能：
    用 read_only 模式一次性读入整个工作表，再按 DU 清单筛选目标行（重复 DU 取
    日期最新）后转置，避免大量数据下全表逐 cell 遍历慢。

输出：
    结果写到同目录「输出结果.xlsx」，表头沿用「输出模板.xlsx」并追加
    「Date ready for configuration」列，带样式。

依赖：openpyxl
    pip install openpyxl
"""


# ============================================================================
#  ⚙ MRDN 系统集成参数对照表（配置系统「工具输入关联」时，照此填「参数名」）
# ============================================================================
#   参数名（填到系统）       对应文件 / 含义
#   --input_template         输入模板.xlsx（物料明细表）
#   --output_template        输出模板.xlsx（结果表头）
#   --type_table             物料类型表.xlsx（itemcode→type 映射）
#   --du_list                DU清单（含 DUID 列的表格，自动识别 DU ID / DU_ID / du id 等列名；只转置清单内的 DU）
#   --duid                   单 DUID（作业清单里 single 节点只跑当前 DUID 时传；优先级高于 --du_list，只转置这一个 DUID）
#   （--out_dir 为系统自动传入的输出目录，无需在系统里配置）
# ============================================================================

import os
import re
import sys
import argparse
import math
from datetime import datetime

import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side


# ============================================================
# 可配置项
# ============================================================
if getattr(sys, 'frozen', False):
    BASE_DIR = os.path.dirname(os.path.abspath(sys.executable))
else:
    BASE_DIR = os.path.dirname(os.path.abspath(__file__))

INPUT_TEMPLATE = os.path.join(BASE_DIR, "输入模板.xlsx")     # 输入模板
OUTPUT_TEMPLATE = os.path.join(BASE_DIR, "输出模板.xlsx")    # 输出模板（提供表头）
TYPE_TABLE = os.path.join(BASE_DIR, "物料类型表.xlsx")       # 物料类型表（itemcode -> type）
OUTPUT_FILE = os.path.join(BASE_DIR, "输出结果.xlsx")        # 最终输出

ITEMCODE_ROW = 2        # itemcode 所在行（Excel 第 2 行）
DESC_ROW = 3            # 描述所在行（Excel 第 3 行）
DATA_START_ROW = 4      # 数据起始行（Excel 第 4 行起）


# ============================================================
# 工具函数
# ============================================================
def to_number(v):
    """把单元格值转成数字（纯数字才返回，其余返回 None）。"""
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
    """DUID 归一化：去空白 + 小写（保留 - 和 _，避免误删 DU 编号里的分隔符）。"""
    return re.sub(r"\s+", "", str(s)).lower()


def has_any_keyword(text, keywords):
    """描述文本是否包含任一关键字（大小写不敏感）。"""
    up = str(text).upper()
    return any(k.upper() in up for k in keywords)


def _parse_date(v):
    """把「Date ready for configuration」单元格值解析成 datetime 用于比较。

    兼容三种情况：
      - datetime 对象（openpyxl 读日期单元格的默认结果）
      - 字符串 "5-Jan" / "5-Jan-2026" / "2026-01-05" 等日期简写
      - None / 空（无日期，返回 None）
    解析失败返回 None。
    """
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
    """把日期值格式化成「5-Jan」简写（datetime -> '5-Jan'；字符串原样；空 -> ''）。"""
    if v is None:
        return ""
    if isinstance(v, datetime):
        return str(v.day) + "-" + v.strftime("%b")
    return str(v).strip()


# ============================================================
# 读取物料类型表：ITEM编码 -> 设备类型
# ============================================================
def read_type_map(path):
    wb = openpyxl.load_workbook(path, data_only=True)
    ws = wb.active

    header = [str(c.value).strip() if c.value is not None else "" for c in ws[1]]
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
        type_map[code] = str(t).strip() if t is not None else ""
    return type_map


def lookup_type(type_map, code):
    """查 type；纯数字编码查不到时补前导 0 到 8 位再查一次。"""
    if code in type_map:
        return type_map[code]
    if code.isdigit() and len(code) < 8:
        padded = code.zfill(8)
        if padded in type_map:
            return type_map[padded]
    return ""


# ============================================================
# 读取 DU 清单（前置条件 DUID 列表）
# ============================================================
def read_du_list(path):
    """读取「DU 清单」——含 DUID 列的前置条件清单。

    自动识别 DUID 列：表头列名归一化后为 duid / du（兼容 DU ID、DU_ID、du id、
    DUID、DU 等写法），只取该列的非空值作为前置条件 DUID。

    支持两种格式：
      - Excel（.xlsx/.xlsm/.xls）：找含 DUID 列名的表头行，取其下一行起该列的值。
      - 文本（.txt/.csv）：优先识别表头（含 DUID 列名）；否则每行一个 DUID。

    返回 dict：{归一化 DUID: 原始 DUID}（去重，保留第一个出现的原始值）。
    找不到 DUID 列时返回空 dict（视为未提供清单，转置全部 DU）。
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

            # 1) 找含 DUID 列名的表头行
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
                print("[提示] DU 清单中未找到 DUID 列（期望列名 DU ID / DU_ID / du id / DUID 等），将转置全部 DU")
                return {}

            # 2) 从表头行下一行起，取 DUID 列的值
            for row in rows[header_idx + 1:]:
                v = row[du_col] if du_col < len(row) else None
                if v is None:
                    continue
                s = str(v).strip()
                if s:
                    du_map.setdefault(_norm_duid(s), s)
        else:
            # 文本：优先识别第一行是否含 DUID 表头；否则每行一个 DUID
            with open(path, "r", encoding="utf-8-sig", errors="replace") as f:
                lines = [ln.strip() for ln in f if ln.strip()]
            if not lines:
                return {}
            first_parts = re.split(r"[\t,，;；]", lines[0])
            du_col = next((i for i, p in enumerate(first_parts)
                           if p and _norm(p) in ("duid", "du", "duid2")), None)
            if du_col is not None:
                # 第一行是表头，从第二行起取该列
                for line in lines[1:]:
                    parts = re.split(r"[\t,，;；]", line)
                    v = parts[du_col] if du_col < len(parts) else ""
                    s = v.strip()
                    if s:
                        du_map.setdefault(_norm_duid(s), s)
            else:
                # 无表头：每行一个 DUID（取每行第一个非空字段）
                for line in lines:
                    for p in re.split(r"[\t,，;；]", line):
                        s = p.strip()
                        if s:
                            du_map.setdefault(_norm_duid(s), s)
                            break
        return du_map
    except Exception as e:
        print("[警告] 读取 DU 清单失败（将转置全部 DU）：%s" % e)
        return {}


# ============================================================
# 读取输入
# ============================================================
def read_input(path):
    """read_only 模式一次性读入整个工作表（性能优于逐 cell 访问）。

    返回 (data_rows, du_col, date_col, code_columns)：
      - data_rows: 数据行列表（0-based 索引 tuple，从 DATA_START_ROW 行起）
      - du_col: DU_ID 列 0-based 索引
      - date_col: 'Date ready for configuration' 列 0-based 索引（找不到为 None）
      - code_columns: [(0-based 列索引, itemcode, desc), ...]
    """
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
        desc = str(desc).strip() if desc is not None else ""
        code_columns.append((i, code, desc))

    data_rows = rows[DATA_START_ROW - 1:]
    return data_rows, du_col, date_col, code_columns


def collect_data(data_rows, du_col, date_col, code_columns, du_set=None):
    """先按 DU 清单筛选目标行（同一 DU 多条时取 Date ready 最新的那条），再转置。

    返回 (result, seen)：
      - result: [(duid, date, items), ...]，items = [(code, desc, qty), ...]
      - seen: 模板里实际出现过的 DUID（归一化）集合
    """
    best = {}   # norm_duid -> (duid, raw_date, row)
    seen = set()

    # 第一遍：只取 DU_ID 列与日期列，筛出目标行；重复 DU 取日期最新
    for row in data_rows:
        if du_col is None or du_col >= len(row):
            continue
        duid = row[du_col]
        if duid is None or str(duid).strip() == "":
            continue
        duid = str(duid).strip()
        norm = _norm_duid(duid)
        seen.add(norm)
        # 前置条件筛选：提供了 DU 清单时，只转置清单内的 DU ID
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

    # 第二遍：只对筛选出的行做物料转置
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


# ============================================================
# 转置后替换处理（itemcode 替换 + 数量按比例折算）
# ============================================================
# 规则：源 itemcode -> 目标 itemcode，数量比例（比例分母）。
#   新数量 = 原数量 / 比例分母（2:1 即 2 个源 = 1 个目标 -> 除以 2）。
CODE_REPLACE_RULES = [
    ("30070011", "02237428", 1, "1:1"),   # 1:1 替换
    ("28010521", "02237428", 2, "2:1"),   # 2:1 替换（2 个 28010521 = 1 个 02237428）
    ("04070012", "04023275", 1, "1:1"),   # 1:1 替换
]


def _norm_code(s):
    """itemcode 归一化：字符串化去空白；纯数字去前导 0（容错 Excel 数字/文本）。"""
    s = str(s).strip()
    if s.isdigit():
        s = s.lstrip("0") or "0"
    return s


def apply_replacements(data):
    """转置完成后，对每个 DU 的物料做 itemcode 替换 + 数量比例折算。

    输入 data 形如 [(duid, date, [(code, desc, qty), ...]), ...]；
    返回 [(duid, date, [(code, desc, qty, 是否替换code, 替换前code, 替换前数量, 替换比例), ...]), ...]。
    未触发替换的物料，替换相关字段留空。
    """
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
                # 1:1 保持原数量；比例 >1 时向上取整（宁多勿少，保证物料整数够用）
                if divisor == 1:
                    new_qty = qty
                else:
                    new_qty = math.ceil(qty / divisor)
                new_items.append((target, desc, new_qty, "是", code, qty, ratio_str))
            else:
                new_items.append((code, desc, qty, "", "", "", ""))
        result.append((duid, date, new_items))
    return result


# ============================================================
# 类型判定（映射 + 替换 + 发包）
# ============================================================
def resolve_type(code, desc, type_map):
    """返回 (最终type, 是否替换类型, 替换前类型, 是否发包)。"""
    orig = lookup_type(type_map, code)

    final = orig
    replaced = False

    # 规则 a：code 前 2 位为 34 -> Type 3
    if code[:2] == "34":
        final = "Type 3"
        replaced = True
    # 规则 b：code 前 1 位为 5 -> Type 1
    elif code[:1] == "5":
        final = "Type 1"
        replaced = True
    # 规则 c：描述含 breaker / AMB -> Type 1
    elif has_any_keyword(desc, ["breaker", "AMB"]):
        final = "Type 1"
        replaced = True

    is_replace = "是" if replaced else ""
    replace_before = orig if replaced else ""

    # 规则 c：最终 type == Type 1 且描述含 RRU/AAU/BBU -> 发包
    is_send = ""
    if final == "Type 1" and has_any_keyword(desc, ["RRU", "AAU", "BBU"]):
        is_send = "发包"

    return final, is_replace, replace_before, is_send


# ============================================================
# 输出（带样式）
# ============================================================
HEADER_FILL = PatternFill("solid", fgColor="305496")
HEADER_FONT = Font(bold=True, color="FFFFFF", name="微软雅黑", size=11)
DATA_FONT = Font(name="微软雅黑", size=11)
CENTER = Alignment(horizontal="center", vertical="center", wrap_text=True)
LEFT = Alignment(horizontal="left", vertical="center", wrap_text=True)
THIN = Side(style="thin", color="BFBFBF")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)


def build_output(data, type_map, template_path, out_path):
    twb = openpyxl.load_workbook(template_path)
    tws = twb.active
    headers = [str(c.value).strip() if c.value is not None else "" for c in tws[1]]
    # 输出增加「Date ready for configuration」列（若模板已有同名列则不重复追加）
    if not any(_norm(h) == "datereadyforconfiguration" for h in headers):
        headers.append("Date ready for configuration")
    ncol = len(headers)

    def col_of(*names):
        for i, h in enumerate(headers):
            if _norm(h) in [_norm(n) for n in names]:
                return i
        return None

    idx_duid = col_of("DUID", "DU ID")
    idx_desc = col_of("描述", "desc")
    idx_code = col_of("itemcode", "item code")
    idx_qty = col_of("qty", "数量")
    idx_type = col_of("type", "设备类型")
    idx_send = col_of("是否发包")
    idx_replace = col_of("是否替换类型")
    idx_replace_before = col_of("替换前类型")
    idx_rcode = col_of("是否替换code")
    idx_rcode_before = col_of("替换前code")
    idx_rqty = col_of("替换前数量")
    idx_rratio = col_of("替换比例")
    idx_date = col_of("Date ready for configuration")

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "物料转置"

    for i, h in enumerate(headers, start=1):
        c = ws.cell(row=1, column=i, value=h)
        c.fill = HEADER_FILL
        c.font = HEADER_FONT
        c.alignment = CENTER
        c.border = BORDER
    ws.freeze_panes = "A2"

    row_idx = 2
    unmapped = set()
    for duid, date, items in data:
        for code, desc, qty, is_rcode, src_code, src_qty, ratio in items:
            final, is_replace, replace_before, is_send = resolve_type(code, desc, type_map)
            if final == "":
                unmapped.add(code)
            row_map = {
                idx_duid: duid,
                idx_desc: desc,
                idx_code: code,
                idx_qty: qty,
                idx_type: final,
                idx_send: is_send,
                idx_replace: is_replace,
                idx_replace_before: replace_before,
                idx_rcode: is_rcode,
                idx_rcode_before: src_code,
                idx_rqty: src_qty,
                idx_rratio: ratio,
                idx_date: _format_date(date),
            }
            for col, val in row_map.items():
                if col is None:
                    continue
                c = ws.cell(row=row_idx, column=col + 1, value=val)
                c.font = DATA_FONT
                c.border = BORDER
                c.alignment = LEFT if col == idx_desc else CENTER
            row_idx += 1

    widths = {"DUID": 26, "描述": 55, "itemcode": 18, "qty": 8, "type": 10,
              "是否发包": 10, "是否替换类型": 12, "替换前类型": 12,
              "是否替换code": 12, "替换前code": 18, "替换前数量": 12, "替换比例": 10,
              "Date ready for configuration": 22}
    for i, h in enumerate(headers, start=1):
        w = 12
        for k, v in widths.items():
            if _norm(h) == _norm(k):
                w = v
                break
        ws.column_dimensions[openpyxl.utils.get_column_letter(i)].width = w

    wb.save(out_path)
    return unmapped


def _configure_stdout():
    if hasattr(sys.stdout, "reconfigure"):
        try:
            sys.stdout.reconfigure(encoding="utf-8", errors="replace")
            sys.stderr.reconfigure(encoding="utf-8", errors="replace")
        except Exception:
            pass


def parse_args():
    p = argparse.ArgumentParser(description="MRDN 物料转置工具")
    p.add_argument("--input_template", default=None, help="输入模板.xlsx 路径")
    p.add_argument("--output_template", default=None, help="输出模板.xlsx 路径")
    p.add_argument("--type_table", default=None, help="物料类型表.xlsx 路径")
    p.add_argument("--du_list", default=None, help="DU清单（前置条件 DUID 列表）路径")
    p.add_argument("--duid", default=None, help="单 DUID（只转置这一个 DUID，优先级高于 --du_list）")
    p.add_argument("--out_dir", default=None, help="输出目录（不传则输出到脚本同目录）")
    return p.parse_args()


def main():
    _configure_stdout()
    args = parse_args()
    input_template = args.input_template or INPUT_TEMPLATE
    output_template = args.output_template or OUTPUT_TEMPLATE
    type_table = args.type_table or TYPE_TABLE
    du_list = args.du_list
    duid = args.duid
    out_dir = args.out_dir or BASE_DIR
    os.makedirs(out_dir, exist_ok=True)
    output_file = os.path.join(out_dir, "输出结果.xlsx")

    for p, name in ((input_template, "输入模板"), (output_template, "输出模板"),
                    (type_table, "物料类型表")):
        if not os.path.exists(p):
            print("[错误] 找不到%s：%s" % (name, p))
            sys.exit(1)

    data_rows, du_col, date_col, code_columns = read_input(input_template)
    print("itemcode 候选列数：", len(code_columns))
    print("DU ID 列：", openpyxl.utils.get_column_letter(du_col + 1))
    print("Date ready for configuration 列：",
          openpyxl.utils.get_column_letter(date_col + 1) if date_col is not None else "未找到")

    type_map = read_type_map(type_table)
    print("物料类型表编码数：", len(type_map))

    # 前置条件：单 DUID（--duid）优先；否则读 DU 清单（--du_list）
    if duid:
        du_map = {_norm_duid(str(duid).strip()): str(duid).strip()}
        du_set = set(du_map.keys())
        print("单 DUID 模式，只转置：", str(duid).strip())
    else:
        du_map = read_du_list(du_list)
        du_set = set(du_map.keys())
        if du_set:
            print("DU 清单（前置条件）DUID 数：", len(du_set))
        else:
            print("未提供 DU 清单（或为空），将转置全部 DU")

    data, template_duids = collect_data(data_rows, du_col, date_col, code_columns, du_set)

    # 转置后替换处理（itemcode 替换 + 数量比例折算）
    data = apply_replacements(data)

    # 校验：DU 清单里有、但输入模板里找不到的 DUID → 报错退出
    if du_set:
        missing = sorted(du_map[n] for n in du_set if n not in template_duids)
        if missing:
            print("[错误] DU 清单中以下 DUID 在输入模板中找不到，请核对清单：")
            for m in missing:
                print("   -", m)
            sys.exit(1)
    for duid, date, items in data:
        print("\nDU:", duid, "（物料", len(items), "种） Date ready:", _format_date(date))
        for code, desc, qty, is_rcode, src_code, src_qty, ratio in items:
            final, is_replace, replace_before, is_send = resolve_type(code, desc, type_map)
            print("   ", code, "=", qty, "type=", final,
                  ("替换(" + replace_before + ")" if is_replace else ""),
                  ("发包" if is_send else ""),
                  ("【code替换:" + str(src_code) + "→" + code + " " + ratio + "】" if is_rcode else ""),
                  "|", desc[:40])

    unmapped = build_output(data, type_map, output_template, output_file)
    if unmapped:
        print("\n[警告] 以下 itemcode 未匹配到 type（已留空）：", sorted(unmapped))
    print("\n已生成输出文件：", output_file)


if __name__ == "__main__":
    try:
        main()
    except SystemExit:
        raise
    except Exception as e:
        _configure_stdout()
        print("[错误] 运行失败：%s" % e)
        sys.exit(1)
