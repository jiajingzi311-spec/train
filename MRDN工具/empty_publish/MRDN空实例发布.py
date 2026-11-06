# -*- coding: utf-8 -*-
"""
MRDN 空实例发布（empty_publish）
=================================
功能：
    从「物料制作清单」中找出需要做空实例的 DUID（该 DUID 下没有任何「发包」标签、
    但有 Type 2 / Type 3 的辅料信息），结合「空实例列表」里的合同号，按 SmartMR
    模板生成 SmartMR 空实例推送清单。

处理规则：
    1. 按 DUID 分组物料制作清单，找出「无发包标签 + 有 Type 2/Type 3」的 DUID（空实例）。
    2. 对每个空实例 DUID，取其 Type 2 / Type 3 的物料行：
       - DUID           → SmartMR 模板的 *DU ID
       - itemcode       → SmartMR 模板的 *Item Code
       - qty            → SmartMR 模板的 *Required Qty
    3. 空实例列表的 HUAWEI CONTRACT NO. *（按 DUID 匹配；列表无 DUID 列则用唯一合同号）
       → SmartMR 模板的 Demand Contract No.
    4. SmartMR 模板的 *Warehouse Type 固定填 "Central"
    5. 若找不到任何空实例 DUID，在结果表 A1（第一行第一列）输出「本批次DUID无需使用空实例」。

依赖：
    pip install openpyxl
"""


# ============================================================================
#  ⚙ MRDN 系统集成参数对照表（配置系统「工具输入关联」时，照此填「参数名」）
# ============================================================================
#   参数名（填到系统）       对应文件 / 含义
#   --material_list          物料制作清单（物料转置表，含 DUID/itemcode/qty/type/是否发包 列）
#   --empty_list             空实例列表（含 HUAWEI CONTRACT NO. * 列，可选 DU ID * 列）
#   --smartmr_template       SmartMR 模板（Du Material 表）
#   （--out_dir 为系统自动传入的输出目录，无需在系统里配置）
# ============================================================================

import os
import sys
import argparse
from collections import OrderedDict

import openpyxl


if getattr(sys, 'frozen', False):
    BASE_DIR = os.path.dirname(os.path.abspath(sys.executable))
else:
    BASE_DIR = os.path.dirname(os.path.abspath(__file__))

OUTPUT_FILENAME = "输出结果.xlsx"
NO_EMPTY_MSG = "本批次DUID无需使用空实例"


def _configure_stdout():
    if hasattr(sys.stdout, "reconfigure"):
        try:
            sys.stdout.reconfigure(encoding="utf-8", errors="replace")
            sys.stderr.reconfigure(encoding="utf-8", errors="replace")
        except Exception:
            pass


def _norm(s):
    if s is None:
        return ""
    return "".join(ch for ch in str(s).lower() if ch.isalnum())


def _s(v):
    if v is None:
        return ""
    return str(v).strip()


def _headers(ws):
    return [_s(c.value) for c in ws[1]]


def _find_col(headers, *names):
    targets = {_norm(n) for n in names}
    for i, h in enumerate(headers):
        if _norm(h) in targets:
            return i
    return -1


def parse_args():
    p = argparse.ArgumentParser(description="MRDN 空实例发布")
    p.add_argument("--material_list", default=None, help="物料制作清单 路径")
    p.add_argument("--empty_list", default=None, help="空实例列表 路径")
    p.add_argument("--smartmr_template", default=None, help="SmartMR 模板 路径")
    p.add_argument("--out_dir", default=None, help="输出目录（系统自动传入）")
    return p.parse_args()


def read_material(path):
    """读物料制作清单，按 DUID 分组，返回 OrderedDict {DUID: [(item, qty, type_norm, is_fb)]}。"""
    wb = openpyxl.load_workbook(path, data_only=True)
    ws = wb.active
    headers = _headers(ws)
    ci_du = _find_col(headers, "DUID", "DU ID", "DU")
    ci_item = _find_col(headers, "itemcode", "Item Code", "item", "material code")
    ci_qty = _find_col(headers, "qty", "quantity")
    ci_type = _find_col(headers, "type")
    ci_fb = _find_col(headers, "是否发包", "发包")
    if ci_du < 0 or ci_item < 0 or ci_qty < 0 or ci_type < 0 or ci_fb < 0:
        print("[错误] 物料制作清单缺少必要列（需要 DUID / itemcode / qty / type / 是否发包），实际表头：%s" % headers)
        sys.exit(1)

    groups = OrderedDict()
    for r in ws.iter_rows(min_row=2, values_only=True):
        du = _s(r[ci_du])
        if not du:
            continue
        item = _s(r[ci_item])
        qty = r[ci_qty]
        t = _norm(r[ci_type])
        fb = _s(r[ci_fb])  # 非空即视为有发包标签
        groups.setdefault(du, []).append((item, qty, t, fb))
    return groups


def find_empty_duids(groups):
    """找出空实例 DUID：无「发包」标签 + 有 Type 2/Type 3 信息。返回按首次出现顺序的 DUID 列表。"""
    empty = []
    for du, rows in groups.items():
        has_fb = any(r[3] != "" for r in rows)
        has_type23 = any(r[2] in ("type2", "type3") for r in rows)
        if (not has_fb) and has_type23:
            empty.append(du)
    return empty


def read_empty_list(path):
    """读空实例列表，返回 (mapping, fallback, has_du_col)。
    mapping: DU ID → HUAWEI CONTRACT NO.（列表含 DU ID 列时）
    fallback: 第一个非空合同号（列表无 DU ID 列时，所有 DU 共用）
    """
    wb = openpyxl.load_workbook(path, data_only=True)
    ws = wb.active
    headers = _headers(ws)
    ci_contract = _find_col(headers, "HUAWEI CONTRACT NO. *", "HUAWEI CONTRACT NO", "Contract No", "contract")
    ci_du = _find_col(headers, "DU ID *", "DU ID", "DUID", "DU")
    if ci_contract < 0:
        print("[错误] 空实例列表缺少必要列（需要 HUAWEI CONTRACT NO. *），实际表头：%s" % headers)
        sys.exit(1)

    mapping = {}
    fallback = ""
    has_du_col = ci_du >= 0
    for r in ws.iter_rows(min_row=2, values_only=True):
        contract = _s(r[ci_contract])
        if not contract:
            continue
        if not fallback:
            fallback = contract
        if has_du_col:
            du = _s(r[ci_du])
            if du and du not in mapping:
                mapping[du] = contract  # 第一个 wins
    return mapping, fallback, has_du_col


def find_main_sheet(wb):
    for sn in wb.sheetnames:
        ws = wb[sn]
        if ws.max_row >= 1:
            if _find_col(_headers(ws), "*DU ID", "DU ID", "DUID") >= 0:
                return ws
    if "Du Material" in wb.sheetnames:
        return wb["Du Material"]
    return wb.active


def build_output(empty_rows, mapping, fallback, has_du_col, template_path, out_path):
    """以 SmartMR 模板为底，写入空实例记录。无空实例时 A1 写提示。返回 (写入行数, 缺合同号DUID列表)。"""
    wb = openpyxl.load_workbook(template_path)
    ws = find_main_sheet(wb)
    headers = _headers(ws)
    ci_du = _find_col(headers, "*DU ID", "DU ID", "DUID")
    ci_item = _find_col(headers, "*Item Code", "Item Code")
    ci_qty = _find_col(headers, "*Required Qty", "Required Qty")
    ci_wh = _find_col(headers, "*Warehouse Type", "Warehouse Type")
    ci_contract = _find_col(headers, "Demand Contract No.", "Demand Contract")
    if ci_du < 0 or ci_item < 0 or ci_qty < 0 or ci_wh < 0 or ci_contract < 0:
        print("[错误] SmartMR 模板缺少必要列（需要 *DU ID / *Item Code / *Required Qty / *Warehouse Type / Demand Contract No.），实际表头：%s" % headers)
        sys.exit(1)

    # 清空第 2 行及以后的示例/旧数据（保留表头）
    for row in ws.iter_rows(min_row=2, max_row=ws.max_row, max_col=ws.max_column):
        for cell in row:
            cell.value = None

    # 无空实例 DUID：A1 输出提示文字
    if not empty_rows:
        ws.cell(row=1, column=1, value=NO_EMPTY_MSG)
        wb.save(out_path)
        return 0, []

    missing = []
    for idx, (du, item, qty) in enumerate(empty_rows):
        r = idx + 2
        ws.cell(row=r, column=ci_du + 1, value=du)
        ws.cell(row=r, column=ci_item + 1, value=item)
        ws.cell(row=r, column=ci_qty + 1, value=qty)
        ws.cell(row=r, column=ci_wh + 1, value="Central")

        if has_du_col:
            contract = mapping.get(du, "")
        else:
            contract = fallback
        ws.cell(row=r, column=ci_contract + 1, value=contract if contract else None)
        if not contract:
            missing.append(du)

    wb.save(out_path)
    return len(empty_rows), missing


def main():
    _configure_stdout()
    args = parse_args()

    out_dir = args.out_dir or BASE_DIR
    os.makedirs(out_dir, exist_ok=True)
    out_path = os.path.join(out_dir, OUTPUT_FILENAME)

    for p, name in ((args.material_list, "物料制作清单"),
                    (args.empty_list, "空实例列表"),
                    (args.smartmr_template, "SmartMR 模板")):
        if not p or not os.path.exists(p):
            print("[错误] 找不到%s：%s" % (name, p))
            sys.exit(1)

    groups = read_material(args.material_list)
    empty_duids = find_empty_duids(groups)
    mapping, fallback, has_du_col = read_empty_list(args.empty_list)

    print("物料制作清单中 DUID 总数：%d" % len(groups))
    print("需要做空实例的 DUID 数：%d" % len(empty_duids))
    if empty_duids:
        print("空实例 DUID：%s" % ", ".join(empty_duids))

    # 组装空实例输出行：空实例 DUID 下的 Type 2 / Type 3 行
    empty_rows = []
    for du in empty_duids:
        for item, qty, t, fb in groups[du]:
            if t in ("type2", "type3"):
                empty_rows.append((du, item, qty))

    print("空实例辅料记录数（Type 2 / Type 3）：%d" % len(empty_rows))

    written, missing = build_output(empty_rows, mapping, fallback, has_du_col, args.smartmr_template, out_path)

    if not empty_rows:
        print("未找到空实例 DUID，结果表 A1 已输出提示：%s" % NO_EMPTY_MSG)
    else:
        print("已写入 SmartMR 记录数：%d" % written)
        if missing:
            seen = []
            for du in missing:
                if du not in seen:
                    seen.append(du)
            print("[警告] 以下 %d 个空实例 DUID 在空实例列表里未找到合同号，Demand Contract No. 留空：%s" % (len(seen), ", ".join(seen)))
        else:
            print("所有空实例 DUID 均已匹配到合同号。")

    print("\n已生成输出文件：", out_path)


if __name__ == "__main__":
    try:
        main()
    except SystemExit:
        raise
    except Exception as e:
        _configure_stdout()
        print("[错误] 运行失败：%s" % e)
        sys.exit(1)
