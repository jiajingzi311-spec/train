# -*- coding: utf-8 -*-
"""
MRDN 辅料发布（aux_publish）
=================================
功能：
    根据「物料制作清单」筛选出辅料（type 为 Type 2 / Type 3）的记录，
    结合「辅料包信息」里的合同号，按 SmartMR 模板生成 SmartMR 推送清单。

处理规则：
    1. 从「物料制作清单」只取 type 为 "Type 2" 和 "Type 3" 的行。
    2. DUID           → SmartMR 模板的 *DU ID
    3. itemcode       → SmartMR 模板的 *Item Code
    4. qty            → SmartMR 模板的 *Required Qty
    5. 辅料包信息的 HUAWEI CONTRACT NO. *（按 DUID 匹配）→ SmartMR 模板的 Demand Contract No.
    6. SmartMR 模板的 *Warehouse Type 固定填 "Central"
    7. 输出结果文件（SmartMR 模板为底，填好数据）

依赖：
    pip install openpyxl
"""


# ============================================================================
#  ⚙ MRDN 系统集成参数对照表（配置系统「工具输入关联」时，照此填「参数名」）
# ============================================================================
#   参数名（填到系统）       对应文件 / 含义
#   --material_list          物料制作清单（物料转置表，含 DUID/itemcode/qty/type 列）
#   --smartmr_template       SmartMR 模板（Du Material 表，含 *DU ID/*Item Code/*Required Qty/*Warehouse Type/Demand Contract No. 列）
#   --aux_pack               辅料包信息（Spart Configuration 表，含 HUAWEI CONTRACT NO. * / DU ID * 列）
#   （--out_dir 为系统自动传入的输出目录，无需在系统里配置）
# ============================================================================

import os
import sys
import argparse

import openpyxl


# 基础路径：打包成 exe 后，脚本同目录 = exe 所在目录
if getattr(sys, 'frozen', False):
    BASE_DIR = os.path.dirname(os.path.abspath(sys.executable))
else:
    BASE_DIR = os.path.dirname(os.path.abspath(__file__))

OUTPUT_FILENAME = "输出结果.xlsx"


def _configure_stdout():
    """Windows 打包后默认 GBK，中文 print 会崩；统一转 UTF-8。"""
    if hasattr(sys.stdout, "reconfigure"):
        try:
            sys.stdout.reconfigure(encoding="utf-8", errors="replace")
            sys.stderr.reconfigure(encoding="utf-8", errors="replace")
        except Exception:
            pass


def _norm(s):
    """归一化：去非字母数字字符、转小写，用于列名/取值匹配。"""
    if s is None:
        return ""
    return "".join(ch for ch in str(s).lower() if ch.isalnum())


def _s(v):
    """单元格值转干净字符串（None → 空串）。"""
    if v is None:
        return ""
    return str(v).strip()


def _headers(ws):
    """读取某 sheet 的第一行表头（字符串列表）。"""
    return [_s(c.value) for c in ws[1]]


def _find_col(headers, *names):
    """在表头里找匹配的列（返回 0-based 列号），找不到返回 -1。"""
    targets = {_norm(n) for n in names}
    for i, h in enumerate(headers):
        if _norm(h) in targets:
            return i
    return -1


def parse_args():
    p = argparse.ArgumentParser(description="MRDN 辅料发布")
    p.add_argument("--material_list", default=None, help="物料制作清单 路径")
    p.add_argument("--smartmr_template", default=None, help="SmartMR 模板 路径")
    p.add_argument("--aux_pack", default=None, help="辅料包信息 路径")
    p.add_argument("--out_dir", default=None, help="输出目录（系统自动传入）")
    return p.parse_args()


def read_material(path):
    """读物料制作清单，筛选 type ∈ {Type 2, Type 3}，返回 [(DUID, itemcode, qty), ...]。"""
    wb = openpyxl.load_workbook(path, data_only=True)
    ws = wb.active
    headers = _headers(ws)
    ci_du = _find_col(headers, "DUID", "DU ID", "DU")
    ci_item = _find_col(headers, "itemcode", "Item Code", "item", "material code")
    ci_qty = _find_col(headers, "qty", "quantity")
    ci_type = _find_col(headers, "type")
    if ci_du < 0 or ci_item < 0 or ci_qty < 0 or ci_type < 0:
        print("[错误] 物料制作清单缺少必要列（需要 DUID / itemcode / qty / type），实际表头：%s" % headers)
        sys.exit(1)

    rows = []
    for r in ws.iter_rows(min_row=2, values_only=True):
        t = _norm(r[ci_type])
        if t in ("type2", "type3"):
            du = _s(r[ci_du])
            item = _s(r[ci_item])
            qty = r[ci_qty]
            if du:
                rows.append((du, item, qty))
    return rows


def read_aux_pack(path):
    """读辅料包信息，建 DU ID → HUAWEI CONTRACT NO. 映射，返回 (mapping, conflicts)。"""
    wb = openpyxl.load_workbook(path, data_only=True)
    ws = wb.active
    headers = _headers(ws)
    ci_contract = _find_col(headers, "HUAWEI CONTRACT NO. *", "HUAWEI CONTRACT NO", "Contract No", "contract")
    ci_du = _find_col(headers, "DU ID *", "DU ID", "DUID", "DU")
    if ci_contract < 0 or ci_du < 0:
        print("[错误] 辅料包信息缺少必要列（需要 HUAWEI CONTRACT NO. * / DU ID *），实际表头：%s" % headers)
        sys.exit(1)

    mapping = {}
    conflicts = []
    for r in ws.iter_rows(min_row=2, values_only=True):
        du = _s(r[ci_du])
        contract = _s(r[ci_contract])
        if not du:
            continue
        if du in mapping and mapping[du] != contract:
            conflicts.append((du, mapping[du], contract))
        else:
            mapping[du] = contract
    return mapping, conflicts


def find_main_sheet(wb):
    """找到 SmartMR 模板的主数据表（含 *DU ID 表头的 sheet）。"""
    for sn in wb.sheetnames:
        ws = wb[sn]
        if ws.max_row >= 1:
            if _find_col(_headers(ws), "*DU ID", "DU ID", "DUID") >= 0:
                return ws
    # 回退：优先 "Du Material"，再 active
    if "Du Material" in wb.sheetnames:
        return wb["Du Material"]
    return wb.active


def build_output(material_rows, aux_mapping, template_path, out_path):
    """以 SmartMR 模板为底，清空示例数据，填入辅料记录，返回 (写入行数, 缺合同号DUID列表)。"""
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

    missing = []
    for idx, (du, item, qty) in enumerate(material_rows):
        r = idx + 2  # 从第 2 行开始
        ws.cell(row=r, column=ci_du + 1, value=du)
        ws.cell(row=r, column=ci_item + 1, value=item)
        ws.cell(row=r, column=ci_qty + 1, value=qty)
        ws.cell(row=r, column=ci_wh + 1, value="Central")
        contract = aux_mapping.get(du, "")
        ws.cell(row=r, column=ci_contract + 1, value=contract if contract else None)
        if not contract:
            missing.append(du)

    wb.save(out_path)
    return len(material_rows), missing


def main():
    _configure_stdout()
    args = parse_args()

    out_dir = args.out_dir or BASE_DIR
    os.makedirs(out_dir, exist_ok=True)
    out_path = os.path.join(out_dir, OUTPUT_FILENAME)

    # 校验输入
    for p, name in ((args.material_list, "物料制作清单"),
                    (args.smartmr_template, "SmartMR 模板"),
                    (args.aux_pack, "辅料包信息")):
        if not p or not os.path.exists(p):
            print("[错误] 找不到%s：%s" % (name, p))
            sys.exit(1)

    material_rows = read_material(args.material_list)
    aux_mapping, conflicts = read_aux_pack(args.aux_pack)

    print("物料制作清单中 Type 2 / Type 3 记录数：%d" % len(material_rows))
    print("辅料包信息中 DUID 映射数：%d" % len(aux_mapping))

    for du, old_c, new_c in conflicts:
        print("[警告] DUID %s 在辅料包信息里对应多个合同号（%s / %s），取第一个 %s" % (du, old_c, new_c, old_c))

    written, missing = build_output(material_rows, aux_mapping, args.smartmr_template, out_path)

    print("已写入 SmartMR 记录数：%d" % written)
    if missing:
        seen = []
        for du in missing:
            if du not in seen:
                seen.append(du)
        print("[警告] 以下 %d 个 DUID 在辅料包信息里未找到合同号，Demand Contract No. 留空：%s" % (len(seen), ", ".join(seen)))
    else:
        print("所有 DUID 均已匹配到合同号，Demand Contract No. 与辅料包信息一致。")

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
