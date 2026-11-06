# -*- coding: utf-8 -*-
"""
MRDN 发包配置工具（fitting_bag）
================================

功能一（发包配置）：
    从「输入模板.xlsx」中找出「是否发包」为「发包」的 DUID 及其 qty，
    按照「输出模板.xlsx」的 Spart Configuration 表头，生成「输出结果.xlsx」。
    其中 HUAWEI CONTRACT NO. / FINANCIAL PRODUCT CODE / SPART NUMBER 三列
    不再用常数，而是取「合同余量查询」筛选出的余量最大记录动态填入。

填写规则（按 Spart Configuration 表头列定位）：
    HUAWEI CONTRACT NO. *    = 余量最大记录的 contractNumber
    Product Type *           = 常数 "Equipment"
    PU NO. *                 = 常数 "001"
    DU ID *                  = 输入模板中「是否发包」为「发包」的 DUID
    FINANCIAL PRODUCT CODE * = 余量最大记录的 financialCode
    SPART NUMBER *           = 余量最大记录的 spartNumber
    SPART QTY. *             = 该 DUID 对应的 qty
    其余列留空。
    若余量最大记录的 remainQty 不足以覆盖所有「发包」行的 qty 之和，则报错退出。

功能二（合同余量查询）：
    读「合同表.xlsx」，筛选 Project Code = "56A0R6K" 且 Sub 列 ∈
    {EOS, EXP, FTK, NSB, Small Cell} 的记录，取所有 Contract No.，
    每 100 个一批调用 CPQ「余量查询」接口 getContractRemainInfo
    （pageSize=150，每批间隔 1~5 秒控制访问频率）。
    返回数据再筛选：contractNumber 结尾非 "AG"/"AX"、spartNumber 开头 "PTAM"，
    保存为「余量查询结果.xlsx」（按 remainQty 降序）。
    Cookie 来自 MRDN 系统「平台网页配置」抓取的平台凭证（CPQ），
    由系统自动生成 JSON 文件经 --platform_cfg 传入。

输出：
    功能一：同目录下「输出结果.xlsx」（保留输出模板的 sheet、表头与格式）。
    功能二：同目录下「余量查询结果.xlsx」。

依赖：openpyxl, requests
"""


# ============================================================================
#  ⚙ MRDN 系统集成参数对照表（配置系统「工具输入关联」时，照此填「参数名」）
# ============================================================================
#   参数名（填到系统）       对应文件 / 含义
#   --input_template         输入模板.xlsx（含「是否发包」列的物料表）
#   --output_template        输出模板.xlsx（Spart Configuration 表头）
#   --contract_table         合同表.xlsx（含 Project Code / Sub / Contract No.）
#   --platform_cfg           平台凭证（关联到系统「平台凭证」下拉选 CPQ，系统自动生成 JSON）
#   （--out_dir 为系统自动传入的输出目录，无需在系统里配置）
# ============================================================================

import os
import re
import sys
import json
import time
import argparse

import openpyxl
import requests
from openpyxl.styles import Font, Alignment, Border, Side


if getattr(sys, 'frozen', False):
    BASE_DIR = os.path.dirname(os.path.abspath(sys.executable))
else:
    BASE_DIR = os.path.dirname(os.path.abspath(__file__))
INPUT_TEMPLATE = os.path.join(BASE_DIR, "输入模板.xlsx")
OUTPUT_TEMPLATE = os.path.join(BASE_DIR, "输出模板.xlsx")
OUTPUT_FILE = os.path.join(BASE_DIR, "输出结果.xlsx")

SHEET_TARGET = "Spart Configuration"   # 输出模板中要填写数据的 sheet

# 常数映射：目标列（归一化名） -> 常数
# 注意：HUAWEI CONTRACT NO. / FINANCIAL PRODUCT CODE / SPART NUMBER 三列
#       不再用常数，而是由「合同余量查询」筛选出的余量最大记录动态填入。
CONSTANTS = {
    "producttype": "Equipment",
    "puno": "001",
}

# 余量记录字段名（接口返回的 key，可能略有大小写差异，_get_field 会容错匹配）
FIELD_CONTRACT = ("contractNumber", "contractNo", "contract")
FIELD_FINANCIAL = ("financialCode", "financialProductCode", "financial")
FIELD_SPART = ("spartNumber", "spartNo", "spart")
FIELD_REMAIN = ("remainQty", "remainQuantity", "remain")


# ============================================================================
#  合同余量查询（功能二）相关常量
# ============================================================================
# CPQ 余量查询接口（固定地址，POST，写死在工具里；Cookie 来自平台凭证 JSON）
CPQ_REMAIN_URL = (
    "https://cpq.huawei.com/unistar/cpq/gateway/CPQBEService/"
    "services/cpqDataRemainCalcService/getContractRemainInfo"
)

PROJECT_CODE_TARGET = "56A0R6K"          # 合同表筛选：Project Code 精确等于该值
SUB_TYPES = ["EOS", "EXP", "FTK", "NSB", "Small Cell"]  # Sub 列允许的取值

CONTRACT_BATCH_SIZE = 100    # 每次给接口传的合同号条数（接口单次上限 100）
PAGE_SIZE = 150              # 每页返回条数（单批 100 个合同号，pageSize 给足 150）

# 查询 body 里的固定 query 参数（除 contractNos 外，其余照 cURL 原样固定）
QUERY_FIXED = {
    "financialCode": "",
    "paymentUnit": "",
    "salesTypes": [1, 3, 4, 5],
    "spartNumbers": [],
    "partDescription": "",
    "model": "",
    "remainQty": [1],
    "isOffset": 0,
    "serviceMaterialFlag": 0,
}


def normalize(s):
    """归一化列名：去除非字母/数字/中文，转小写，便于匹配。"""
    return re.sub(r"[^A-Za-z0-9\u4e00-\u9fff]", "", str(s)).lower()


def find_col(headers, *names):
    """按归一化名在表头中找列索引（0-based），找不到返回 None。"""
    for i, h in enumerate(headers):
        if normalize(h) in [normalize(n) for n in names]:
            return i
    return None


def _match_du(cell_value, target):
    """判断单元格里的 DUID 是否包含目标 DU（支持多个 DUID 用分隔符相连）。"""
    if cell_value is None or not target:
        return False
    tokens = [d.strip() for d in re.split(r"[、,，\s;；]+", str(cell_value)) if d.strip()]
    return (target or "").strip() in tokens


def read_fabag_data(path):
    """从输入模板读取「是否发包=发包」的 (DUID, qty) 列表。"""
    wb = openpyxl.load_workbook(path, data_only=True)
    ws = wb.active
    headers = [str(c.value).strip() if c.value is not None else "" for c in ws[1]]

    col_duid = find_col(headers, "DUID")
    col_qty = find_col(headers, "qty")
    col_send = find_col(headers, "是否发包")
    if col_duid is None or col_qty is None or col_send is None:
        raise RuntimeError("输入模板缺少必需列：DUID / qty / 是否发包")

    result = []
    for r in range(2, ws.max_row + 1):
        send = ws.cell(row=r, column=col_send + 1).value
        if send is None:
            continue
        if str(send).strip() != "发包":
            continue
        duid = ws.cell(row=r, column=col_duid + 1).value
        qty = ws.cell(row=r, column=col_qty + 1).value
        if duid is None or str(duid).strip() == "":
            continue
        result.append((str(duid).strip(), qty))
    return result, headers


# 数据行样式
DATA_FONT = Font(name="微软雅黑", size=11)
CENTER = Alignment(horizontal="center", vertical="center", wrap_text=True)
THIN = Side(style="thin", color="BFBFBF")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)


def build_output(rows, template_path, out_path, best_record=None):
    wb = openpyxl.load_workbook(template_path)

    if SHEET_TARGET not in wb.sheetnames:
        raise RuntimeError("输出模板缺少 sheet：%s" % SHEET_TARGET)
    ws = wb[SHEET_TARGET]

    headers = [str(c.value).strip() if c.value is not None else "" for c in ws[1]]

    col_contract = find_col(headers, "HUAWEI CONTRACT NO. *")
    col_product = find_col(headers, "Product Type *")
    col_puno = find_col(headers, "PU NO. *")
    col_duid = find_col(headers, "DU ID *")
    col_fin = find_col(headers, "FINANCIAL PRODUCT CODE *")
    col_spart = find_col(headers, "SPART NUMBER *")
    col_qty = find_col(headers, "SPART QTY. *")

    # 从「余量最大记录」取三个动态值（替换原常数；无记录则留空）
    contract_val = _cell_str(_get_field(best_record, *FIELD_CONTRACT)) if best_record else ""
    fin_val = _cell_str(_get_field(best_record, *FIELD_FINANCIAL)) if best_record else ""
    spart_val = _cell_str(_get_field(best_record, *FIELD_SPART)) if best_record else ""

    # 清空原有数据行（第 2 行起），保留表头与格式
    if ws.max_row >= 2:
        ws.delete_rows(2, ws.max_row - 1)

    for i, (duid, qty) in enumerate(rows):
        r = 2 + i
        values = {
            col_contract: contract_val,
            col_product: CONSTANTS["producttype"],
            col_puno: CONSTANTS["puno"],
            col_duid: duid,
            col_fin: fin_val,
            col_spart: spart_val,
            col_qty: qty,
        }
        for col, val in values.items():
            if col is None:
                continue
            c = ws.cell(row=r, column=col + 1, value=val)
            c.font = DATA_FONT
            c.alignment = CENTER
            c.border = BORDER

    wb.save(out_path)


# ============================================================================
#  合同余量查询（功能二）
# ============================================================================
def _cell_str(v):
    """单元格值转干净字符串：float 整数去 .0，None/空返回空串。"""
    if v is None:
        return ""
    if isinstance(v, float) and v.is_integer():
        return str(int(v))
    return str(v).strip()


def _norm_sub(s):
    """规范化 Sub 列值：去空白、转大写，用于与枚举比较。"""
    return re.sub(r"\s+", "", str(s)).upper()


SUB_TYPES_NORM = [_norm_sub(s) for s in SUB_TYPES]  # EOS/EXP/FTK/NSB/SMALLCELL


def _sub_hit(sub_value):
    """判断 Sub 列值是否命中枚举（规范化后相等，或包含枚举子串）。"""
    v = _norm_sub(sub_value)
    if not v:
        return False
    for t in SUB_TYPES_NORM:
        if v == t or t in v:
            return True
    return False


def read_contract_nos(path):
    """读合同表，筛选 Project Code=56A0R6K 且 Sub 命中枚举的记录，返回去重后的 Contract No. 列表。"""
    wb = openpyxl.load_workbook(path, data_only=True)
    ws = wb.active
    headers = [str(c.value).strip() if c.value is not None else "" for c in ws[1]]

    col_project = find_col(headers, "Project Code", "ProjectCode", "项目代码", "项目编码", "Project")
    col_sub = find_col(headers, "Sub", "Sub Type", "SubType", "子项目", "类型")
    col_contract = find_col(headers, "Contract No.", "Contract No", "ContractNo",
                            "Contract Number", "合同号", "合同编号")
    missing = []
    if col_project is None:
        missing.append("Project Code")
    if col_sub is None:
        missing.append("Sub")
    if col_contract is None:
        missing.append("Contract No.")
    if missing:
        raise RuntimeError("合同表缺少必需列：%s" % " / ".join(missing))

    result = []
    seen = set()
    for r in range(2, ws.max_row + 1):
        project = ws.cell(row=r, column=col_project + 1).value
        sub = ws.cell(row=r, column=col_sub + 1).value
        contract = ws.cell(row=r, column=col_contract + 1).value

        if _cell_str(project) != PROJECT_CODE_TARGET:
            continue
        if not _sub_hit(sub):
            continue
        cno = _cell_str(contract)
        if not cno:
            continue
        if cno in seen:
            continue
        seen.add(cno)
        result.append(cno)
    return result


def load_platform_cfg(path):
    """读平台凭证 JSON，返回 headers（含 Cookie，并补齐接口必需 header）。"""
    if not path or not os.path.exists(path):
        raise RuntimeError("找不到平台凭证文件：%s" % path)
    with open(path, encoding="utf-8") as f:
        cfg = json.load(f)
    headers = dict(cfg.get("headers") or {})
    if not headers.get("Cookie"):
        cookie = cfg.get("cookie")
        if cookie:
            headers["Cookie"] = cookie
    if not headers.get("Cookie"):
        raise RuntimeError("平台凭证里没有 Cookie，请先在「平台网页配置」里一键登录抓取 CPQ Cookie")
    headers.setdefault("Content-Type", "application/json")
    headers.setdefault("Accept", "application/json, text/plain, */*")
    return headers


def query_remain_info(contract_nos, headers):
    """分批（每批 CONTRACT_BATCH_SIZE 个）POST 查询 CPQ 余量，返回合并后的原始返回列表。"""
    payloads = []
    total = len(contract_nos)
    n_batch = (total + CONTRACT_BATCH_SIZE - 1) // CONTRACT_BATCH_SIZE
    for i in range(0, total, CONTRACT_BATCH_SIZE):
        batch = contract_nos[i:i + CONTRACT_BATCH_SIZE]
        batch_no = i // CONTRACT_BATCH_SIZE + 1
        body = {
            "page": {"curPage": 1, "pageSize": PAGE_SIZE},
            "query": dict(QUERY_FIXED, contractNos=batch),
        }
        print("查询第 %d/%d 批（%d 个合同号）..." % (batch_no, n_batch, len(batch)))
        try:
            resp = requests.post(CPQ_REMAIN_URL, headers=headers, json=body,
                                 verify=False, timeout=60)
        except requests.RequestException as e:
            raise RuntimeError("第 %d 批请求失败：%s" % (batch_no, e))
        if resp.status_code != 200:
            raise RuntimeError("第 %d 批请求失败：HTTP %s" % (batch_no, resp.status_code))
        try:
            data = resp.json()
        except ValueError:
            data = {"_raw_text": resp.text}
        payloads.append(data)
        # 访问频率控制：每次请求间隔 1~5 秒，模拟人工操作速度（最后一批除外）
        if i + CONTRACT_BATCH_SIZE < total:
            time.sleep(2)
    return payloads


def _extract_rows(node, _depth=0):
    """从接口返回里递归找第一个「元素全是 dict 的 list」，作为结果数据行。"""
    if _depth > 10 or node is None:
        return []
    if isinstance(node, list):
        if node and all(isinstance(x, dict) for x in node):
            return node
        for x in node:
            r = _extract_rows(x, _depth + 1)
            if r:
                return r
        return []
    if isinstance(node, dict):
        for v in node.values():
            r = _extract_rows(v, _depth + 1)
            if r:
                return r
    return []


def _get_field(row, *names):
    """容错查找字段：先精确匹配，再归一化匹配（大小写/分隔符不敏感）。"""
    if not isinstance(row, dict):
        return None
    for n in names:
        if n in row:
            return row[n]
    norm_map = {normalize(k): v for k, v in row.items() if isinstance(k, str)}
    for n in names:
        nn = normalize(n)
        if nn in norm_map:
            return norm_map[nn]
    return None


def _to_float(v):
    """安全转 float（余量/数量可能是数字或字符串）。"""
    if v is None:
        return 0.0
    if isinstance(v, bool):
        return 1.0 if v else 0.0
    if isinstance(v, (int, float)):
        return float(v)
    try:
        return float(str(v).strip())
    except (ValueError, TypeError):
        return 0.0


def filter_remain_rows(rows):
    """筛选余量记录：contractNumber 结尾非 AG/AX，spartNumber 开头 PTAM。"""
    result = []
    for row in rows:
        contract = _cell_str(_get_field(row, *FIELD_CONTRACT))
        spart = _cell_str(_get_field(row, *FIELD_SPART))
        if not contract or not spart:
            continue
        cu = contract.upper()
        if cu.endswith("AG") or cu.endswith("AX"):
            continue
        if not spart.upper().startswith("PTAM"):
            continue
        result.append(row)
    return result


def pick_best_record(rows):
    """从筛选后的余量记录里取 remainQty 最大的那条（无则返回 None）。"""
    if not rows:
        return None
    return max(rows, key=lambda r: _to_float(_get_field(r, *FIELD_REMAIN)))


def _cell_value(v):
    """把接口返回的任意值转成 Excel 可写：dict/list 转 JSON 字符串。"""
    if isinstance(v, (dict, list)):
        try:
            return json.dumps(v, ensure_ascii=False)
        except Exception:
            return str(v)
    return v


def save_remain_result(rows, out_dir):
    """保存筛选后的余量结果到 Excel（按 remainQty 降序），返回文件路径。"""
    xlsx_path = os.path.join(out_dir, "余量查询结果.xlsx")

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "余量查询结果"
    if rows:
        rows_sorted = sorted(rows, key=lambda r: _to_float(_get_field(r, *FIELD_REMAIN)), reverse=True)
        headers = []
        for row in rows_sorted:
            for k in row.keys():
                if k not in headers:
                    headers.append(k)
        ws.append(headers)
        for row in rows_sorted:
            ws.append([_cell_value(row.get(h)) for h in headers])
    else:
        ws.append(["没有符合筛选条件的余量记录（合同号非 AG/AX 结尾、物料以 PTAM 开头）"])
    wb.save(xlsx_path)
    return xlsx_path


def _configure_stdout():
    if hasattr(sys.stdout, "reconfigure"):
        try:
            sys.stdout.reconfigure(encoding="utf-8", errors="replace")
            sys.stderr.reconfigure(encoding="utf-8", errors="replace")
        except Exception:
            pass


def parse_args():
    p = argparse.ArgumentParser(description="MRDN 发包配置工具")
    p.add_argument("--input_template", default=None, help="输入模板.xlsx 路径")
    p.add_argument("--output_template", default=None, help="输出模板.xlsx 路径")
    p.add_argument("--contract_table", default=None, help="合同表.xlsx 路径（合同余量查询用）")
    p.add_argument("--platform_cfg", default=None, help="平台凭证 JSON 文件路径（系统自动传入）")
    p.add_argument("--out_dir", default=None, help="输出目录（不传则输出到脚本同目录）")
    p.add_argument("--duid", default=None, help="只处理指定 DUID（单 DU 作业时由系统传入）")
    return p.parse_args()


def main():
    _configure_stdout()
    args = parse_args()
    out_dir = args.out_dir or BASE_DIR
    os.makedirs(out_dir, exist_ok=True)

    # 判断跑哪些功能：
    #  - 显式传了 contract_table → 跑「合同余量查询」
    #  - 显式传了 input/output_template，或啥都没传（退回旧行为）→ 跑「发包配置」
    do_contract = bool(args.contract_table)
    do_fabag = bool(args.input_template or args.output_template) or (not do_contract)

    # 筛选后的余量记录（发包配置用它来替换常数）
    remain_rows = []

    # ===== 功能二：合同余量查询 =====
    if do_contract:
        if not args.platform_cfg:
            print("[错误] 已提供合同表，但缺少平台凭证（--platform_cfg）；"
                  "请在系统「工具输入关联」里把 platform_cfg 关联到 CPQ 平台凭证")
            sys.exit(1)
        if not os.path.exists(args.contract_table):
            print("[错误] 找不到合同表：%s" % args.contract_table)
            sys.exit(1)

        headers = load_platform_cfg(args.platform_cfg)
        contract_nos = read_contract_nos(args.contract_table)
        print("从合同表筛选出 Contract No. %d 条" % len(contract_nos))
        if not contract_nos:
            print("[提示] 没有符合条件的合同号（Project Code=%s 且 Sub∈%s）"
                  % (PROJECT_CODE_TARGET, SUB_TYPES))

        payloads = query_remain_info(contract_nos, headers) if contract_nos else []
        all_rows = []
        for p in payloads:
            all_rows.extend(_extract_rows(p))
        print("接口返回数据行共 %d 条" % len(all_rows))

        remain_rows = filter_remain_rows(all_rows)
        print("筛选后（合同非 AG/AX 结尾、物料 PTAM 开头）可用余量记录 %d 条" % len(remain_rows))

        xlsx_path = save_remain_result(remain_rows, out_dir)
        print("已生成余量查询结果：", xlsx_path)

    # ===== 功能一：发包配置 =====
    if do_fabag:
        input_template = args.input_template or INPUT_TEMPLATE
        output_template = args.output_template or OUTPUT_TEMPLATE
        for p, name in ((input_template, "输入模板"), (output_template, "输出模板")):
            if not os.path.exists(p):
                print("[错误] 找不到%s：%s" % (name, p))
                sys.exit(1)

        output_file = os.path.join(out_dir, "输出结果.xlsx")
        rows, headers = read_fabag_data(input_template)
        if args.duid and (args.duid or "").strip():
            duid_filter = (args.duid or "").strip()
            rows = [(d, q) for (d, q) in rows if _match_du(d, duid_filter)]
            print("已按单 DUID 过滤：", duid_filter)
        print("找到「发包」记录 %d 条：" % len(rows))
        for duid, qty in rows:
            print("   DU ID =", duid, "| qty =", qty)

        if not rows:
            print("[提示] 没有「发包」记录，仍会生成只有表头的输出文件。")

        # 用「余量最大」的记录替换原常数（合同/财务产品/物料）
        total_qty = sum(_to_float(q) for _, q in rows)
        best_record = pick_best_record(remain_rows) if remain_rows else None
        if best_record is None:
            print("[错误] 没有可用的余量记录，无法确定合同/财务产品/物料；"
                  "请先提供合同表并执行余量查询（合同号非 AG/AX 结尾、物料 PTAM 开头）")
            sys.exit(1)
        best_remain = _to_float(_get_field(best_record, *FIELD_REMAIN))
        if best_remain < total_qty:
            print("[错误] 余量不足：余量最大的记录余量 %s，不足以覆盖发包总需求 %s"
                  % (best_remain, total_qty))
            sys.exit(1)
        print("采用余量记录：合同=%s 财务产品=%s 物料=%s 余量=%s"
              % (_cell_str(_get_field(best_record, *FIELD_CONTRACT)),
                 _cell_str(_get_field(best_record, *FIELD_FINANCIAL)),
                 _cell_str(_get_field(best_record, *FIELD_SPART)),
                 best_remain))

        build_output(rows, output_template, output_file, best_record)
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
