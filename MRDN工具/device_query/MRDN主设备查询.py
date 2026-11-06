# -*- coding: utf-8 -*-
"""
MRDN 主设备查询工具（device_query）
=====================================

功能：
    从「物料制作清单」筛选出每个 DU 下 type=Type 1 的主设备，生成「主设备列表」；
    再基于主设备列表里的 itemcode 调 ISC「库存现有量」接口逐项查询库存，
    按筛选规则过滤出可用库存，生成「库存查询结果表」。

处理规则：
    1. 物料制作清单：只取 type 为 "Type 1" 的行 → 主设备列表
       （DUID、描述、itemcode、qty、type）。
    2. 主设备列表的 itemcode 去重后，逐个调 ISC 接口查询（每次间隔控制频率）。
    3. ISC 接口返回的库存记录，提取以下 15 个字段（展示列 → 接口真实字段名）：
       onhandId→onhandId / Entity→entityId / Warehouse→warehouse /
       Project Code→projectCode / Project Name→projectName /
       Contract No.→contractNumber / Item→item / Inv Qty.→invQty /
       Reserved Qty.→reservedQty / Avi Qty.→actQty / Unit→unit /
       SubInv→subInv / ERP Org.→actPropertyTypeDesc /
       Quality Type→qualityTypeDesc / Goods Age(DAY)→goodsAge
    4. 库存记录筛选条件：
       - Entity(entityId) 等于 3449
       - Warehouse 为 GCW_BIN_PH 或 SLM_BIN_PH
       - Project Code 精确等于 56A0R6K
       - Project Name 模糊包含 "Globe Wireless"（不区分大小写）
       - Item 与查询的 itemcode 一致
       - Contract No. 结尾非 "0AX" / "0AG"（不区分大小写）
    5. 用筛选出的 Contract No. 到「合同信息表」查 Sub 列，Sub=EXP 才算符合。
    6. 占用逻辑：相同 itemcode 被多个 DU 使用时，按 DUID 顺序依次分配，
       每个 onhandId 记录剩余量递减，Avi Qty.（剩余）>= 本行 qty 才可分配，
       避免同一条库存被重复占用导致超分。
    7. CPQ 余量查询（第 3 步）：
       对「库存可用」!=「无可用库存」的行，用（库存可用合同号 + itemcode）调 CPQ
       「余量查询」接口 getContractRemainInfo，批量查询（合同号去重 + itemcode 去重，
       每批最多 CPQ_BATCH_SIZE 个合同号，pageSize=150）；返回 6 字段（见下）；
       多条记录选 remainQty 最多那条；相同 合同+物料 被多个 DU 使用时按剩余量递减
       防超分；remainQty >= 本行 qty 则「余量是否满足」填 contractNumber，否则填「余量不足」。
    8. 输出（共 5 个文件）：
       - 主设备列表.xlsx（DUID / 描述 / itemcode / qty / type / 库存可用 / 余量是否满足 / Sub）
         库存可用 = 匹配到的合同号；无匹配则「无可用库存」
         Sub = 该合同号在合同信息表里的 Sub 列值
         余量是否满足 = CPQ 满足填 contractNumber，不满足填「余量不足」，无库存留空
       - 库存查询结果表.xlsx（15 字段的可用库存记录，按 onhandId 去重）
       - 余量查询结果表.xlsx（CPQ 余量 6 字段：contractNumber / paymentUnit /
         financialCode / spartNumber / spartDescription / remainQty）
       - 主设备不满足清单.xlsx（主设备列表中「库存可用」=「无可用库存」的行）
       - 最终CPQ导入表.xlsx（CPQ Spart 模板：库存可用且余量满足的行）

依赖：openpyxl, requests
"""


# ============================================================================
#  ⚙ MRDN 系统集成参数对照表（配置系统「工具输入关联」时，照此填「参数名」）
# ============================================================================
#   参数名（填到系统）       对应文件 / 含义
#   --material_list          物料制作清单（物料转置表，含 DUID/描述/itemcode/qty/type 列）
#   --contract_table         合同信息表（含 Contract No. / Sub 列）
#   --cpq_template           CPQ模板（CPQ Spart导入模板，最终CPQ导入表用）
#   --isc_cfg                ISC 平台凭证（系统「平台凭证」下拉选 ISC，自动生成 JSON）
#   --cpq_cfg                CPQ 平台凭证（系统「平台凭证」下拉选 CPQ，自动生成 JSON）
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


# 基础路径：打包成 exe 后，脚本同目录 = exe 所在目录
if getattr(sys, 'frozen', False):
    BASE_DIR = os.path.dirname(os.path.abspath(sys.executable))
else:
    BASE_DIR = os.path.dirname(os.path.abspath(__file__))

DEVICE_LIST_FILENAME = "主设备列表.xlsx"
INV_RESULT_FILENAME = "库存查询结果表.xlsx"
REMAIN_RESULT_FILENAME = "余量查询结果表.xlsx"
UNSATISFIED_FILENAME = "主设备不满足清单.xlsx"
CPQ_OUTPUT_FILENAME = "最终CPQ导入表.xlsx"


# ============================================================================
#  接口与筛选相关常量
# ============================================================================
# ISC 库存现有量查询接口（固定地址，POST；路径 /500/1 表示 pageSize=500, pageNo=1）
ISC_URL_TEMPLATE = (
    "https://isc.huawei.com/isc/gw/com.huawei.supply.warehouse.iscwws:wws/"
    "iscp/wws/services/inner/invOnhandQuery/findOnhandDetailList/{page_size}/{page_no}"
)
ISC_PAGE_SIZE = 500

# 请求 body 里的固定值（照 cURL 原样；itemCode 逐项动态填）
ISC_ENTITY_ID = "3449"

# 库存记录筛选常量
WAREHOUSE_ALLOWED = ["GCW_BIN_PH", "SLM_BIN_PH"]
PROJECT_CODE_TARGET = "56A0R6K"
PROJECT_NAME_KEYWORD = "Globe Wireless"   # Project Name 模糊匹配关键字
CONTRACT_EXCLUDE_SUFFIX = ["0AX", "0AG"]  # Contract No. 结尾排除（不区分大小写）
SUB_TARGET = "EXP"                         # 合同信息表 Sub 列筛选值

# 访问频率控制：每次 ISC 请求间隔秒数（开发指导要求 1~5 秒）
REQUEST_INTERVAL = 2


# ============================================================================
#  CPQ 余量查询（第 3 步）相关常量
# ============================================================================
# CPQ 余量查询接口（固定地址，POST；Cookie 来自 CPQ 平台凭证 JSON）
CPQ_REMAIN_URL = (
    "https://cpq.huawei.com/unistar/cpq/gateway/CPQBEService/"
    "services/cpqDataRemainCalcService/getContractRemainInfo"
)
CPQ_PAGE_SIZE = 150     # 用户确认：pageSize 数字 150（不用文本格式）
CPQ_BATCH_SIZE = 100    # CPQ 每批最多合同号数量（参考 fitting_bag 批量场景）
CPQ_INTERVAL = 2        # CPQ 每次请求间隔秒数

# CPQ 查询 body 固定参数（除 contractNos/spartNumbers 外照 cURL 原样）
CPQ_QUERY_FIXED = {
    "financialCode": "",
    "paymentUnit": "",
    "salesTypes": [1, 3, 4, 5],
    "partDescription": "",
    "model": "",
    "remainQty": [1],
    "isOffset": 0,
    "serviceMaterialFlag": 0,
}

# CPQ Spart 导入模板 sheet 名与常数
CPQ_SHEET_TARGET = "Spart Configuration"
CPQ_PRODUCT_TYPE = "Equipment"
CPQ_PU_NO = "001"          # 文本格式，不能是数字格式

# 主设备列表「余量是否满足」列不满足时的标记
SATISFY_NO = "余量不足"


# ============================================================================
#  ISC 接口返回字段名别名（接口字段可能驼峰/下划线/大小写差异，容错匹配）
# ============================================================================
FIELD_ONHAND = ("onhandId", "onhandID", "onhand_id", "id")
FIELD_ENTITY = ("entityId", "entityID", "entityid", "entity")       # 实体：真实字段 entityId
FIELD_WAREHOUSE = ("warehouse", "warehouseName", "warehouseCode", "warehouseId")
FIELD_PROJ_CODE = ("projectCode", "projectcode")
FIELD_PROJ_NAME = ("projectName", "projectname")
FIELD_CONTRACT = ("contractNumber", "contractNo", "contractno", "contract")  # 合同号：真实字段 contractNumber
FIELD_ITEM = ("item", "itemCode", "itemcode")                       # 物料：真实字段 item
FIELD_INV_QTY = ("invQty", "invQuantity", "invqty", "onhandQty", "onhandQuantity")
FIELD_RESERVED = ("reservedQty", "reservedQuantity", "reserveQty")
FIELD_AVI = ("actQty", "actqty", "aviQty", "availableQty", "availQty")   # 可用量：真实字段 actQty
FIELD_UNIT = ("unit", "uom", "unitOfMeasure")
FIELD_SUBINV = ("subInv", "subinventory", "subInventory", "subInvCode")
FIELD_ERP = ("actPropertyTypeDesc", "actpropertytypedesc", "erpOrg", "erpOrganization")  # ERP Org：真实字段 actPropertyTypeDesc
FIELD_QUALITY = ("qualityTypeDesc", "qualitytypedesc", "qualityType", "quality")  # 质量类型：真实字段 qualityTypeDesc
FIELD_AGE = ("goodsAge", "goodsAgeDay", "goodsAgeDays", "age")

# 结果表 15 个展示列名（与需求一致）
RESULT_HEADERS = [
    "onhandId", "Entity", "Warehouse", "Project Code", "Project Name",
    "Contract No.", "Item", "Inv Qty.", "Reserved Qty.", "Avi Qty.",
    "Unit", "SubInv", "ERP Org.", "Quality Type", "Goods Age(DAY)",
]
FIELD_ALIASES = [
    FIELD_ONHAND, FIELD_ENTITY, FIELD_WAREHOUSE, FIELD_PROJ_CODE, FIELD_PROJ_NAME,
    FIELD_CONTRACT, FIELD_ITEM, FIELD_INV_QTY, FIELD_RESERVED, FIELD_AVI,
    FIELD_UNIT, FIELD_SUBINV, FIELD_ERP, FIELD_QUALITY, FIELD_AGE,
]

# ============================================================================
#  CPQ 接口返回字段名别名（接口字段可能大小写差异，_get_field 容错匹配）
# ============================================================================
FIELD_CPQ_CONTRACT = ("contractNumber", "contractNo", "contract")
FIELD_CPQ_PAYMENT = ("paymentUnit", "paymentunit")
FIELD_CPQ_FINANCIAL = ("financialCode", "financialProductCode", "financial")
FIELD_CPQ_SPART = ("spartNumber", "spartNo", "spart")
FIELD_CPQ_SPART_DESC = ("spartDescription", "spartdesc", "partDescription")
FIELD_CPQ_REMAIN = ("remainQty", "remainQuantity", "remain")

# 余量查询结果表 6 个展示列（与需求字段名一致）
CPQ_RESULT_HEADERS = [
    "contractNumber", "paymentUnit", "financialCode",
    "spartNumber", "spartDescription", "remainQty",
]
CPQ_FIELD_ALIASES = [
    FIELD_CPQ_CONTRACT, FIELD_CPQ_PAYMENT, FIELD_CPQ_FINANCIAL,
    FIELD_CPQ_SPART, FIELD_CPQ_SPART_DESC, FIELD_CPQ_REMAIN,
]


# ============================================================================
#  基础工具函数
# ============================================================================
def _configure_stdout():
    """Windows 打包后默认 GBK，中文 print 会崩；统一转 UTF-8。"""
    if hasattr(sys.stdout, "reconfigure"):
        try:
            sys.stdout.reconfigure(encoding="utf-8", errors="replace")
            sys.stderr.reconfigure(encoding="utf-8", errors="replace")
        except Exception:
            pass


def normalize(s):
    """归一化列名：去非字母数字中文，转小写，便于匹配。"""
    return re.sub(r"[^A-Za-z0-9\u4e00-\u9fff]", "", str(s)).lower()


def find_col(headers, *names):
    """按归一化名在表头中找列索引（0-based），找不到返回 None。"""
    for i, h in enumerate(headers):
        if normalize(h) in [normalize(n) for n in names]:
            return i
    return None


def _cell_str(v):
    """单元格/接口值转干净字符串：float 整数去 .0，None/空返回空串。"""
    if v is None:
        return ""
    if isinstance(v, float) and v.is_integer():
        return str(int(v))
    if isinstance(v, (dict, list)):
        try:
            return json.dumps(v, ensure_ascii=False)
        except Exception:
            return str(v)
    return str(v).strip()


def _to_float(v):
    """安全转 float（数量可能是数字或字符串）。"""
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


def _extract_total(node, _depth=0):
    """从接口返回里递归找「总数」字段（total/count），找不到返回 None。"""
    if _depth > 10 or node is None:
        return None
    if isinstance(node, dict):
        for k, v in node.items():
            nk = normalize(k)
            if nk in ("total", "totalcount", "count", "totalnum", "totalrows"):
                if isinstance(v, (int, float)):
                    return int(v)
                if isinstance(v, str) and v.strip().isdigit():
                    return int(v.strip())
            r = _extract_total(v, _depth + 1)
            if r is not None:
                return r
    elif isinstance(node, list):
        for x in node:
            r = _extract_total(x, _depth + 1)
            if r is not None:
                return r
    return None


# ============================================================================
#  输入读取
# ============================================================================
def read_material(path):
    """读物料制作清单，筛 type=Type 1 的行，返回主设备列表
    [(duid, 描述, itemcode, qty, type), ...]，保持原始顺序。"""
    wb = openpyxl.load_workbook(path, data_only=True)
    ws = wb.active
    headers = [str(c.value).strip() if c.value is not None else "" for c in ws[1]]

    col_duid = find_col(headers, "DUID")
    col_desc = find_col(headers, "描述", "Description", "物料描述")
    col_item = find_col(headers, "itemcode", "Item Code", "ItemCode", "物料编码")
    col_qty = find_col(headers, "qty", "Qty", "数量")
    col_type = find_col(headers, "type", "Type", "类型")
    missing = [n for n, c in [("DUID", col_duid), ("itemcode", col_item),
                              ("qty", col_qty), ("type", col_type)] if c is None]
    if missing:
        raise RuntimeError("物料制作清单缺少必需列：%s" % " / ".join(missing))

    rows = []
    for r in range(2, ws.max_row + 1):
        type_val = ws.cell(row=r, column=col_type + 1).value
        if type_val is None:
            continue
        if normalize(type_val) != "type1":
            continue
        duid = _cell_str(ws.cell(row=r, column=col_duid + 1).value)
        item = _cell_str(ws.cell(row=r, column=col_item + 1).value)
        qty = ws.cell(row=r, column=col_qty + 1).value
        desc = _cell_str(ws.cell(row=r, column=col_desc + 1).value) if col_desc is not None else ""
        if not duid or not item:
            continue
        rows.append((duid, desc, item, qty, _cell_str(type_val)))
    return rows


def read_contract_exp(path):
    """读合同信息表，筛选 Sub=EXP 的记录，返回 {规范化合同号: Sub原始值} 映射。"""
    wb = openpyxl.load_workbook(path, data_only=True)
    ws = wb.active
    headers = [str(c.value).strip() if c.value is not None else "" for c in ws[1]]

    col_contract = find_col(headers, "Contract No.", "Contract No", "ContractNo",
                            "Contract Number", "合同号", "合同编号")
    col_sub = find_col(headers, "Sub", "Sub Type", "SubType", "子项目", "类型")
    if col_contract is None or col_sub is None:
        raise RuntimeError("合同信息表缺少必需列：Contract No. / Sub")

    result = {}
    for r in range(2, ws.max_row + 1):
        sub = ws.cell(row=r, column=col_sub + 1).value
        if sub is None:
            continue
        if normalize(sub) != SUB_TARGET.lower():
            continue
        cno = _cell_str(ws.cell(row=r, column=col_contract + 1).value)
        if not cno:
            continue
        result.setdefault(_norm_contract(cno), _cell_str(sub))
    return result


def _norm_contract(cno):
    """规范化合同号：去空白、转大写，用于对比。"""
    return re.sub(r"\s+", "", str(cno)).upper()


# ============================================================================
#  平台凭证读取
# ============================================================================
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
        raise RuntimeError("平台凭证里没有 Cookie，请先在「平台网页配置」里一键登录抓取 Cookie")
    headers.setdefault("Content-Type", "application/json")
    headers.setdefault("Accept", "application/json, text/plain, */*")
    return headers


# ============================================================================
#  ISC 查询
# ============================================================================
def build_isc_body(item_code):
    """构造 ISC 接口请求 body（照 cURL 原样，仅 itemCode 动态填）。"""
    return {
        "projectCode": "", "projectName": "", "contractNo": "", "subInv": "",
        "locatorCode": "", "zoneCode": "", "clNo": "", "boxName": "",
        "originalCl": "", "lpnClass": "",
        "itemCode": item_code,
        "frozenFlag": "", "boxNoFrom": "", "boxNoTo": "", "itemRevision": "",
        "actualProperty": "", "invQtyFrom": "", "invQtyTo": "", "boxType": "",
        "property": "", "propertyDesc": "", "siteType": "", "spu": "", "site": "",
        "goodsAgeFrom": "", "goodsAgeTo": "", "itemWithoutBox": "", "frozenLockedFlag": "",
        "entityId": ISC_ENTITY_ID, "warehouseId": "",
        "boxNameList": [], "locatorList": [], "itemCodeList": [], "subInvList": [],
        "contractNoList": [], "clNoList": [], "originalCls": [], "locator": "",
    }


def query_inv(item_code, headers):
    """查 ISC 库存接口，返回该 itemcode 的原始记录 list（可能翻页）。"""
    url = ISC_URL_TEMPLATE.format(page_size=ISC_PAGE_SIZE, page_no=1)
    body = build_isc_body(item_code)
    try:
        resp = requests.post(url, json=body, headers=headers, verify=False, timeout=90)
    except Exception as e:
        raise RuntimeError("ISC 接口请求失败（itemcode=%s）：%s" % (item_code, e))

    if resp.status_code != 200:
        raise RuntimeError("ISC 接口返回非 200（itemcode=%s）：HTTP %s\n%s"
                           % (item_code, resp.status_code, resp.text[:500]))

    try:
        data = resp.json()
    except ValueError:
        raise RuntimeError("ISC 接口返回非 JSON（itemcode=%s）：%s" % (item_code, resp.text[:300]))

    rows = _extract_rows(data)
    total = _extract_total(data)

    # 翻页：若返回条数不足 total 且还有数据，继续取下一页
    if total is not None and len(rows) < total and rows:
        page_no = 2
        while len(rows) < total:
            url = ISC_URL_TEMPLATE.format(page_size=ISC_PAGE_SIZE, page_no=page_no)
            try:
                r2 = requests.post(url, json=body, headers=headers, verify=False, timeout=90)
            except Exception as e:
                print("[警告] 翻页失败（page=%d）：%s" % (page_no, e))
                break
            if r2.status_code != 200:
                print("[警告] 翻页返回非 200（page=%d）" % page_no)
                break
            try:
                d2 = r2.json()
            except ValueError:
                break
            page_rows = _extract_rows(d2)
            if not page_rows:
                break
            rows.extend(page_rows)
            page_no += 1
            if page_no > 20:   # 安全上限，避免死循环
                break

    return rows


def extract_record(row):
    """把接口原始 dict 提取成 15 字段列表（按 RESULT_HEADERS 顺序），字段缺失为空串。"""
    return [_cell_str(_get_field(row, *alias)) for alias in FIELD_ALIASES]


def filter_inv_record(row, item_code):
    """基础筛选（不含 Avi Qty 与 Sub，这两项在分配阶段按具体 qty 判断）。"""
    entity = _cell_str(_get_field(row, *FIELD_ENTITY))
    warehouse = _cell_str(_get_field(row, *FIELD_WAREHOUSE)).upper()
    proj_code = _cell_str(_get_field(row, *FIELD_PROJ_CODE)).upper()
    proj_name = _cell_str(_get_field(row, *FIELD_PROJ_NAME))
    contract = _cell_str(_get_field(row, *FIELD_CONTRACT)).upper()
    item = _cell_str(_get_field(row, *FIELD_ITEM))

    if entity != ISC_ENTITY_ID:
        return False
    if warehouse not in [w.upper() for w in WAREHOUSE_ALLOWED]:
        return False
    if proj_code != PROJECT_CODE_TARGET.upper():
        return False
    if PROJECT_NAME_KEYWORD.lower() not in proj_name.lower():
        return False
    if item != item_code:
        return False
    if not contract:
        return False
    if any(contract.endswith(sfx.upper()) for sfx in CONTRACT_EXCLUDE_SUFFIX):
        return False
    return True


# ============================================================================
#  输出写入
# ============================================================================
def build_device_list(rows, out_dir):
    """写主设备列表.xlsx，返回文件路径。"""
    path = os.path.join(out_dir, DEVICE_LIST_FILENAME)
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "主设备列表"
    headers = ["DUID", "描述", "itemcode", "qty", "type", "库存可用", "余量是否满足", "Sub"]
    ws.append(headers)
    for r in rows:
        ws.append(list(r))
    _style_header(ws)
    _style_data(ws)
    wb.save(path)
    return path


def build_inv_result(records, out_dir):
    """写库存查询结果表.xlsx（15 字段，去重后的可用库存），返回文件路径。"""
    path = os.path.join(out_dir, INV_RESULT_FILENAME)
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "库存查询结果"
    ws.append(RESULT_HEADERS)
    for rec in records:
        ws.append(rec)
    _style_header(ws)
    _style_data(ws)
    wb.save(path)
    return path


def build_remain_result(records, out_dir):
    """写余量查询结果表.xlsx（CPQ 余量 6 字段），返回文件路径。"""
    path = os.path.join(out_dir, REMAIN_RESULT_FILENAME)
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "余量查询结果"
    ws.append(CPQ_RESULT_HEADERS)
    if records:
        for rec in records:
            ws.append([rec.get(h, "") for h in CPQ_RESULT_HEADERS])
    else:
        ws.append(["无余量查询记录（未配置 CPQ 凭证或无可用库存）"])
    _style_header(ws)
    _style_data(ws)
    wb.save(path)
    return path


def build_unsatisfied(rows, out_dir):
    """写主设备不满足清单.xlsx（库存可用=无可用库存的行），返回文件路径。"""
    path = os.path.join(out_dir, UNSATISFIED_FILENAME)
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "主设备不满足清单"
    headers = ["DUID", "描述", "itemcode", "qty", "type", "库存可用", "Sub"]
    ws.append(headers)
    for r in rows:
        ws.append(list(r))
    _style_header(ws)
    _style_data(ws)
    wb.save(path)
    return path


def build_cpq_output(rows, template_path, out_dir):
    """写最终CPQ导入表.xlsx（CPQ Spart 模板填数据），返回文件路径。

    rows: [(duid, item, qty, contract, financialCode), ...]
    无数据且无模板时返回 None（跳过输出）。
    """
    if not rows and (not template_path or not os.path.exists(template_path)):
        return None
    if not template_path or not os.path.exists(template_path):
        raise RuntimeError("生成最终CPQ导入表需要 CPQ模板（--cpq_template）")

    path = os.path.join(out_dir, CPQ_OUTPUT_FILENAME)
    wb = openpyxl.load_workbook(template_path)
    if CPQ_SHEET_TARGET not in wb.sheetnames:
        raise RuntimeError("CPQ模板缺少 sheet：%s" % CPQ_SHEET_TARGET)
    ws = wb[CPQ_SHEET_TARGET]
    headers = [str(c.value).strip() if c.value is not None else "" for c in ws[1]]

    col_contract = find_col(headers, "HUAWEI CONTRACT NO. *")
    col_product = find_col(headers, "Product Type *")
    col_puno = find_col(headers, "PU NO. *")
    col_duid = find_col(headers, "DU ID *")
    col_fin = find_col(headers, "FINANCIAL PRODUCT CODE *")
    col_spart = find_col(headers, "SPART NUMBER *")
    col_qty = find_col(headers, "SPART QTY. *")

    # 清空第 2 行起的数据，保留表头与格式
    if ws.max_row >= 2:
        ws.delete_rows(2, ws.max_row - 1)

    for i, (duid, item, qty, contract, fin) in enumerate(rows):
        r = 2 + i
        if col_contract is not None:
            ws.cell(row=r, column=col_contract + 1, value=contract)
        if col_product is not None:
            ws.cell(row=r, column=col_product + 1, value=CPQ_PRODUCT_TYPE)
        if col_puno is not None:
            c = ws.cell(row=r, column=col_puno + 1, value=CPQ_PU_NO)
            c.number_format = "@"   # 文本格式，保留前导 0，不能是数字格式
        if col_duid is not None:
            ws.cell(row=r, column=col_duid + 1, value=duid)
        if col_fin is not None:
            ws.cell(row=r, column=col_fin + 1, value=fin)
        if col_spart is not None:
            ws.cell(row=r, column=col_spart + 1, value=item)
        if col_qty is not None:
            ws.cell(row=r, column=col_qty + 1, value=qty)
    wb.save(path)
    return path


def build_cpq_body(contract_nos, spart_numbers):
    """构造 CPQ 余量批量查询 body（contractNos/spartNumbers 传列表，支持批量）。"""
    return {
        "page": {"curPage": 1, "pageSize": CPQ_PAGE_SIZE},
        "query": dict(CPQ_QUERY_FIXED, contractNos=list(contract_nos), spartNumbers=list(spart_numbers)),
    }


def query_cpq_remain(contract_nos, spart_numbers, headers):
    """批量查 CPQ 余量接口，返回原始记录 list（元素为 dict）。"""
    body = build_cpq_body(contract_nos, spart_numbers)
    try:
        resp = requests.post(CPQ_REMAIN_URL, json=body, headers=headers,
                             verify=False, timeout=90)
    except Exception as e:
        raise RuntimeError("CPQ 接口请求失败（contracts=%d, items=%d）：%s"
                           % (len(contract_nos), len(spart_numbers), e))
    if resp.status_code != 200:
        raise RuntimeError("CPQ 接口返回非 200（contracts=%d, items=%d）：HTTP %s\n%s"
                           % (len(contract_nos), len(spart_numbers), resp.status_code, resp.text[:500]))
    try:
        data = resp.json()
    except ValueError:
        raise RuntimeError("CPQ 接口返回非 JSON（contracts=%d, items=%d）：%s"
                           % (len(contract_nos), len(spart_numbers), resp.text[:300]))
    return _extract_rows(data)


HEADER_FONT = Font(name="微软雅黑", size=11, bold=True)
DATA_FONT = Font(name="微软雅黑", size=11)
CENTER = Alignment(horizontal="center", vertical="center", wrap_text=True)
THIN = Side(style="thin", color="BFBFBF")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)


def _style_header(ws):
    for c in ws[1]:
        c.font = HEADER_FONT
        c.alignment = CENTER
        c.border = BORDER


def _style_data(ws):
    for row in ws.iter_rows(min_row=2):
        for c in row:
            c.font = DATA_FONT
            c.alignment = CENTER
            c.border = BORDER


# ============================================================================
#  主流程
# ============================================================================
def main():
    _configure_stdout()
    p = argparse.ArgumentParser(description="MRDN 主设备查询工具")
    p.add_argument("--material_list", required=True, help="物料制作清单路径")
    p.add_argument("--contract_table", required=True, help="合同信息表路径")
    p.add_argument("--cpq_template", default=None, help="CPQ模板路径（CPQ Spart导入模板）")
    p.add_argument("--isc_cfg", required=True, help="ISC 平台凭证 JSON 路径")
    p.add_argument("--cpq_cfg", default=None, help="CPQ 平台凭证 JSON 路径")
    p.add_argument("--out_dir", required=True, help="输出目录")
    args = p.parse_args()

    out_dir = args.out_dir
    os.makedirs(out_dir, exist_ok=True)

    # 1. 读物料清单 → 主设备列表
    print("[1/7] 读取物料制作清单，筛选 Type 1 主设备...")
    device_rows = read_material(args.material_list)
    if not device_rows:
        print("[提示] 物料制作清单中没有 type=Type 1 的主设备，仅生成空表。")
        build_device_list([], out_dir)
        build_inv_result([], out_dir)
        build_remain_result([], out_dir)
        build_unsatisfied([], out_dir)
        build_cpq_output([], args.cpq_template, out_dir)
        return
    print("      主设备共 %d 行（去重 itemcode %d 个）"
          % (len(device_rows), len(set(r[2] for r in device_rows))))

    # 2. 读合同信息表 → Sub=EXP 的合同号集合
    print("[2/7] 读取合同信息表，筛选 Sub=EXP 的合同号...")
    exp_contracts = read_contract_exp(args.contract_table)
    print("      符合条件的合同号 %d 个" % len(exp_contracts))

    # 3. 加载 ISC 平台凭证
    print("[3/7] 加载 ISC 平台凭证...")
    isc_headers = load_platform_cfg(args.isc_cfg)

    # 4. 逐 itemcode 查询 ISC 库存
    print("[4/7] 调 ISC 接口逐 itemcode 查询库存（每次间隔 %d 秒）..." % REQUEST_INTERVAL)
    unique_items = []
    seen_items = set()
    for r in device_rows:
        if r[2] not in seen_items:
            seen_items.add(r[2])
            unique_items.append(r[2])

    inv_by_item = {}      # itemcode -> [筛选后的记录 dict（含 onhandId/contract/avi 等）]
    for idx, item in enumerate(unique_items, 1):
        print("      (%d/%d) 查询 itemcode=%s ..." % (idx, len(unique_items), item))
        raw_rows = query_inv(item, isc_headers)
        filtered = []
        for row in raw_rows:
            if filter_inv_record(row, item):
                filtered.append(row)
        inv_by_item[item] = filtered
        print("            返回 %d 条，筛选后 %d 条" % (len(raw_rows), len(filtered)))
        if idx < len(unique_items):
            time.sleep(REQUEST_INTERVAL)

    # 5. 占用分配：按 DUID 顺序，剩余量递减
    print("[5/7] 分配库存（占用逻辑：Avi Qty. 剩余 >= qty 且合同 Sub=EXP）...")
    remain_by_onhand = {}   # onhandId -> 剩余 Avi Qty.
    used_by_onhand = {}     # onhandId -> 15 字段列表（用于库存查询结果表，去重）
    result_rows = []
    for (duid, desc, item, qty, type_val) in device_rows:
        q = _to_float(qty)
        records = inv_by_item.get(item, [])
        chosen = None
        for rec in records:
            onhand = _cell_str(_get_field(rec, *FIELD_ONHAND))
            if not onhand:
                continue
            contract = _cell_str(_get_field(rec, *FIELD_CONTRACT))
            if _norm_contract(contract) not in exp_contracts:
                continue
            avi = _to_float(_get_field(rec, *FIELD_AVI))
            remain = remain_by_onhand.get(onhand, avi)
            if remain >= q:
                chosen = (onhand, contract, rec)
                remain_by_onhand[onhand] = remain - q
                break
        if chosen:
            onhand, contract, rec = chosen
            sub_val = exp_contracts.get(_norm_contract(contract), "")
            result_rows.append((duid, desc, item, qty, type_val, contract, sub_val))
            if onhand not in used_by_onhand:
                used_by_onhand[onhand] = extract_record(rec)
        else:
            result_rows.append((duid, desc, item, qty, type_val, "无可用库存", ""))

    # 6. CPQ 余量查询：批量查询（合同号去重 + itemcode 去重，分批传）
    avail_contracts = []   # 去重原始合同号（库存可用的行）
    seen_c = set()
    avail_items = []       # 去重 itemcode
    seen_i = set()
    for r in result_rows:
        if r[5] == "无可用库存":
            continue
        if r[5] not in seen_c:
            seen_c.add(r[5])
            avail_contracts.append(r[5])
        if r[2] not in seen_i:
            seen_i.add(r[2])
            avail_items.append(r[2])

    cpq_records = {}   # (norm_contract, item) -> [6字段 dict，按 remainQty 降序]
    all_cpq_rows = []  # 余量查询结果表所有记录
    do_cpq = bool(args.cpq_cfg) and os.path.exists(args.cpq_cfg)
    if do_cpq and avail_contracts:
        print("[6/7] 调 CPQ 接口批量查询余量（%d 个合同号 + %d 个物料，每批 %d 合同号，间隔 %d 秒）..."
              % (len(avail_contracts), len(avail_items), CPQ_BATCH_SIZE, CPQ_INTERVAL))
        cpq_headers = load_platform_cfg(args.cpq_cfg)
        n_batch = (len(avail_contracts) + CPQ_BATCH_SIZE - 1) // CPQ_BATCH_SIZE
        for bi in range(0, len(avail_contracts), CPQ_BATCH_SIZE):
            batch = avail_contracts[bi:bi + CPQ_BATCH_SIZE]
            bno = bi // CPQ_BATCH_SIZE + 1
            print("      (%d/%d) 查询 %d 个合同号 ..." % (bno, n_batch, len(batch)))
            raw_rows = query_cpq_remain(batch, avail_items, cpq_headers)
            for row in raw_rows:
                rec = {}
                for h, alias in zip(CPQ_RESULT_HEADERS, CPQ_FIELD_ALIASES):
                    rec[h] = _cell_str(_get_field(row, *alias))
                key = (_norm_contract(rec["contractNumber"]), rec["spartNumber"])
                cpq_records.setdefault(key, []).append(rec)
                all_cpq_rows.append(rec)
            print("            返回 %d 条记录" % len(raw_rows))
            if bi + CPQ_BATCH_SIZE < len(avail_contracts):
                time.sleep(CPQ_INTERVAL)
        # 每个组合的记录按 remainQty 降序
        for key in cpq_records:
            cpq_records[key].sort(key=lambda x: _to_float(x["remainQty"]), reverse=True)
    else:
        if not do_cpq:
            print("[6/7] 未配置 CPQ 凭证，跳过 CPQ 余量查询。")
        else:
            print("[6/7] 无库存可用的行，跳过 CPQ 余量查询。")

    # 7. 组装最终主设备列表（8 列）+ CPQ 余量占用分配
    print("[7/7] 组装结果（CPQ 余量占用分配）...")
    device_final = []   # 9 元组 (duid, desc, item, qty, type, contract, sub, satisfy, fin)
    cpq_remain = {}     # (norm_contract, item, idx) -> 剩余 remainQty
    for (duid, desc, item, qty, type_val, contract, sub_val) in result_rows:
        if contract == "无可用库存":
            device_final.append((duid, desc, item, qty, type_val, contract, sub_val, "", ""))
            continue
        q = _to_float(qty)
        key = (_norm_contract(contract), item)
        recs = cpq_records.get(key, [])
        chosen = None
        for idx, rec in enumerate(recs):
            remain_key = (key[0], key[1], idx)
            remain = cpq_remain.get(remain_key, _to_float(rec["remainQty"]))
            if remain >= q:
                cpq_remain[remain_key] = remain - q
                chosen = rec
                break
        if chosen:
            contract_no = chosen["contractNumber"] or contract
            fin = chosen["financialCode"]
            device_final.append((duid, desc, item, qty, type_val, contract, sub_val, contract_no, fin))
        else:
            device_final.append((duid, desc, item, qty, type_val, contract, sub_val, SATISFY_NO, ""))

    # 写结果（5 个文件）
    inv_records = list(used_by_onhand.values())
    inv_records.sort(key=lambda r: (_norm_contract(r[5]), -_to_float(r[9])))
    p1 = build_device_list([(r[0], r[1], r[2], r[3], r[4], r[5], r[7], r[6]) for r in device_final], out_dir)
    p2 = build_inv_result(inv_records, out_dir)
    p3 = build_remain_result(all_cpq_rows, out_dir)

    unsat_rows = [r for r in result_rows if r[5] == "无可用库存"]
    p4 = build_unsatisfied(unsat_rows, out_dir)

    cpq_rows = [(r[0], r[2], r[3], r[5], r[8])
                for r in device_final if r[7] and r[7] != SATISFY_NO]
    p5 = build_cpq_output(cpq_rows, args.cpq_template, out_dir)

    avail = sum(1 for r in result_rows if r[5] != "无可用库存")
    satisfied = sum(1 for r in device_final if r[7] and r[7] != SATISFY_NO)
    print("")
    print("完成！")
    print("  主设备列表：%s（共 %d 行，库存可用 %d 行，余量满足 %d 行）"
          % (p1, len(result_rows), avail, satisfied))
    print("  库存查询结果表：%s（%d 条可用库存）" % (p2, len(inv_records)))
    print("  余量查询结果表：%s（%d 条）" % (p3, len(all_cpq_rows)))
    print("  主设备不满足清单：%s（%d 行）" % (p4, len(unsat_rows)))
    if p5:
        print("  最终CPQ导入表：%s（%d 行）" % (p5, len(cpq_rows)))
    else:
        print("  最终CPQ导入表：无余量满足数据，跳过")


if __name__ == "__main__":
    main()
