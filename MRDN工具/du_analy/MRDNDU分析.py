# -*- coding: utf-8 -*-
"""
MRDN DU 分析工具（du_analyzer）
=================================
功能：
    读取 DUID 清单，按规则判定每个 DU 的站点类型（SiteStype）和项目类型，
    输出到带样式的 Excel（表头带颜色、字体加粗）。

判定规则（按优先级从上到下，大小写不敏感）：
    1. DU ID 含独立段 "SC"          → SiteStype = "Small cell"
    2. DU ID 含 "Changeout"/"changout" → SiteStype = "EOS"
    3. DU ID 含独立段 "FTK"          → SiteStype = "FTK&NSB"
    4. 结尾不含以上关键词（SC/Changeout/FTK）→ SiteStype = "EXP"，项目类型 = "R6K/DQN"

输出列：
    DU ID、SiteStype、项目类型。
    项目类型仅当 SiteStype == "EXP" 时填 "R6K/DQN"，其余留空（其他列先空出来）。

依赖：
    pip install openpyxl
"""


# ============================================================================
#  ⚙ MRDN 系统集成参数对照表（配置系统「工具输入关联」时，照此填「参数名」）
# ============================================================================
#   参数名（填到系统）       对应文件 / 含义
#   --du_list                DUID 清单（关联到系统「DUID 清单」输入项，系统自动生成文本文件）
#   --output_template        输出模板.xlsx（提供表头列；可选，缺省时用默认三列表头）
#   --duid                   单 DUID（单 DUID 作业只处理当前 DUID 时由系统传入）
#   （--out_dir 为系统自动传入的输出目录，无需在系统里配置）
# ============================================================================

import os
import re
import sys
import argparse

import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter


# 基础路径：打包成 exe 后，脚本同目录 = exe 所在目录
if getattr(sys, 'frozen', False):
    BASE_DIR = os.path.dirname(os.path.abspath(sys.executable))
else:
    BASE_DIR = os.path.dirname(os.path.abspath(__file__))

OUTPUT_FILE = "输出结果.xlsx"  # 固定、可预期的输出文件名


# ============================================================
# 工具函数
# ============================================================
def _norm(s):
    """归一化表头：只保留字母/数字/汉字，转小写。"""
    return re.sub(r"[^0-9a-zA-Z\u4e00-\u9fff]+", "", str(s)).lower()


def _configure_stdout():
    """Windows 打包后默认 GBK，中文 print 会崩；统一转 UTF-8。"""
    if hasattr(sys.stdout, "reconfigure"):
        try:
            sys.stdout.reconfigure(encoding="utf-8", errors="replace")
            sys.stderr.reconfigure(encoding="utf-8", errors="replace")
        except Exception:
            pass


def is_du_header(v):
    n = _norm(v)
    return n in ("duid", "du", "du号", "duid号", "du编号", "duid编号", "duidlist", "duid列表")


def is_sitestype_header(v):
    n = _norm(v)
    return n in ("sitestype", "sitetype", "站点类型", "站点", "stationtype")


def is_project_header(v):
    n = _norm(v)
    return n in ("项目类型", "projecttype", "项目", "project")


# ============================================================
# 读取 DUID 清单
# ============================================================
def read_duids(path):
    """读取 DUID 清单，返回去重后的 DUID 列表（保序，保留原始写法）。

    支持两种格式：
      - 文本（.txt/.csv/.list）：每行一个 DUID；若首行是表头（含 duid/du）则跳过。
      - Excel（.xlsx/.xlsm/.xls）：找含 DU ID 列名的表头，取其下非空值；
        找不到表头则回退为取第一列所有非空值。
    """
    if not path or not os.path.exists(path):
        return []

    result = []
    seen = set()

    def add(x):
        x = str(x).strip().lstrip("\ufeff")
        if x and x.lower() not in seen:
            seen.add(x.lower())
            result.append(x)

    ext = os.path.splitext(path)[1].lower()
    if ext in (".xlsx", ".xlsm", ".xls"):
        wb = openpyxl.load_workbook(path, data_only=True)
        ws = wb.active
        col = None
        data_start = None
        for r in range(1, min(ws.max_row, 10) + 1):
            for i in range(1, ws.max_column + 1):
                v = ws.cell(row=r, column=i).value
                if is_du_header(v):
                    col = i
                    data_start = r + 1
                    break
            if col is not None:
                break
        if col is not None:
            for rr in range(data_start, ws.max_row + 1):
                v = ws.cell(row=rr, column=col).value
                if v is None:
                    continue
                add(v)
        else:
            for rr in range(1, ws.max_row + 1):
                v = ws.cell(row=rr, column=1).value
                if v is None:
                    continue
                add(v)
    else:
        with open(path, encoding="utf-8", errors="ignore") as f:
            lines = [ln.strip() for ln in f if ln.strip()]
        start = 0
        if lines and is_du_header(lines[0]):
            start = 1
        for ln in lines[start:]:
            add(ln)

    return result


# ============================================================
# 站点类型 / 项目类型 判定
# ============================================================
def classify(duid):
    """判定单个 DUID，返回 (SiteStype, 项目类型)。"""
    s = str(duid).strip()
    up = s.upper()

    # 独立段（以非字母数字分隔的段），用于匹配 SC / FTK 这类短关键词，避免误命中
    segs = [x for x in re.split(r"[^0-9A-Za-z]+", up) if x]

    # 规则 1：含独立段 SC -> Small cell
    if "SC" in segs:
        return "Small cell", ""
    # 规则 2：含 Changeout / changout -> EOS
    if "CHANGEOUT" in up or "CHANGOUT" in up:
        return "EOS", ""
    # 规则 3：含独立段 FTK -> FTK&NSB
    if "FTK" in segs:
        return "FTK&NSB", ""
    # 规则 4：结尾不含以上关键词（SC/Changeout/FTK）-> EXP（项目类型 R6K/DQN）
    return "EXP", "R6K/DQN"


# ============================================================
# 输出（带样式）
# ============================================================
HEADER_FILL = PatternFill("solid", fgColor="305496")       # 深蓝表头
HEADER_FONT = Font(bold=True, color="FFFFFF", name="微软雅黑", size=11)
DATA_FONT = Font(name="微软雅黑", size=11)
CENTER = Alignment(horizontal="center", vertical="center", wrap_text=True)
THIN = Side(style="thin", color="BFBFBF")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)


def _read_template_headers(template_path):
    """读输出模板第一行表头，返回表头名列表（无模板返回空列表）。"""
    if not template_path or not os.path.exists(template_path):
        return []
    try:
        twb = openpyxl.load_workbook(template_path)
        tws = twb.active
        return [str(c.value).strip() if c.value is not None else "" for c in tws[1]]
    except Exception:
        return []


def build_output(rows, template_headers, out_path):
    """rows: [(duid, sitestype, project)]。写带样式的结果 Excel。"""
    # 确定输出表头：优先用输出模板表头；缺省三列
    if template_headers:
        headers = list(template_headers)
    else:
        headers = ["DU ID", "SiteStype", "项目类型"]

    def col_of(pred):
        for i, h in enumerate(headers):
            if pred(h):
                return i
        return None

    idx_duid = col_of(is_du_header)
    idx_stype = col_of(is_sitestype_header)
    idx_proj = col_of(is_project_header)

    # 若模板缺失关键列，补到表头末尾
    if idx_duid is None:
        idx_duid = len(headers)
        headers.append("DU ID")
    if idx_stype is None:
        idx_stype = len(headers)
        headers.append("SiteStype")
    if idx_proj is None:
        idx_proj = len(headers)
        headers.append("项目类型")

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "DU分析"

    # 表头（带色 + 加粗 + 居中）
    for i, h in enumerate(headers, start=1):
        c = ws.cell(row=1, column=i, value=h)
        c.fill = HEADER_FILL
        c.font = HEADER_FONT
        c.alignment = CENTER
        c.border = BORDER
    ws.freeze_panes = "A2"

    # 数据行
    for r, (duid, stype, proj) in enumerate(rows, start=2):
        vals = {idx_duid: duid, idx_stype: stype, idx_proj: proj}
        for col, val in vals.items():
            c = ws.cell(row=r, column=col + 1, value=val)
            c.font = DATA_FONT
            c.alignment = CENTER
            c.border = BORDER

    # 列宽
    for i, h in enumerate(headers, start=1):
        if is_du_header(h):
            w = 40
        elif is_sitestype_header(h):
            w = 16
        elif is_project_header(h):
            w = 14
        else:
            w = 12
        ws.column_dimensions[get_column_letter(i)].width = w

    wb.save(out_path)


# ============================================================
# 入口
# ============================================================
def parse_args():
    p = argparse.ArgumentParser(description="MRDN DU 分析工具")
    p.add_argument("--du_list", default=None, help="DUID 清单文件路径（系统自动生成）")
    p.add_argument("--output_template", default=None, help="输出模板.xlsx 路径（可选）")
    p.add_argument("--duid", default=None, help="单 DUID（只处理这一个 DUID，优先级高于 --du_list）")
    p.add_argument("--out_dir", default=None, help="输出目录（不传则输出到脚本同目录）")
    return p.parse_args()


def main():
    _configure_stdout()
    args = parse_args()

    out_dir = args.out_dir or BASE_DIR
    os.makedirs(out_dir, exist_ok=True)
    output_file = os.path.join(out_dir, OUTPUT_FILE)

    # 拿到 DUID 列表：单 DUID 优先，否则读清单
    if args.duid:
        duids = [str(args.duid).strip()]
        print("单 DUID 模式，只处理：", duids[0])
    else:
        if not args.du_list or not os.path.exists(args.du_list):
            print("[错误] 找不到 DUID 清单：%s" % args.du_list)
            sys.exit(1)
        duids = read_duids(args.du_list)
        if not duids:
            print("[错误] DUID 清单为空，没有可处理的 DUID")
            sys.exit(1)
        print("DUID 清单数量：", len(duids))

    # 判定
    rows = []
    for d in duids:
        stype, proj = classify(d)
        rows.append((d, stype, proj))

    # 读输出模板表头（可选）
    template_headers = _read_template_headers(args.output_template)

    build_output(rows, template_headers, output_file)

    # 打印汇总
    print("\n判定结果：")
    for d, stype, proj in rows:
        line = "  %s -> %s" % (d, stype)
        if proj:
            line += "（项目类型：%s）" % proj
        print(line)
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
