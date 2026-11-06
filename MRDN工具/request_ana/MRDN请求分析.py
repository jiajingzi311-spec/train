# -*- coding: utf-8 -*-
"""
MRDN 请求分析工具（request_ana）
=================================

功能：
1. 读「输入模板.xlsx」里的 DU ID；
2. 读「输入邮件模板.msg」：
   - 发送时间 -> request receive date
   - 正文里「主题:」那一行的内容 -> email title
3. week 按本地当天算，写成 weekXX（如 week42）；
4. Request type 写常数 DN；
5. 按规则判断 SiteStype 与 项目类型（规则 1~4，全部填充）；
6. 以「输出模板.xlsx」的 17 列表头为准生成结果，其余列留空；
7. 表头带底色、字体加粗、边框、冻结首行、自适应列宽。

依赖：
    pip install openpyxl extract_msg

用法：
    把 输入模板.xlsx / 输入邮件模板.msg / 输出模板.xlsx 放到本脚本同目录，
    直接运行：python MRDN请求分析.py
    结果写到同目录「输出结果.xlsx」（不会改动原模板）。
"""


# ============================================================================
#  ⚙ MRDN 系统集成参数对照表（配置系统「工具输入关联」时，照此填「参数名」）
# ============================================================================
#   参数名（填到系统）       对应文件 / 含义
#   --input_template         输入模板.xlsx（DU ID 清单）
#   --email_file             输入邮件模板.msg
#   --output_template        输出模板.xlsx（结果表头）
#   （--out_dir 为系统自动传入的输出目录，无需在系统里配置）
# ============================================================================

import os
import re
import sys
import datetime
import argparse

import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
import extract_msg


# ============================================================
# 可配置项（按需修改）
# ============================================================
if getattr(sys, 'frozen', False):
    BASE_DIR = os.path.dirname(os.path.abspath(sys.executable))
else:
    BASE_DIR = os.path.dirname(os.path.abspath(__file__))

INPUT_TEMPLATE = os.path.join(BASE_DIR, "输入模板.xlsx")   # 输入模板（DU ID）
EMAIL_FILE = os.path.join(BASE_DIR, "输入邮件模板.msg")     # 邮件模板（.msg）
OUTPUT_TEMPLATE = os.path.join(BASE_DIR, "输出模板.xlsx")   # 输出模板（17 列表头）
OUTPUT_FILE = os.path.join(BASE_DIR, "输出结果.xlsx")       # 最终输出（另存，不改模板）

REQUEST_TYPE = "DN"                                        # Request type 常数

# 站点类型 -> 项目类型（规则 1~4，全部填充）
#   规则1 含 SC                -> Small cell  / R6K,EU0
#   规则2 含 Changeout         -> EOS         / R6K,DQN,N0M
#   规则3 含 FTK               -> FTK&NSB     / R6K,DQN
#   规则4 结尾 WL 且不含以上   -> EXP         / R6K,DQN
SITE_PROJECT = {
    "Small cell": "56A0R6K/56A0EU0",
    "EOS": "56A0R6K/56A0DQN/56A0N0M",
    "FTK&NSB": "56A0R6K/56A0DQN",
    "EXP": "56A0R6K/56A0DQN",
}

# week 用「本地今天」计算
WEEK_FROM_TODAY = True


# ============================================================
# 站点类型判断
# ============================================================
def classify_site(du_id):
    """返回 (站点类型, 项目类型)。"""
    name = str(du_id).strip()
    if not name:
        return "请咨询TD", ""

    upper = name.upper()
    tokens = re.split(r"[_\-\+\s]+", upper)   # 按 _ - + 空格 拆段

    if "SC" in tokens:
        site = "Small cell"
    elif "CHANGEOUT" in upper:
        site = "EOS"
    elif "FTK" in tokens:
        site = "FTK&NSB"
    elif tokens and tokens[-1] == "WL":
        site = "EXP"
    else:
        return "请咨询TD", ""

    return site, SITE_PROJECT.get(site, "")


# ============================================================
# 读取输入 / 邮件
# ============================================================
def read_du_ids(path):
    """读取输入模板 / DUID 清单里的 DU ID（自动定位含 DU ID / DUID 的列）。

    支持两种格式：
      - Excel（.xlsx/.xlsm/.xls）：自动定位含 DU ID / DUID 的列，取该列非空值。
      - 文本（.txt/.csv）：优先识别表头（含 DU ID / DUID）；否则每行一个 DUID。
    系统批次作业时，这里传入的是动态生成的 DUID 清单（txt，每行一个 DUID）。
    """
    if not path or not os.path.exists(path):
        return []
    ext = os.path.splitext(path)[1].lower()
    if ext in (".xlsx", ".xlsm", ".xls"):
        wb = openpyxl.load_workbook(path, data_only=True)
        ws = wb.active

        du_col = None
        for row in ws.iter_rows(min_row=1, max_row=3, values_only=True):
            for idx, cell in enumerate(row):
                if cell is None:
                    continue
                if _norm(cell) in ("duid", "du", "duid2"):
                    du_col = idx
                    break
            if du_col is not None:
                break
        if du_col is None:
            du_col = 0

        du_ids = []
        for row in ws.iter_rows(min_row=2, values_only=True):
            if du_col >= len(row):
                continue
            val = row[du_col]
            if val is None or str(val).strip() == "":
                continue
            du_ids.append(str(val).strip())
        return du_ids

    # 文本清单：优先识别表头；否则每行一个 DUID（取每行第一个非空字段）
    with open(path, "r", encoding="utf-8-sig", errors="replace") as f:
        lines = [ln.strip() for ln in f if ln.strip()]
    if not lines:
        return []
    first_parts = re.split(r"[\t,，;；]", lines[0])
    du_col = next((i for i, p in enumerate(first_parts)
                   if p and _norm(p) in ("duid", "du", "duid2")), None)
    du_ids = []
    if du_col is not None:
        for line in lines[1:]:
            parts = re.split(r"[\t,，;；]", line)
            v = parts[du_col] if du_col < len(parts) else ""
            s = v.strip()
            if s:
                du_ids.append(s)
    else:
        for line in lines:
            for p in re.split(r"[\t,，;；]", line):
                s = p.strip()
                if s:
                    du_ids.append(s)
                    break
    return du_ids


def read_email(path):
    """读取 .msg，返回 (发送时间 datetime 或 None, 正文「主题:」后的标题)。"""
    msg = extract_msg.Message(path)
    dt = msg.date
    if dt is not None and getattr(dt, "tzinfo", None) is not None:
        dt = dt.astimezone().replace(tzinfo=None)   # 转本地并去掉时区

    title = _extract_subject_from_body(msg.body or "")
    return dt, title


def _extract_subject_from_body(body):
    """从邮件正文里找「主题:」那一行，返回冒号后的内容（优先中文「主题:」）。"""
    for line in body.splitlines():
        line = line.strip()
        m = re.match(r"^主题\s*[:：]\s*(.+)$", line)
        if m:
            return m.group(1).strip()
    for line in body.splitlines():
        line = line.strip()
        m = re.match(r"^Subject\s*[:：]\s*(.+)$", line, flags=re.IGNORECASE)
        if m:
            return m.group(1).strip()
    return ""


def _norm(s):
    return re.sub(r"[\s\-_]+", "", str(s)).lower()


def week_label():
    """返回 weekXX，XX = 本地当天的 ISO 周数。"""
    anchor = datetime.datetime.now() if WEEK_FROM_TODAY else None
    return "week{}".format(anchor.isocalendar()[1])


# ============================================================
# 输出（带样式）
# ============================================================
HEADER_FILL = PatternFill("solid", fgColor="305496")   # 表头深蓝
HEADER_FONT = Font(bold=True, color="FFFFFF", name="微软雅黑", size=11)
DATA_FONT = Font(name="微软雅黑", size=11)
CENTER = Alignment(horizontal="center", vertical="center", wrap_text=True)
LEFT = Alignment(horizontal="left", vertical="center", wrap_text=True)
THIN = Side(style="thin", color="BFBFBF")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
WARN_FILL = PatternFill("solid", fgColor="FFF2CC")     # “请咨询TD” 高亮


def _disp_width(s):
    w = 0
    for ch in str(s):
        w += 2 if ord(ch) > 127 else 1
    return w


def build_output(du_ids, dt, title, out_path, template_path):
    """以输出模板表头为准生成结果。"""
    wb = openpyxl.load_workbook(template_path)
    ws = wb.active

    # 读表头（第 1 行）
    headers = []
    for c in ws[1]:
        headers.append(str(c.value).strip() if c.value is not None else "")
    ncol = len(headers)

    # 列名 -> 列号（0 基）
    def col_of(*names):
        for i, h in enumerate(headers):
            nh = _norm(h)
            for n in names:
                if nh == _norm(n):
                    return i
        return None

    idx_date = col_of("request receive date", "receive date")
    idx_title = col_of("email title", "title")
    idx_du = col_of("DU ID", "DUID")
    idx_week = col_of("week")
    idx_req = col_of("request type")
    idx_site = col_of("sitestype", "site type")
    idx_proj = col_of("项目类型", "project type")

    # 清空旧数据（保留第 1 行表头）
    if ws.max_row > 1:
        ws.delete_rows(2, ws.max_row - 1)

    # 表头样式
    for i in range(ncol):
        c = ws.cell(row=1, column=i + 1)
        c.value = headers[i]
        c.fill = HEADER_FILL
        c.font = HEADER_FONT
        c.alignment = CENTER
        c.border = BORDER
    ws.freeze_panes = "A2"

    date_str = dt.strftime("%Y-%m-%d %H:%M:%S") if dt else ""
    wk = week_label()

    row_map = {
        idx_du: lambda d, s, p: d,
        idx_date: lambda d, s, p: date_str,
        idx_title: lambda d, s, p: title,
        idx_week: lambda d, s, p: wk,
        idx_req: lambda d, s, p: REQUEST_TYPE,
        idx_site: lambda d, s, p: s,
        idx_proj: lambda d, s, p: p,
    }

    for r, du in enumerate(du_ids, start=2):
        site, proj = classify_site(du)
        for col, fn in row_map.items():
            if col is None:
                continue
            c = ws.cell(row=r, column=col + 1, value=fn(du, site, proj))
            c.font = DATA_FONT
            c.border = BORDER
            c.alignment = LEFT if col in (idx_du, idx_title, idx_proj) else CENTER
            if fn(du, site, proj) == "请咨询TD":
                c.fill = WARN_FILL

    # 列宽自适应
    for i, h in enumerate(headers):
        max_w = _disp_width(h) + 2
        for r in range(2, len(du_ids) + 2):
            v = ws.cell(row=r, column=i + 1).value
            if v is not None:
                max_w = max(max_w, _disp_width(v))
        ws.column_dimensions[openpyxl.utils.get_column_letter(i + 1)].width = min(max_w + 4, 42)

    wb.save(out_path)


def _configure_stdout():
    """Windows 打包后默认 GBK，中文 print 会崩；统一转 UTF-8。"""
    if hasattr(sys.stdout, "reconfigure"):
        try:
            sys.stdout.reconfigure(encoding="utf-8", errors="replace")
            sys.stderr.reconfigure(encoding="utf-8", errors="replace")
        except Exception:
            pass


def parse_args():
    p = argparse.ArgumentParser(description="MRDN 请求分析工具")
    p.add_argument("--input_template", default=None, help="输入模板.xlsx 路径（不传则用脚本同目录默认文件）")
    p.add_argument("--email_file", default=None, help="输入邮件模板.msg 路径")
    p.add_argument("--output_template", default=None, help="输出模板.xlsx 路径")
    p.add_argument("--out_dir", default=None, help="输出目录（不传则输出到脚本同目录）")
    return p.parse_args()


def main():
    _configure_stdout()
    args = parse_args()
    input_template = args.input_template or INPUT_TEMPLATE
    email_file = args.email_file or EMAIL_FILE
    output_template = args.output_template or OUTPUT_TEMPLATE
    out_dir = args.out_dir or BASE_DIR
    os.makedirs(out_dir, exist_ok=True)
    output_file = os.path.join(out_dir, "输出结果.xlsx")

    for p, name in ((input_template, "输入模板"), (email_file, "邮件模板"),
                    (output_template, "输出模板")):
        if not os.path.exists(p):
            print("[错误] 找不到%s：%s" % (name, p))
            sys.exit(1)

    du_ids = read_du_ids(input_template)
    dt, title = read_email(email_file)

    print("DU ID 数量：", len(du_ids))
    for d in du_ids:
        print("  -", d)
    print("邮件发送时间：", dt)
    print("邮件正文「主题:」：", title)
    print("周数：", week_label())

    build_output(du_ids, dt, title, output_file, output_template)
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
