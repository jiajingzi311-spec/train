import os
# -*- coding: utf-8 -*-
import openpyxl, os

BASE = os.path.join(os.path.dirname(os.path.abspath(__file__)), "dist", "MRDN系统", "data", "inputs")  # 指向本机 dist 下的数据目录

files = [
    "2025 EXP Material for configuration_MOCK.xlsx",
    "物料设计文档.xlsx",
    "SmartMR模板.xlsx",
    "物料类型表.xlsx",
    "物料转置输出模板.xlsx",
    "输入模板.xlsx",
    "输出模板.xlsx",
]

def show(fname, maxr=8, maxc=14):
    p = os.path.join(BASE, fname)
    print("=" * 70)
    print("###", fname)
    if not os.path.exists(p):
        print("  [不存在]", p)
        return
    try:
        wb = openpyxl.load_workbook(p, data_only=True)
    except Exception as e:
        print("  [打开失败]", e)
        return
    for ws in wb.worksheets:
        print("  sheet:", repr(ws.title), "max_row:", ws.max_row, "max_col:", ws.max_column)
        for r in range(1, min(ws.max_row, maxr) + 1):
            vals = ["" if ws.cell(row=r, column=c).value is None
                    else str(ws.cell(row=r, column=c).value)[:16]
                    for c in range(1, min(ws.max_column, maxc) + 1)]
            print("   r%d:" % r, vals)
    wb.close()

for f in files:
    show(f)
