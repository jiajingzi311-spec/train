# -*- coding: utf-8 -*-
import openpyxl, os

BASE = r"D:\work_document\DOS工具开发\MRDN系统\dist\MRDN系统\data\inputs"

for fname in ["输入模板.xlsx", "输出模板.xlsx", "输出结果.xlsx"]:
    p = os.path.join(BASE, fname)
    print("=" * 70)
    print("###", fname)
    if not os.path.exists(p):
        print("  [不存在]")
        continue
    wb = openpyxl.load_workbook(p, data_only=True)
    for ws in wb.worksheets:
        print("  sheet:", repr(ws.title), "max_row:", ws.max_row, "max_col:", ws.max_column)
        for r in range(1, min(ws.max_row, 6) + 1):
            vals = ["" if ws.cell(row=r, column=c).value is None
                    else str(ws.cell(row=r, column=c).value)[:20]
                    for c in range(1, min(ws.max_column, 15) + 1)]
            print("   r%d:" % r, vals)
    wb.close()
