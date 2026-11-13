import os
# -*- coding: utf-8 -*-
import openpyxl, os

BASE = os.path.join(os.path.dirname(os.path.abspath(__file__)), "dist", "MRDN系统", "data", "inputs")  # 指向本机 dist 下的数据目录

# 1. Item replace sheet
p = os.path.join(BASE, "2025 EXP Material for configuration_MOCK.xlsx")
wb = openpyxl.load_workbook(p, data_only=True)
ws = wb["Item replace definite V1 "]
print("### Item replace definite V1 (前 15 行)")
for r in range(1, min(ws.max_row, 15) + 1):
    vals = ["" if ws.cell(row=r, column=c).value is None else str(ws.cell(row=r, column=c).value)[:22]
            for c in range(1, min(ws.max_column, 8) + 1)]
    print("  r%d:" % r, vals)
wb.close()

# 2. 物料转置输出模板表头
p2 = os.path.join(BASE, "物料转置输出模板.xlsx")
wb2 = openpyxl.load_workbook(p2, data_only=True)
ws2 = wb2.active
print("\n### 物料转置输出模板 表头（第 1 行）")
hdr = ["" if ws2.cell(row=1, column=c).value is None else str(ws2.cell(row=1, column=c).value)
       for c in range(1, ws2.max_column + 1)]
for i, h in enumerate(hdr, 1):
    print("   col%d: %s" % (i, h))
wb2.close()

# 3. SmartMR 模板 Du Material 表头
p3 = os.path.join(BASE, "SmartMR模板.xlsx")
wb3 = openpyxl.load_workbook(p3, data_only=True)
ws3 = wb3["Du Material"]
print("\n### SmartMR模板 Du Material 表头（第 1 行）")
hdr3 = ["" if ws3.cell(row=1, column=c).value is None else str(ws3.cell(row=1, column=c).value)
        for c in range(1, min(ws3.max_column, 40) + 1)]
for i, h in enumerate(hdr3, 1):
    print("   col%d: %s" % (i, h))
wb3.close()
