# -*- coding: utf-8 -*-
import openpyxl
p = r"D:\work_document\DOS工具开发\MRDN系统\dist\MRDN系统\data\works\_batch\26年10月20日第一次作业\02_物料清单制作\输出结果.xlsx"
wb = openpyxl.load_workbook(p, data_only=True)
print("sheets:", wb.sheetnames)
ws = wb.active
print("--- active sheet:", ws.title, "rows:", ws.max_row, "cols:", ws.max_column, "---")
# 打印前 4 行（表头 + 说明行 + 数据）
for i, row in enumerate(ws.iter_rows(max_row=6, values_only=True)):
    print(i, [c if c is not None else "" for c in row])
