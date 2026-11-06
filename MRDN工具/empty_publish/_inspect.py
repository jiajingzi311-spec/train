# -*- coding: utf-8 -*-
import openpyxl
p = r"D:\work_document\python_doc\pythonProject\MRDN工具\empty_publish\空实例列表.xlsx"
wb = openpyxl.load_workbook(p, data_only=True)
print("sheets:", wb.sheetnames)
for sn in wb.sheetnames:
    ws = wb[sn]
    print("---", sn, "--- rows:", ws.max_row, "cols:", ws.max_column)
    for i, row in enumerate(ws.iter_rows(max_row=6, values_only=True)):
        print(i, [c if c is not None else "" for c in row])
