# -*- coding: utf-8 -*-
import openpyxl, os

BASE = r"D:\work_document\DOS工具开发\MRDN系统\dist\MRDN系统\data\inputs"

def sheets(fname):
    p = os.path.join(BASE, fname)
    wb = openpyxl.load_workbook(p, data_only=True, read_only=True)
    print("###", fname)
    for sn in wb.sheetnames:
        ws = wb[sn]
        print("   -", repr(sn), "| max_row=%s max_col=%s" % (ws.max_row, ws.max_column))
    wb.close()
    print()

for f in ["2025 EXP Material for configuration_080425au.xlsx",
          "物料设计文档.xlsx", "SmartMR模板.xlsx",
          "物料类型表.xlsx", "物料转置输出模板.xlsx"]:
    sheets(f)
