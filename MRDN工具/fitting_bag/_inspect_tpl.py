# -*- coding: utf-8 -*-
import openpyxl, os

for f in ['输入模板.xlsx', 'CPQ Spart导入模板.xlsx']:
    print('===', f, '===')
    wb = openpyxl.load_workbook(f, data_only=True)
    print('sheets:', wb.sheetnames)
    for sn in wb.sheetnames:
        ws = wb[sn]
        print('  sheet:', sn, 'max_row', ws.max_row, 'max_col', ws.max_column)
        print('  headers:', [c.value for c in ws[1]])
        if ws.max_row >= 2:
            print('  row2:', [c.value for c in ws[2]])
        if ws.max_row >= 3:
            print('  row3:', [c.value for c in ws[3]])
    print()
