# -*- coding: utf-8 -*-
import openpyxl
from collections import Counter

print("=== 物料设计文档.xlsx ===")
try:
    wb = openpyxl.load_workbook(r'D:\work_document\DOS工具开发\MRDN系统\dist\MRDN系统\data\inputs\物料设计文档.xlsx', data_only=True)
    for ws in wb.worksheets:
        print(' sheet:', repr(ws.title), 'dims:', ws.dimensions)
        for r in range(1, min(ws.max_row, 5) + 1):
            row = [ws.cell(row=r, column=c).value for c in range(1, min(ws.max_column, 15) + 1)]
            print('  ', r, row)
except Exception as e:
    print(' ERROR', e)

print()
print("=== 2025 EXP Material：重复 DU + 日期列 ===")
f = r'D:\work_document\DOS工具开发\MRDN系统\dist\MRDN系统\data\inputs\2025 EXP Material for configuration_080425au.xlsx'
wb = openpyxl.load_workbook(f, data_only=True)
ws = wb.active
# 第3行表头
hdr = [ws.cell(row=3, column=c).value for c in range(1, 12)]
print(' 表头(第3行前11列):', hdr)

# 找 Date ready for configuration 列和 DU_ID 列
du_col = date_col = None
for c in range(1, ws.max_column + 1):
    v = ws.cell(row=3, column=c).value
    if v is None:
        continue
    s = str(v).strip().lower()
    if 'du_id' in s or s in ('duid', 'du', 'du_id'):
        du_col = c
    if 'date ready' in s or 'ready for configuration' in s.lower().replace('ready for configuration', 'date ready'):
        pass
    if 'date' in s and 'ready' in s:
        date_col = c
print(' DU_ID 列:', du_col, ' Date ready for configuration 列:', date_col)

# 检查重复 DU 和日期值
du_counter = Counter()
du_dates = {}
for r in range(4, ws.max_row + 1):
    duid = ws.cell(row=r, column=2).value  # B 列
    if duid is None or str(duid).strip() == '':
        continue
    duid = str(duid).strip()
    du_counter[duid] += 1
    d = ws.cell(row=r, column=5).value if date_col else None  # E 列
    if duid not in du_dates:
        du_dates[duid] = []
    du_dates[duid].append((r, d))

dups = {k: v for k, v in du_counter.items() if v > 1}
print(' 总 DU 数:', len(du_counter), ' 重复 DU 数:', len(dups))
print(' 重复 DU 示例（前 5 个）:')
for k, v in list(dups.items())[:5]:
    print('   ', k, '出现', v, '次，日期:', du_dates[k])
