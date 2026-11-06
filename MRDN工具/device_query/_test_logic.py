# -*- coding: utf-8 -*-
import importlib.util

spec = importlib.util.spec_from_file_location('m', 'MRDN主设备查询.py')
m = importlib.util.module_from_spec(spec)
spec.loader.exec_module(m)

# 用真实字段名构造假 ISC 响应
recs = {
    '03058754': [
        {'onhandId': 'O1', 'entityId': 3449, 'warehouse': 'GCW_BIN_PH',
         'projectCode': '56A0R6K', 'projectName': 'Globe Wireless Solution',
         'contractNumber': '0006082300010E', 'item': '03058754',
         'invQty': 1, 'reservedQty': 0, 'actQty': 1,
         'unit': 'EA', 'subInv': 'SUB1', 'actPropertyTypeDesc': 'xxx',
         'qualityTypeDesc': 'good', 'goodsAge': 10},
        {'onhandId': 'O2', 'entityId': 3449, 'warehouse': 'SLM_BIN_PH',
         'projectCode': '56A0R6K', 'projectName': 'Globe Wireless XX',
         'contractNumber': '0006082300010E', 'item': '03058754',
         'invQty': 1, 'reservedQty': 0, 'actQty': 1,
         'unit': 'EA', 'subInv': 'SUB2', 'actPropertyTypeDesc': 'yyy',
         'qualityTypeDesc': 'good', 'goodsAge': 20},
        {'onhandId': 'O3', 'entityId': 9999, 'warehouse': 'GCW_BIN_PH',
         'projectCode': '56A0R6K', 'projectName': 'Globe Wireless',
         'contractNumber': '0006082300010E', 'item': '03058754', 'actQty': 5},  # entityId 不符
        {'onhandId': 'O4', 'entityId': 3449, 'warehouse': 'OTHER_WH',
         'projectCode': '56A0R6K', 'projectName': 'Globe Wireless',
         'contractNumber': '0006082300010E', 'item': '03058754', 'actQty': 5},  # warehouse 不符
        {'onhandId': 'O5', 'entityId': 3449, 'warehouse': 'GCW_BIN_PH',
         'projectCode': '56A0R6K', 'projectName': 'Globe Wireless',
         'contractNumber': '0006082300010AX', 'item': '03058754', 'actQty': 5},  # contract 结尾 0AX 排除
        {'onhandId': 'O6', 'entityId': 3449, 'warehouse': 'GCW_BIN_PH',
         'projectCode': '56A0R6K', 'projectName': 'Globe Wireless',
         'contractNumber': '0006082300010AG', 'item': '03058754', 'actQty': 5},  # contract 结尾 0AG 排除
    ],
}
device_rows = [('SL4078', 'd1', '03058754', 1, 'Type 1'),
               ('SL18', 'd2', '03058754', 1, 'Type 1')]
exp = {'0006082300010E': 'EXP'}

inv_by_item = {}
for item, raw in recs.items():
    inv_by_item[item] = [r for r in raw if m.filter_inv_record(r, item)]
print('filtered onhand:', [r['onhandId'] for r in inv_by_item['03058754']])
assert [r['onhandId'] for r in inv_by_item['03058754']] == ['O1', 'O2'], 'filter wrong'

# 15 字段提取验证（真实字段名）
ext = m.extract_record(recs['03058754'][0])
print('extract:', ext)
assert ext[0] == 'O1'                       # onhandId
assert ext[1] == '3449'                     # entityId
assert ext[2] == 'GCW_BIN_PH'               # warehouse
assert ext[5] == '0006082300010E'           # contractNumber
assert ext[6] == '03058754'                 # item
assert ext[9] == '1'                        # actQty
assert ext[12] == 'xxx'                     # actPropertyTypeDesc
assert ext[13] == 'good'                    # qualityTypeDesc
assert ext[14] == '10'                      # goodsAge

# 占用分配
remain = {}
result = []
for (duid, desc, item, qty, tv) in device_rows:
    q = m._to_float(qty)
    chosen = None
    for rec in inv_by_item.get(item, []):
        onhand = m._cell_str(m._get_field(rec, *m.FIELD_ONHAND))
        contract = m._cell_str(m._get_field(rec, *m.FIELD_CONTRACT))
        if m._norm_contract(contract) not in exp:
            continue
        avi = m._to_float(m._get_field(rec, *m.FIELD_AVI))
        rem = remain.get(onhand, avi)
        if rem >= q:
            chosen = (onhand, contract)
            remain[onhand] = rem - q
            break
    if chosen:
        sub = exp.get(m._norm_contract(chosen[1]), "")
        result.append((duid, item, qty, chosen[1], sub))
    else:
        result.append((duid, item, qty, '无可用库存', ''))

print('result:', result)
assert result[0][3] == '0006082300010E' and result[0][4] == 'EXP', 'SL4078 应可用且 Sub=EXP'
assert result[1][3] == '0006082300010E' and result[1][4] == 'EXP', 'SL18 应可用且 Sub=EXP'

print('ALL TESTS PASSED (真实字段名)')
