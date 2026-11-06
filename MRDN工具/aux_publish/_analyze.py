# -*- coding: utf-8 -*-
import openpyxl
from collections import defaultdict

def norm(s):
    if s is None: return ""
    return "".join(ch for ch in str(s).lower() if ch.isalnum())

mat = r"D:\work_document\DOS工具开发\MRDN系统\dist\MRDN系统\data\works\_batch\26年10月20日第一次作业\02_物料清单制作\输出结果.xlsx"
wb = openpyxl.load_workbook(mat, data_only=True)
ws = wb.active
headers = [str(c.value) if c.value is not None else "" for c in ws[1]]
print("headers:", headers)
def col_of(*names):
    for i,h in enumerate(headers):
        if norm(h) in [norm(n) for n in names]:
            return i
    return -1
ci_du = col_of("DUID","DU ID","DU")
ci_item = col_of("itemcode","Item Code")
ci_qty = col_of("qty","quantity")
ci_type = col_of("type")
ci_fb = col_of("是否发包","发包")
print("cols: du=%d item=%d qty=%d type=%d 发包=%d" % (ci_du,ci_item,ci_qty,ci_type,ci_fb))

groups = defaultdict(list)
for r in ws.iter_rows(min_row=2, values_only=True):
    du = str(r[ci_du]).strip() if ci_du>=0 and r[ci_du] else ""
    if not du: continue
    fb = str(r[ci_fb]).strip() if ci_fb>=0 and r[ci_fb] else ""
    t = str(r[ci_type]).strip() if ci_type>=0 and r[ci_type] else ""
    item = r[ci_item]
    qty = r[ci_qty]
    groups[du].append((fb, t, item, qty))

print("\n=== 按 DUID 分组 ===")
for du, rows in groups.items():
    has_fb = any(r[0] != "" for r in rows)  # 是否有发包标签
    types = set(r[1] for r in rows)
    has_type23 = any(norm(r[1]) in ("type2","type3") for r in rows)
    print("\nDUID:", du)
    print("  有发包标签:", has_fb, "| type分布:", types, "| 有Type2/3:", has_type23)
    for r in rows:
        print("    fb=%-6s type=%-8s item=%-16s qty=%s" % (repr(r[0]), r[1], r[2], r[3]))
    # 判定空实例
    if (not has_fb) and has_type23:
        print("  >>> 需要空实例")
    else:
        print("  >>> 不需要空实例")
