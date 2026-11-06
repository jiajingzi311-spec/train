# -*- coding: utf-8 -*-
import sys
sys.stdout.reconfigure(encoding="utf-8", errors="replace")
import importlib.util

spec = importlib.util.spec_from_file_location(
    "mc", r"D:\work_document\python_doc\pythonProject\MRDN工具\material_check\MRDN物料清单核查.py")
m = importlib.util.module_from_spec(spec)
spec.loader.exec_module(m)

# 旧长表：DU1 itemA=2, itemB=5, itemC=1
old = {
    (m._norm_duid("DU1"), m._norm_code("11111111")): {"duid":"DU1","code":"11111111","qty":2,"desc":"A","type":"","date":None},
    (m._norm_duid("DU1"), m._norm_code("22222222")): {"duid":"DU1","code":"22222222","qty":5,"desc":"B","type":"","date":None},
    (m._norm_duid("DU1"), m._norm_code("33333333")): {"duid":"DU1","code":"33333333","qty":1,"desc":"C","type":"","date":None},
}
# 新长表：DU1 itemA=3, itemB=5, itemD=2
new = {
    (m._norm_duid("DU1"), m._norm_code("11111111")): {"duid":"DU1","code":"11111111","qty":3,"desc":"A","type":"","date":None},
    (m._norm_duid("DU1"), m._norm_code("22222222")): {"duid":"DU1","code":"22222222","qty":5,"desc":"B","type":"","date":None},
    (m._norm_duid("DU1"), m._norm_code("44444444")): {"duid":"DU1","code":"44444444","qty":2,"desc":"D","type":"","date":None},
}

rows = m.compare(old, new)
print("差异条数:", len(rows))
for r in rows:
    print("  %s | %s | 旧=%s 新=%s | %s" % (r["duid"], r["code"], r["old_qty"], r["new_qty"], r["label"]))

# 期望: 11111111 数量变化, 33333333 需剔除, 44444444 需新增
assert len(rows) == 3, "应为 3 条差异"
labels = {r["code"]: r["label"] for r in rows}
assert labels["11111111"] == "数量变化", labels
assert labels["33333333"] == "需剔除", labels
assert labels["44444444"] == "需新增", labels
print("PASS: 数量变化 / 需剔除 / 需新增 逻辑全部正确")
