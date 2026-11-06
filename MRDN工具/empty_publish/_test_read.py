# -*- coding: utf-8 -*-
import openpyxl, os, shutil
import importlib.util

spec = importlib.util.spec_from_file_location("ep", r"D:\work_document\python_doc\pythonProject\MRDN工具\empty_publish\MRDN空实例发布.py")
ep = importlib.util.module_from_spec(spec)
spec.loader.exec_module(ep)

# 1) 测试 read_empty_list 对真实空实例列表的解析
mapping, fallback, has_du_col = ep.read_empty_list(r"D:\work_document\python_doc\pythonProject\MRDN工具\empty_publish\空实例列表.xlsx")
print("mapping:", mapping)
print("fallback:", repr(fallback))
print("has_du_col:", has_du_col)

# 2) 测试 read_material 对真实物料清单
groups = ep.read_material(r"D:\work_document\DOS工具开发\MRDN系统\dist\MRDN系统\data\works\_batch\26年10月20日第一次作业\02_物料清单制作\输出结果.xlsx")
print("DUIDs:", list(groups.keys()))
for du, rows in groups.items():
    fb = [r for r in rows if r[3] != ""]
    t23 = [r for r in rows if r[2] in ("type2", "type3")]
    print(du, "行数", len(rows), "发包行", len(fb), "Type2/3行", len(t23))
empty = ep.find_empty_duids(groups)
print("空实例 DUID:", empty)
