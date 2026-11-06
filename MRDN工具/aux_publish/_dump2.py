# -*- coding: utf-8 -*-
import json

path = r"D:\work_document\DOS工具开发\MRDN系统\dist\MRDN系统\data\data.json"
with open(path, encoding="utf-8") as f:
    d = json.load(f)

print("=== checklist ===")
for f in d.get("flows", []):
    cl = f.get("checklist", [])
    print("checklist count:", len(cl))
    for c in cl:
        print("  ", c.get("id"), "|", c.get("name"), "| group=", c.get("group"), "| file=", c.get("file"))

print()
print("=== steps 详细（含 output.files 和 tool） ===")
for f in d.get("flows", []):
    for s in f.get("steps", []):
        print("-" * 50)
        print("step", s.get("seq"), s.get("name"), "id=", s.get("id"))
        for it in (s.get("input", {}) or {}).get("items", []):
            print("   item:", it.get("id"), it.get("name"), "kind=", it.get("kind"), "refId=", it.get("refId"), "refFile=", it.get("refFile"))
        t = s.get("tool")
        if t:
            print("   tool:", t.get("name"), "config=", t.get("config"))
        out = s.get("output", {}) or {}
        print("   output.type:", out.get("type"), "text:", (out.get("text") or "")[:60])
        for fo in out.get("files", []):
            print("     file:", fo.get("id"), fo.get("name"), fo.get("path"))
