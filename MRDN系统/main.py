# -*- coding: utf-8 -*-
"""
MRDN 系统 · 业务操作手册（桌面应用）
后端：数据持久化 + 工具执行 + Excel 导出
"""
import os
import re
import sys
import json
import copy
import shutil
import base64
import shlex
import threading
import uuid
import subprocess
import datetime
import time

import webview


def base_dir():
    """用户数据目录：开发时=脚本目录，打包后=exe所在目录"""
    if getattr(sys, 'frozen', False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))


def resource_dir():
    """静态资源目录（index.html）：打包后=临时解压目录"""
    if getattr(sys, 'frozen', False):
        return getattr(sys, '_MEIPASS', os.path.dirname(sys.executable))
    return os.path.dirname(os.path.abspath(__file__))


def ensure_webview_runtimes():
    """pywebview 的 edgechromium 后端导入时会遍历 win-arm64/win-x64/win-x86 三个 RID 目录，
    其中 win-arm64/native 在源包里是空目录，PyInstaller 不会打包空目录，
    导致打包后运行报 FileNotFoundError。这里在打包环境补建该空目录。"""
    if not getattr(sys, 'frozen', False):
        return
    try:
        base = getattr(sys, '_MEIPASS', None) or os.path.dirname(os.path.realpath(__file__))
        target = os.path.join(base, 'webview', 'lib', 'runtimes', 'win-arm64', 'native')
        os.makedirs(target, exist_ok=True)
    except Exception:
        pass


BASE = base_dir()
DATA_DIR = os.path.join(BASE, 'data')
TOOLS_DIR = os.path.join(BASE, 'tools')
SHOTS_DIR = os.path.join(DATA_DIR, 'shots')
EXPORT_DIR = os.path.join(DATA_DIR, '导出')
INPUTS_DIR = os.path.join(DATA_DIR, 'inputs')
OUTPUTS_DIR = os.path.join(DATA_DIR, 'outputs')
WORKS_DIR = os.path.join(DATA_DIR, 'works')
DATA_FILE = os.path.join(DATA_DIR, 'data.json')
EXCEL_FILE = os.path.join(DATA_DIR, '操作手册.xlsx')
DU_STATE_FILE = os.path.join(DATA_DIR, 'du_state.json')   # 独立的 DU 状态持久化文件（任务锚点完整清单 + 完成/未完成标记）

# ---- 版本更新 ----
APP_VERSION = "1.0.0"                                   # 当前应用版本（每次发布新版时修改）
DEFAULT_UPDATE_URL = "http://your-internal-server/MRDN系统/version.json"   # 最新版本信息文件默认地址（内网 HTTP）
UPDATE_CONFIG_FILE = os.path.join(BASE, 'update_config.json')               # exe 同层级可选配置文件，可覆盖默认地址


def _get_update_url():
    """读取更新地址：优先 exe 同层级 update_config.json 的 update_url，否则用默认地址。"""
    url = DEFAULT_UPDATE_URL
    try:
        if os.path.exists(UPDATE_CONFIG_FILE):
            with open(UPDATE_CONFIG_FILE, encoding='utf-8') as f:
                cfg = json.load(f)
            if isinstance(cfg, dict) and (cfg.get('update_url') or '').strip():
                url = cfg['update_url'].strip()
    except Exception:
        pass
    return url


def _version_gt(a, b):
    """语义化版本对比：a > b 返回 True。忽略非数字字符，按数字分段比较。"""
    def parts(v):
        nums = re.findall(r'\d+', str(v))
        return [int(x) for x in nums] if nums else [0]
    pa, pb = parts(a), parts(b)
    for i in range(max(len(pa), len(pb))):
        x = pa[i] if i < len(pa) else 0
        y = pb[i] if i < len(pb) else 0
        if x != y:
            return x > y
    return False


# 首次运行时的初始数据（后续以 data/data.json 为准）
DEFAULT_DATA = {
    "version": 1,
    "appName": "MRDN 系统",
    "tags": [
        {"id": "data", "name": "数据处理", "icon": "📊"},
        {"id": "web", "name": "网页操作", "icon": "🌐"},
        {"id": "notify", "name": "消息/邮件告知", "icon": "📧"}
    ],
    "flows": [
        {
            "id": "f1", "name": "订单数据导出", "icon": "📦", "color": "#5b6cff",
            "goal": "把订单系统数据导出 → 用工具清洗 → 再回传系统，一条龙搞定。",
            "checklist": [
                {"id": "C001", "name": "客户码表", "desc": "客户名称与编码对照表，录入前先备好。", "file": None},
                {"id": "C002", "name": "SKU 参照表", "desc": "商品编码与规格的参照表。", "file": None}
            ],
            "steps": [
                {"id": "N001", "seq": 1, "name": "登录订单系统",
                 "desc": "打开订单系统网址，输入账号密码，点「登录」。",
                 "input": {"type": "text", "text": "账号、密码"},
                 "output": {"type": "text", "text": "登录成功，进入订单首页"},
                 "note": "账号密码已存浏览器，直接点登录即可。",
                 "screenshot": None, "tool": None, "files": []},
                {"id": "N002", "seq": 2, "name": "筛选本月订单",
                 "desc": "1. 进入「订单查询」页\n2. 时间范围选「本月」\n3. 状态选「已完成」\n4. 点「查询」",
                 "input": {"type": "ref", "refId": "N001", "text": "登录态（引用 N001 的输出）"},
                 "output": {"type": "text", "text": "本月已完成订单列表（约 320 条）"},
                 "note": "时区选「北京时间」，否则会漏单。",
                 "screenshot": None, "tool": None, "files": []},
                {"id": "N003", "seq": 3, "name": "导出订单表",
                 "desc": "点右上角「导出 Excel」，选「全部字段」，下载到本地。",
                 "input": {"type": "ref", "refId": "N002", "text": "订单列表（引用 N002 的输出）"},
                 "output": {"type": "text", "text": "原始订单表 order_20261013.xlsx"},
                 "note": "文件名自动带日期，别改。",
                 "screenshot": None, "tool": None,
                 "files": [{"name": "order_20261013.xlsx"}]},
                {"id": "N004", "seq": 4, "name": "工具清洗数据",
                 "desc": "运行清洗工具，自动去重、格式化日期、剔除空行。",
                 "input": {"type": "ref", "refId": "N003", "text": "order_20261013.xlsx（引用 N003 的输出）"},
                 "output": {"type": "text", "text": "清洗后 cleaned_20261013.xlsx"},
                 "note": "工具在 tools 文件夹，运行前先确认输入文件路径正确。",
                 "screenshot": None,
                 "tool": {"name": "clean_data.exe"},
                 "files": [{"name": "cleaned_20261013.xlsx"}]},
                {"id": "N005", "seq": 5, "name": "回传清洗结果",
                 "desc": "回到订单系统，进入「批量导入」，上传清洗结果，点「确认导入」。",
                 "input": {"type": "ref", "refId": "N004", "text": "cleaned_20261013.xlsx（引用 N004 的输出）"},
                 "output": {"type": "text", "text": "系统提示「导入成功：320 条」"},
                 "note": "导入前会弹二次确认，看清条数再点确认。",
                 "screenshot": None, "tool": None, "files": []}
            ]
        },
        {
            "id": "f2", "name": "日报发送", "icon": "📊", "color": "#06b6d4",
            "goal": "每天汇总订单数据，整理成日报并发邮件给团队。",
            "steps": [
                {"id": "N006", "seq": 1, "name": "整理日报数据",
                 "desc": "汇总当天订单量、异常单，填进日报模板。",
                 "input": {"type": "text", "text": "当天订单记录"},
                 "output": {"type": "text", "text": "日报 Excel 一份"},
                 "note": "", "screenshot": None, "tool": None, "files": []},
                {"id": "N007", "seq": 2, "name": "发送日报邮件",
                 "desc": "用 Outlook 把日报发给团队群组。",
                 "input": {"type": "ref", "refId": "N006", "text": "日报 Excel（引用 N006 的输出）"},
                 "output": {"type": "text", "text": "邮件已发送"},
                 "note": "", "screenshot": None, "tool": None, "files": []}
            ]
        },
        {
            "id": "f3", "name": "客户信息维护", "icon": "👥", "color": "#10b981",
            "goal": "统一维护客户资料，新增或修改都有章可循。",
            "steps": [
                {"id": "N008", "seq": 1, "name": "新增客户",
                 "desc": "进入客户管理，点「新增」，填写客户名称、联系方式。",
                 "input": {"type": "text", "text": "客户名称、联系方式"},
                 "output": {"type": "text", "text": "客户建档成功"},
                 "note": "", "screenshot": None, "tool": None, "files": []}
            ]
        }
    ]
}


def decode_bytes(b):
    for enc in ("utf-8", "gbk", "utf-16"):
        try:
            return b.decode(enc)
        except Exception:
            continue
    return b.decode("utf-8", "replace")


class Api:
    def __init__(self):
        for d in (DATA_DIR, TOOLS_DIR, SHOTS_DIR, EXPORT_DIR, INPUTS_DIR, OUTPUTS_DIR, WORKS_DIR):
            os.makedirs(d, exist_ok=True)
        self._login_sessions = {}

    # ---- 版本更新（轻量版：检查 → 对比 → 提示下载） ----
    def check_update(self):
        """请求 version.json，对比当前版本，返回是否有更新及下载信息。"""
        result = {"ok": True, "current_version": APP_VERSION, "has_update": False,
                  "latest_version": "", "url": "", "notes": "", "md5": "", "error": ""}
        try:
            import urllib.request
            req = urllib.request.Request(_get_update_url(), headers={"User-Agent": "MRDN-System"})
            with urllib.request.urlopen(req, timeout=10) as resp:
                raw = resp.read().decode("utf-8")
            info = json.loads(raw)
            latest = str(info.get("version", "") or "").strip()
            url = str(info.get("url", "") or "").strip()
            notes = str(info.get("notes", "") or "").strip()
            md5 = str(info.get("md5", "") or "").strip()
            result["latest_version"] = latest
            result["url"] = url
            result["notes"] = notes
            result["md5"] = md5
            result["has_update"] = bool(latest and _version_gt(latest, APP_VERSION))
        except Exception as e:
            result["error"] = "无法连接更新服务器：%s" % (e,)
        return result

    # ---- 数据 ----
    def load_data(self):
        data = None
        if os.path.exists(DATA_FILE):
            try:
                with open(DATA_FILE, encoding="utf-8") as f:
                    data = json.load(f)
            except Exception:
                data = None
        if data is None:
            data = copy.deepcopy(DEFAULT_DATA)
        # 兼容旧数据：给每个流程补 progress 进度字段（已完成节点 id 列表）
        for flow in data.get("flows", []):
            if not isinstance(flow.get("progress"), dict):
                flow["progress"] = {"done": []}
            # 兼容旧数据：给每个流程补任务初始清单
            if not isinstance(flow.get("checklist"), list):
                flow["checklist"] = []
            # 兼容旧数据：节点输入由单选引用(type/refId)迁移为多选引用(refs)
            for s in flow.get("steps", []):
                inp = s.get("input") or {}
                if "refs" not in inp:
                    refs = []
                    text = inp.get("text") or ""
                    if inp.get("type") == "ref" and inp.get("refId"):
                        refs.append({"kind": "node", "refId": inp["refId"], "required": False})
                        text = ""
                    elif inp.get("type") == "checklist" and inp.get("refId"):
                        refs.append({"kind": "checklist", "refId": inp["refId"], "required": False})
                        text = ""
                    inp["refs"] = refs
                    inp["text"] = text
                    inp.pop("type", None)
                    inp.pop("refId", None)
                    s["input"] = inp
            # 兼容旧数据：输入从 {text,refs,file,fileRequired} 迁移为 {text, items:[{id,name,required,kind,file,refId,refLabel}]}
            for s in flow.get("steps", []):
                inp = s.get("input") or {}
                if "items" in inp:
                    continue
                items = []
                old_file = inp.get("file")
                file_required = bool(inp.get("fileRequired"))
                if old_file or file_required:
                    items.append({
                        "id": "in_file", "name": "上传文件", "required": file_required,
                        "kind": "file", "file": old_file, "refId": None, "refLabel": None,
                    })
                for r in inp.get("refs") or []:
                    rid = r.get("refId")
                    if r.get("kind") == "node":
                        tgt = next((x for x in flow.get("steps", []) if x.get("id") == rid), None)
                        items.append({
                            "id": "in_ref_" + str(rid), "name": (tgt.get("name") if tgt else "节点输出"),
                            "required": bool(r.get("required")), "kind": "node", "file": None,
                            "refId": rid, "refLabel": (tgt.get("name") if tgt else None),
                        })
                    elif r.get("kind") == "checklist":
                        ci = next((x for x in (flow.get("checklist") or []) if x.get("id") == rid), None)
                        items.append({
                            "id": "in_cl_" + str(rid), "name": (ci.get("name") if ci else "初始清单"),
                            "required": bool(r.get("required")), "kind": "checklist", "file": None,
                            "refId": rid, "refLabel": (ci.get("name") if ci else None),
                        })
                inp["items"] = items
                inp["text"] = inp.get("text") or ""
                inp.pop("file", None)
                inp.pop("fileRequired", None)
                inp.pop("refs", None)
                s["input"] = inp
            # 兼容旧数据：工具槽位迁移为 {id, itemId}（itemId=本节点输入项 id，仅本节点）
            for s in flow.get("steps", []):
                tool = s.get("tool") or {}
                cfg = tool.get("config")
                if not isinstance(cfg, list):
                    continue
                migrated = []
                for c in cfg:
                    if not isinstance(c, dict):
                        continue
                    if "itemId" in c and "ref" not in c:
                        migrated.append(c)  # 已是新格式
                        continue
                    # 旧 ref/source 格式都是跨节点或临时引用，无法自动映射到本节点输入项 → itemId=None
                    migrated.append({"id": c.get("id"), "itemId": None})
                tool["config"] = migrated
                s["tool"] = tool
            # 兼容旧数据：节点输出补 files 字段（真实输出文件，如输出模板）
            for s in flow.get("steps", []):
                out = s.get("output") or {}
                if not isinstance(out.get("files"), list):
                    out["files"] = []
                if not isinstance(out.get("results"), list):
                    out["results"] = []
                s["output"] = out
            # 兼容新数据：流程补批次(batches)与作业项(workItems)
            if not isinstance(flow.get("batches"), list):
                flow["batches"] = []
            if not isinstance(flow.get("workItems"), list):
                flow["workItems"] = []
            # 兼容新数据：补物料断点状态标签(materialStatusOptions)；任务锚点改由节点级 anchor 配置
            if not isinstance(flow.get("materialStatusOptions"), list):
                flow["materialStatusOptions"] = ["缺料", "缺余量", "等CRF", "无料", "等货", "其他"]
            # 兼容新数据：作业项补物料状态(materials)
            for w in flow.get("workItems", []):
                if not isinstance(w.get("materials"), list):
                    w["materials"] = []
        # 兼容旧数据：补节点标签列表
        if not isinstance(data.get("tags"), list) or not data.get("tags"):
            data["tags"] = copy.deepcopy(DEFAULT_DATA["tags"])
        with open(DATA_FILE, "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False, indent=2)
        self._write_excel(data)
        return data

    def save_data(self, data):
        try:
            with open(DATA_FILE, "w", encoding="utf-8") as f:
                json.dump(data, f, ensure_ascii=False, indent=2)
            self._write_excel(data)
            return {"ok": True}
        except Exception as e:
            return {"ok": False, "error": str(e)}

    # ---- DU 状态独立持久化（du_state.json）----
    def _load_du_state(self):
        if os.path.exists(DU_STATE_FILE):
            try:
                with open(DU_STATE_FILE, encoding="utf-8") as f:
                    return json.load(f)
            except Exception:
                pass
        return {}

    def _save_du_state(self, state):
        try:
            with open(DU_STATE_FILE, "w", encoding="utf-8") as f:
                json.dump(state, f, ensure_ascii=False, indent=2)
        except Exception:
            pass

    def _cache_du_state(self, flow_id, work_item_id, duid, materials):
        """把某个作业项的完整锚点物料清单（含状态）写入 du_state.json。"""
        state = self._load_du_state()
        key = flow_id + "|" + work_item_id
        state[key] = {
            "duid": duid,
            "materials": [
                {
                    "itemcode": m.get("itemcode"),
                    "desc": m.get("desc", "") or "",
                    "qty": m.get("qty", "") or "",
                    "type": m.get("type", "") or "",
                    "sitestype": m.get("sitestype", "") or "",
                    "project": m.get("project", "") or "",
                    "fromItemcode": m.get("fromItemcode", "") or "",
                    "fromType": m.get("fromType", "") or "",
                    "status": m.get("status", "") or "",
                    "remark": m.get("remark", "") or "",
                    "done": bool(m.get("done")),
                    "extra": m.get("extra") or {},
                }
                for m in materials if m.get("itemcode")
            ],
        }
        self._save_du_state(state)

    def clear_du_state(self, flow_id, work_item_ids):
        """清空 du_state.json 里指定作业项的锚点缓存（重跑锚点节点时调用，让清单刷新到重跑后的新结果）。"""
        try:
            state = self._load_du_state()
            keys = [str(flow_id) + "|" + str(wid) for wid in (work_item_ids or [])]
            changed = False
            for k in keys:
                if k in state:
                    del state[k]
                    changed = True
            if changed:
                self._save_du_state(state)
            return {"ok": True}
        except Exception as e:
            return {"ok": False, "error": str(e)}

    # ---- Excel ----
    def export_excel(self, data):
        try:
            fn = "操作手册_" + datetime.datetime.now().strftime("%Y%m%d_%H%M%S") + ".xlsx"
            path = os.path.join(EXPORT_DIR, fn)
            self._write_excel(data, path)
            return {"ok": True, "path": path}
        except Exception as e:
            return {"ok": False, "error": str(e)}

    def _write_excel(self, data, path=None):
        import openpyxl
        from openpyxl.styles import Font, PatternFill, Alignment
        from openpyxl.utils import get_column_letter
        path = path or EXCEL_FILE
        wb = openpyxl.Workbook()

        HDR_FILL = PatternFill("solid", fgColor="5B6CFF")
        HDR_FONT = Font(bold=True, color="FFFFFF")
        HDR_ALIGN = Alignment(horizontal="center", vertical="center", wrap_text=True)
        WRAP = Alignment(vertical="top", wrap_text=True)

        def style_header(ws):
            for c in ws[1]:
                c.font = HDR_FONT
                c.fill = HDR_FILL
                c.alignment = HDR_ALIGN

        def set_widths(ws, widths):
            for i, w in enumerate(widths, 1):
                ws.column_dimensions[get_column_letter(i)].width = w

        KIND_LABEL = {"file": "拖入文件", "checklist": "引用清单", "node": "引用节点输出", "desc": "仅描述"}

        def item_display_name(it):
            return it.get("name") or it.get("refLabel") or ""

        flows = data.get("flows", [])
        # 收集 (flow, step) 对，并建立节点索引（用于解析「引用节点输出」）
        pairs = []
        for flow in flows:
            for s in flow.get("steps", []):
                pairs.append((flow, s))
        step_by_id = {s.get("id"): s for _, s in pairs}

        def checklist_item(flow, cid):
            for c in (flow.get("checklist") or []):
                if c.get("id") == cid:
                    return c
            return None

        def resolve_node_file(step):
            """节点输出解析：运行结果 → 输出模板文件 → 旧文字文件名。返回 (文件名, 路径)"""
            if not step:
                return "", ""
            out = step.get("output") or {}
            res = out.get("results") or []
            if res:
                return (res[0].get("name") or ""), (res[0].get("rel") or "")
            files = out.get("files") or []
            if files:
                return (files[0].get("name") or ""), (files[0].get("path") or "")
            old = step.get("files") or []
            if old:
                return (old[0].get("name") or ""), ""
            return "", ""

        def item_file_info(flow, item):
            """输入项解析。返回 (引用对象名, 文件名, 文件路径)"""
            kind = item.get("kind") or ""
            if kind == "file":
                f = item.get("file") or {}
                return "", (f.get("name") or ""), (f.get("path") or "")
            if kind == "checklist":
                c = checklist_item(flow, item.get("refId"))
                ref = item.get("refLabel") or (c.get("name") if c else "") or ""
                f = (c or {}).get("file") or {}
                return ref, (f.get("name") or ""), (f.get("path") or "")
            if kind == "node":
                ref = item.get("refLabel") or ""
                tgt = step_by_id.get(item.get("refId"))
                fn, fp = resolve_node_file(tgt)
                return ref, fn, fp
            return "", "", ""

        def config_ref_label(step, item_id):
            """工具配置的「关联对象」名称：输出文件(out_前缀) → 本节点输入项 → 原样"""
            if not item_id:
                return "未配置"
            if str(item_id).startswith("out_"):
                out = step.get("output") or {}
                for f in (out.get("files") or []):
                    if f.get("id") == item_id:
                        return f.get("name") or item_id
                return item_id
            inp = step.get("input") or {}
            for it in (inp.get("items") or []):
                if it.get("id") == item_id:
                    return it.get("name") or item_id
            return item_id

        # ============ Sheet 1：节点总览 ============
        ws = wb.active
        ws.title = "节点总览"
        ws.append(["节点ID", "所属流程", "步骤序号", "节点名称", "节点简介",
                   "网址", "站点备注", "操作描述", "操作指导", "输入说明", "输入项",
                   "输出说明", "输出模板文件", "输出结果文件", "工具", "注意事项"])
        style_header(ws)
        for flow, s in pairs:
            inp = s.get("input") or {}
            out = s.get("output") or {}
            tool = s.get("tool") or {}
            items = inp.get("items") or []
            items_sum = "；".join(
                "%s(%s、%s)" % (
                    item_display_name(it),
                    KIND_LABEL.get(it.get("kind"), "拖入文件"),
                    "必选" if it.get("required") else "可选",
                ) for it in items
            )
            guide_sum = "\n".join(
                "%d. %s" % (i + 1, (st.get("text") or ""))
                for i, st in enumerate(s.get("steps") or [])
            )
            out_tpl = "、".join((f.get("name") or "") for f in (out.get("files") or []))
            out_res = "、".join((r.get("name") or "") for r in (out.get("results") or []))
            ws.append([
                s.get("id"), flow.get("name"), s.get("seq"), s.get("name"),
                s.get("brief") or "", s.get("url") or "", s.get("siteNote") or "",
                s.get("desc") or "", guide_sum, inp.get("text") or "", items_sum,
                out.get("text") or "", out_tpl, out_res,
                tool.get("name") or "", s.get("note") or "",
            ])
        set_widths(ws, [10, 14, 6, 14, 18, 30, 14, 40, 40, 20, 40, 20, 20, 20, 18, 24])
        for row in ws.iter_rows(min_row=2):
            for c in row:
                c.alignment = WRAP

        # ============ Sheet 2：操作指导 ============
        ws2 = wb.create_sheet("操作指导")
        ws2.append(["节点序号", "节点名称", "步骤号", "步骤文字", "截图文件"])
        style_header(ws2)
        for flow, s in pairs:
            for i, st in enumerate(s.get("steps") or []):
                shot = st.get("shot") or ""
                shot_name = os.path.basename(shot) if shot else ""
                ws2.append([s.get("seq"), s.get("name"), i + 1, st.get("text") or "", shot_name])
        set_widths(ws2, [8, 16, 8, 50, 30])
        for row in ws2.iter_rows(min_row=2):
            row[3].alignment = WRAP

        # ============ Sheet 3：输入项明细 ============
        ws3 = wb.create_sheet("输入项明细")
        ws3.append(["节点序号", "节点名称", "输入项名称", "类型", "是否必选", "引用对象", "文件名称", "文件路径"])
        style_header(ws3)
        for flow, s in pairs:
            inp = s.get("input") or {}
            for it in (inp.get("items") or []):
                ref, fn, fp = item_file_info(flow, it)
                ws3.append([
                    s.get("seq"), s.get("name"), item_display_name(it),
                    KIND_LABEL.get(it.get("kind"), "拖入文件"),
                    "必选" if it.get("required") else "可选",
                    ref, fn, fp,
                ])
        set_widths(ws3, [8, 16, 18, 12, 8, 18, 28, 40])

        # ============ Sheet 4：工具配置 ============
        ws4 = wb.create_sheet("工具配置")
        ws4.append(["节点序号", "节点名称", "工具名称", "参数名", "关联对象"])
        style_header(ws4)
        for flow, s in pairs:
            tool = s.get("tool") or {}
            for cfg in (tool.get("config") or []):
                ws4.append([
                    s.get("seq"), s.get("name"), tool.get("name") or "",
                    cfg.get("id") or "", config_ref_label(s, cfg.get("itemId")),
                ])
        set_widths(ws4, [8, 16, 24, 18, 24])

        # ============ Sheet 5：输出文件 ============
        ws5 = wb.create_sheet("输出文件")
        ws5.append(["节点序号", "节点名称", "文件类型", "文件名称", "文件路径"])
        style_header(ws5)
        for flow, s in pairs:
            out = s.get("output") or {}
            for f in (out.get("files") or []):
                ws5.append([s.get("seq"), s.get("name"), "模板文件", f.get("name") or "", f.get("path") or ""])
            for r in (out.get("results") or []):
                ws5.append([s.get("seq"), s.get("name"), "运行结果", r.get("name") or "", r.get("rel") or ""])
        set_widths(ws5, [8, 16, 10, 28, 40])

        # ============ Sheet 6：初始清单 ============
        ws6 = wb.create_sheet("初始清单")
        ws6.append(["所属流程", "清单项名称", "描述", "文件名称", "文件路径"])
        style_header(ws6)
        for flow in flows:
            for c in (flow.get("checklist") or []):
                f = c.get("file") or {}
                ws6.append([flow.get("name"), c.get("name") or "", c.get("desc") or "", f.get("name") or "", f.get("path") or ""])
        set_widths(ws6, [14, 20, 30, 28, 40])

        wb.save(path)

    # ---- 工具 ----
    def list_tools(self):
        try:
            if not os.path.isdir(TOOLS_DIR):
                return []
            return sorted(f for f in os.listdir(TOOLS_DIR) if f.lower().endswith(".exe"))
        except Exception:
            return []

    def get_tool_manifest(self, name):
        """读取工具声明文件 tools/<name>.tool.json，返回输入/输出契约。
        没有声明文件时返回 None（前端按无契约工具处理）。"""
        try:
            name = (name or "").strip()
            if not name:
                return None
            # 若传入的是 exe 文件名，声明文件与 exe 同名、后缀换成 .tool.json
            stem = os.path.splitext(name)[0]
            candidates = [
                os.path.join(TOOLS_DIR, stem + ".tool.json"),
                os.path.join(TOOLS_DIR, name + ".tool.json"),
            ]
            for p in candidates:
                if os.path.isfile(p):
                    with open(p, encoding="utf-8") as f:
                        m = json.load(f)
                    m["_file"] = p
                    return m
            return None
        except Exception:
            return None

    def resolve_file(self, name):
        """把「inputs/xx」「outputs/xx」相对路径或纯文件名解析成真实绝对路径。
        供前端把工具输入槽位对应文件转成可传给工具的绝对路径。"""
        try:
            name = (name or "").strip()
            if not name:
                return {"ok": False, "error": "空文件名"}
            candidates = []
            if name.startswith(("inputs/", "shots/", "outputs/")):
                candidates.append(os.path.join(DATA_DIR, name))
            else:
                for d in (BASE, DATA_DIR, EXPORT_DIR, INPUTS_DIR, TOOLS_DIR):
                    candidates.append(os.path.join(d, name))
            for c in candidates:
                if os.path.isfile(c):
                    return {"ok": True, "path": c}
            return {"ok": False, "error": "文件不存在：" + name}
        except Exception as e:
            return {"ok": False, "error": str(e)}

    def run_tool(self, name, args=None):
        exe = os.path.join(TOOLS_DIR, name)
        if not os.path.isfile(exe):
            return {"ok": False, "log": "工具不存在：" + name + "\n请确认已放到 tools 文件夹"}
        try:
            args = list(args or [])
            # 为本次运行创建独立输出目录，工具通过 --out_dir 把结果写到这里
            run_dir = os.path.join(
                OUTPUTS_DIR,
                os.path.splitext(name)[0] + "_" + datetime.datetime.now().strftime("%Y%m%d_%H%M%S_%f"))
            os.makedirs(run_dir, exist_ok=True)
            cmd = [exe] + args + ["--out_dir", run_dir]
            r = subprocess.run(cmd, cwd=TOOLS_DIR, capture_output=True, timeout=600)
            log = decode_bytes(r.stdout) + decode_bytes(r.stderr)
            if not log.strip():
                log = "（无输出）"
            # 收集本次运行产生的文件（供前端展示/预览）
            outputs = []
            for fn in sorted(os.listdir(run_dir)):
                fp = os.path.join(run_dir, fn)
                if os.path.isfile(fp):
                    outputs.append({"name": fn,
                                    "rel": "outputs/" + os.path.basename(run_dir) + "/" + fn})
            return {"ok": True, "log": log, "exit_code": r.returncode,
                    "outputs": outputs, "run_dir": run_dir}
        except subprocess.TimeoutExpired:
            return {"ok": False, "log": "运行超时（10 分钟）"}
        except Exception as e:
            return {"ok": False, "log": "运行出错：" + str(e)}

    # ---- 截图 ----
    def save_screenshot(self, data_url):
        try:
            _, b64 = data_url.split(",", 1)
            img = base64.b64decode(b64)
            fn = "shot_" + datetime.datetime.now().strftime("%Y%m%d_%H%M%S_%f") + ".jpg"
            with open(os.path.join(SHOTS_DIR, fn), "wb") as f:
                f.write(img)
            return {"ok": True, "path": "shots/" + fn}
        except Exception as e:
            return {"ok": False, "error": str(e)}

    def get_shot(self, rel_path):
        """返回图片的 base64 data URL（pywebview 的 NavigateToString 无法加载 file:// 资源，
        所以图片统一走后端转 base64 再内联显示）"""
        try:
            rel_path = (rel_path or "").strip()
            if not rel_path:
                return {"ok": False, "error": "空路径"}
            if rel_path.startswith(("inputs/", "shots/")):
                path = os.path.join(DATA_DIR, rel_path)
            else:
                path = os.path.join(BASE, rel_path)
            if not os.path.isfile(path):
                return {"ok": False, "error": "文件不存在"}
            with open(path, "rb") as f:
                b = f.read()
            ext = os.path.splitext(path)[1].lower()
            mime = {
                ".jpg": "image/jpeg", ".jpeg": "image/jpeg", ".png": "image/png",
                ".gif": "image/gif", ".bmp": "image/bmp", ".webp": "image/webp",
            }.get(ext, "image/jpeg")
            return {"ok": True, "data": "data:" + mime + ";base64," + base64.b64encode(b).decode("ascii")}
        except Exception as e:
            return {"ok": False, "error": str(e)}

    # ---- 输入文档 ----
    def save_input_file(self, filename, data_url):
        try:
            _, b64 = data_url.split(",", 1)
            raw = base64.b64decode(b64)
            safe = os.path.basename(filename or "") or ("input_" + datetime.datetime.now().strftime("%Y%m%d_%H%M%S"))
            path = os.path.join(INPUTS_DIR, safe)
            with open(path, "wb") as f:
                f.write(raw)
            return {"ok": True, "name": safe, "path": "inputs/" + safe}
        except Exception as e:
            return {"ok": False, "error": str(e)}

    def preview_file(self, name):
        """预览文件内容：Excel 转成表格、图片转 base64、文本直接返回。
        供前端在操作台下方做「数据预览」，替代直接打开本地文件。"""
        try:
            name = (name or "").strip()
            if not name:
                return {"ok": False, "error": "空文件名"}
            candidates = []
            if name.startswith(("inputs/", "shots/", "outputs/")):
                candidates.append(os.path.join(DATA_DIR, name))
            else:
                for d in (BASE, DATA_DIR, EXPORT_DIR, INPUTS_DIR, TOOLS_DIR):
                    candidates.append(os.path.join(d, name))
            target = next((c for c in candidates if os.path.isfile(c)), None)
            if not target:
                return {"ok": False, "error": "文件不存在：" + name}
            base_name = os.path.basename(target)
            ext = os.path.splitext(target)[1].lower()
            # Excel（新格式）
            if ext in (".xlsx", ".xlsm"):
                import openpyxl
                wb = openpyxl.load_workbook(target, read_only=True, data_only=True)
                sheets = []
                for ws in wb.worksheets[:3]:
                    rows = []
                    for i, row in enumerate(ws.iter_rows(values_only=True)):
                        if i >= 100:
                            rows.append(["……（数据过多，仅显示前 100 行）"])
                            break
                        rows.append([("" if v is None else str(v)) for v in row][:50])
                    sheets.append({"name": ws.title, "rows": rows})
                wb.close()
                return {"ok": True, "type": "excel", "name": base_name, "sheets": sheets}
            # 图片
            if ext in (".jpg", ".jpeg", ".png", ".gif", ".bmp", ".webp"):
                with open(target, "rb") as f:
                    b = f.read()
                mime = {
                    ".jpg": "image/jpeg", ".jpeg": "image/jpeg", ".png": "image/png",
                    ".gif": "image/gif", ".bmp": "image/bmp", ".webp": "image/webp",
                }.get(ext, "image/jpeg")
                return {"ok": True, "type": "image", "name": base_name,
                        "data": "data:" + mime + ";base64," + base64.b64encode(b).decode("ascii")}
            # 文本类
            if ext in (".txt", ".csv", ".log", ".md", ".json", ".py"):
                with open(target, "rb") as f:
                    b = f.read()
                return {"ok": True, "type": "text", "name": base_name,
                        "content": decode_bytes(b)}
            # 其他类型：仅提示
            return {"ok": True, "type": "unsupported", "name": base_name,
                    "hint": "该文件类型暂不支持预览（" + (ext or "未知") + "），可点右上角「打开本地文件」查看"}
        except Exception as e:
            return {"ok": False, "error": str(e)}

    @staticmethod
    def _cell_str(v):
        """把单元格值规整成字符串；float 整数去掉 .0（如 19.0 → 19）。"""
        if v is None:
            return ""
        if isinstance(v, float) and v.is_integer():
            return str(int(v))
        return str(v)

    def _read_table(self, target, ext):
        """读取表格文件（Excel/CSV），返回 (headers, rows)。
        headers 是列名列表；rows 是 [{列名: 值字符串}, ...]。非表格返回 (None, None)。"""
        if ext in (".xlsx", ".xlsm"):
            import openpyxl
            wb = openpyxl.load_workbook(target, read_only=True, data_only=True)
            ws = wb.worksheets[0]
            raw = list(ws.iter_rows(values_only=True))
            wb.close()
            if not raw:
                return None, None
            headers = [self._cell_str(h).strip() for h in raw[0]]
            rows = []
            for row in raw[1:]:
                if all(v is None or self._cell_str(v).strip() == "" for v in row):
                    continue
                d = {}
                for i, h in enumerate(headers):
                    d[h] = self._cell_str(row[i]) if i < len(row) else ""
                rows.append(d)
            return headers, rows
        if ext == ".csv":
            import csv
            text = None
            for enc in ("utf-8-sig", "gbk", "utf-8"):
                try:
                    with open(target, "r", encoding=enc, newline="") as f:
                        text = f.read()
                    break
                except UnicodeDecodeError:
                    continue
            if text is None:
                with open(target, "r", encoding="utf-8", errors="replace", newline="") as f:
                    text = f.read()
            raw = list(csv.reader(text.splitlines()))
            if not raw:
                return None, None
            headers = [h.strip() for h in raw[0]]
            rows = []
            for row in raw[1:]:
                if not row or all(not c.strip() for c in row):
                    continue
                d = {}
                for i, h in enumerate(headers):
                    d[h] = row[i].strip() if i < len(row) else ""
                rows.append(d)
            return headers, rows
        return None, None

    def build_message(self, rel_path, template, filter_str):
        """根据消息模板 + 筛选条件，从工具输出的表格里生成可粘贴消息。
        rel_path: 结果文件相对路径；template: 含 {列名} 的消息模板；filter_str: 多行「列名=值」精确匹配（AND）。"""
        try:
            rel_path = (rel_path or "").strip()
            template = (template or "").strip()
            filter_str = (filter_str or "").strip()
            if not rel_path:
                return {"ok": False, "error": "无结果文件"}
            if not template:
                return {"ok": False, "error": "未配置消息模板"}
            candidates = []
            if rel_path.startswith(("inputs/", "shots/", "outputs/")):
                candidates.append(os.path.join(DATA_DIR, rel_path))
            else:
                for d in (BASE, DATA_DIR, EXPORT_DIR, INPUTS_DIR, TOOLS_DIR):
                    candidates.append(os.path.join(d, rel_path))
            target = next((c for c in candidates if os.path.isfile(c)), None)
            if not target:
                return {"ok": False, "error": "文件不存在：" + rel_path}
            ext = os.path.splitext(target)[1].lower()
            headers, rows = self._read_table(target, ext)
            if headers is None:
                return {"ok": False, "error": "结果不是表格文件（仅支持 Excel/CSV）"}
            # 解析筛选条件（多行「列名=值」，AND 精确匹配）
            conds = []
            for line in filter_str.splitlines():
                line = line.strip()
                if not line:
                    continue
                if "=" in line:
                    k, v = line.split("=", 1)
                    conds.append((k.strip(), v.strip()))
            matched = []
            for r in rows:
                ok = True
                for k, v in conds:
                    if self._cell_str(r.get(k, "")) != v:
                        ok = False
                        break
                if ok:
                    matched.append(r)

            def fill_row(row):
                def repl(m):
                    key = m.group(1).strip()
                    return self._cell_str(row.get(key, ""))
                return re.sub(r"\{([^{}]+)\}", repl, template)

            lines = [fill_row(r) for r in matched]
            return {"ok": True, "message": "\n".join(lines), "matched": len(matched),
                    "headers": headers}
        except Exception as e:
            return {"ok": False, "error": str(e)}

    def open_file(self, name):
        try:
            name = (name or "").strip()
            if not name:
                return {"ok": False, "error": "空文件名"}
            candidates = []
            if name.startswith(("inputs/", "shots/", "outputs/")):
                candidates.append(os.path.join(DATA_DIR, name))
            else:
                for d in (BASE, DATA_DIR, EXPORT_DIR, INPUTS_DIR, TOOLS_DIR):
                    candidates.append(os.path.join(d, name))
            target = next((c for c in candidates if os.path.isfile(c)), None)
            if not target:
                return {"ok": False, "error": "文件不存在：" + name}
            os.startfile(target)
            return {"ok": True}
        except Exception as e:
            return {"ok": False, "error": str(e)}

    def open_folder(self, name):
        """解析文件路径，打开其所在的文件夹（而不是直接打开文件本身）。"""
        try:
            name = (name or "").strip()
            if not name:
                return {"ok": False, "error": "空文件名"}
            candidates = []
            if name.startswith(("inputs/", "shots/", "outputs/")):
                candidates.append(os.path.join(DATA_DIR, name))
            else:
                for d in (BASE, DATA_DIR, EXPORT_DIR, INPUTS_DIR, TOOLS_DIR):
                    candidates.append(os.path.join(d, name))
            target = next((c for c in candidates if os.path.isfile(c)), None)
            if not target:
                return {"ok": False, "error": "文件不存在：" + name}
            os.startfile(os.path.dirname(target))
            return {"ok": True}
        except Exception as e:
            return {"ok": False, "error": str(e)}

    def open_url(self, url):
        try:
            url = (url or "").strip()
            if not url:
                return {"ok": False, "error": "空网址"}
            if not (url.startswith("http://") or url.startswith("https://")):
                url = "https://" + url
            import webbrowser
            webbrowser.open(url)
            return {"ok": True}
        except Exception as e:
            return {"ok": False, "error": str(e)}

    # ---- 打开目录 ----
    def open_data_dir(self):
        try:
            os.startfile(DATA_DIR)
            return {"ok": True}
        except Exception as e:
            return {"ok": False, "error": str(e)}

    def open_tools_dir(self):
        try:
            os.makedirs(TOOLS_DIR, exist_ok=True)
            os.startfile(TOOLS_DIR)
            return {"ok": True}
        except Exception as e:
            return {"ok": False, "error": str(e)}

    # ---- 作业清单（批次 / DUID 作业项）----
    def _abs_path(self, name):
        """把「inputs/xx」「outputs/xx」等相对路径或纯文件名解析为绝对路径，找不到返回 None。"""
        name = (name or "").strip()
        if not name:
            return None
        if name.startswith(("inputs/", "shots/", "outputs/", "works/")):
            p = os.path.join(DATA_DIR, name)
            return p if os.path.isfile(p) else None
        for d in (BASE, DATA_DIR, EXPORT_DIR, INPUTS_DIR, TOOLS_DIR, OUTPUTS_DIR):
            p = os.path.join(d, name)
            if os.path.isfile(p):
                return p
        return None

    def _safe_name(self, s):
        """清理 DUID / 节点名里的非法文件名字符。"""
        s = re.sub(r'[\\/:*?"<>|]', "_", (s or "").strip())
        return s or "未命名"

    def _batch_name(self, batch_id):
        """根据批次 id 查批次名（用于归档目录命名）。
        有批次名用批次名，无则默认「年月日时分 + 批次」。"""
        try:
            data = self.load_data()
            for fl in data.get("flows", []):
                for b in fl.get("batches", []):
                    if b.get("id") == batch_id:
                        name = (b.get("name") or "").strip()
                        if name:
                            return name
        except Exception:
            pass
        return datetime.datetime.now().strftime("%Y-%m-%d %H:%M") + "批次"

    def make_duid_list(self, flow_id, batch_id, step_id=None):
        """生成批次「当前 DUID 清单」文本文件（每行一个 DUID），返回绝对路径。
        供 batch 节点跑工具时，把「当前批次剩余 DUID」动态传给工具的 --du_list 参数。
        拆出的 DUID（batchId=null）自动排除，只统计仍在批次内的 DUID。
        step_id 传入时，进一步排除已单独完成该节点的 DUID（避免批量作业重复处理）。"""
        try:
            data = self.load_data()
            flow = next((f for f in data.get("flows", []) if f.get("id") == flow_id), None)
            if not flow:
                return {"ok": False, "error": "流程不存在"}
            items = [w for w in flow.get("workItems", [])
                     if w.get("batchId") == batch_id and w.get("duid")]
            if step_id:
                items = [w for w in items
                         if step_id not in (w.get("doneStepIds") or [])]
            duids = [w.get("duid") for w in items]
            if not duids:
                return {"ok": False, "error": "该节点没有未完成的 DUID（可能都已单独完成，无需重复运行）"}
            os.makedirs(INPUTS_DIR, exist_ok=True)
            fn = "_duidlist_" + str(batch_id) + "_" + datetime.datetime.now().strftime("%Y%m%d_%H%M%S_%f") + ".txt"
            path = os.path.join(INPUTS_DIR, fn)
            with open(path, "w", encoding="utf-8") as f:
                f.write("\n".join(duids) + "\n")
            return {"ok": True, "path": path, "count": len(duids), "duids": duids}
        except Exception as e:
            return {"ok": False, "error": str(e)}

    def make_single_duid_list(self, duid):
        """生成「仅当前 DUID」的清单文件（拆出 DUID 单独作业时用），返回绝对路径。
        供 single 节点跑工具时，把当前 DUID 动态传给工具的 --du_list 参数。"""
        try:
            duid = (duid or "").strip()
            if not duid:
                return {"ok": False, "error": "DUID 为空"}
            os.makedirs(INPUTS_DIR, exist_ok=True)
            fn = "_duidlist_single_" + datetime.datetime.now().strftime("%Y%m%d_%H%M%S_%f") + ".txt"
            path = os.path.join(INPUTS_DIR, fn)
            with open(path, "w", encoding="utf-8") as f:
                f.write(duid + "\n")
            return {"ok": True, "path": path, "count": 1, "duids": [duid]}
        except Exception as e:
            return {"ok": False, "error": str(e)}

    def make_duid_list_file(self, duids):
        """把给定的 DUID 数组写成清单文件（每行一个），返回绝对路径。
        供「作业准备」模块跑工具时，把粘贴解析出的 DUID 动态传给工具的 --du_list 参数。"""
        try:
            duids = [str(d).strip() for d in (duids or []) if str(d).strip()]
            if not duids:
                return {"ok": False, "error": "没有 DUID，请先粘贴并解析"}
            os.makedirs(INPUTS_DIR, exist_ok=True)
            fn = "_duidlist_prep_" + datetime.datetime.now().strftime("%Y%m%d_%H%M%S_%f") + ".txt"
            path = os.path.join(INPUTS_DIR, fn)
            with open(path, "w", encoding="utf-8") as f:
                f.write("\n".join(duids) + "\n")
            return {"ok": True, "path": path, "count": len(duids), "duids": duids}
        except Exception as e:
            return {"ok": False, "error": str(e)}

    def make_itemcode_list(self, itemcodes):
        """把给定的 itemcode 数组写成清单文件（每行一个），返回绝对路径。
        供任务锚点「查询」操作把勾选的 itemcode 动态传给查询工具的 --itemcodes 参数。"""
        try:
            itemcodes = [str(x).strip() for x in (itemcodes or []) if str(x).strip()]
            if not itemcodes:
                return {"ok": False, "error": "没有勾选的 itemcode"}
            os.makedirs(INPUTS_DIR, exist_ok=True)
            fn = "_itemcodelist_" + datetime.datetime.now().strftime("%Y%m%d_%H%M%S_%f") + ".txt"
            path = os.path.join(INPUTS_DIR, fn)
            with open(path, "w", encoding="utf-8") as f:
                f.write("\n".join(itemcodes) + "\n")
            return {"ok": True, "path": path, "count": len(itemcodes), "itemcodes": itemcodes}
        except Exception as e:
            return {"ok": False, "error": str(e)}

    @staticmethod
    def _norm_col(s):
        """列名规范化：去掉空格/下划线/连字符/点，转小写，用于容错匹配。"""
        return re.sub(r"[\s_\-\.]", "", str(s or "")).lower()

    def _find_col(self, headers, keywords):
        """在表头里找列索引（容错匹配：先精确匹配，再子串匹配）。找不到返回 None。"""
        nkws = []
        for kw in keywords:
            nkw = self._norm_col(kw)
            if nkw:
                nkws.append(nkw)
        # 1) 精确匹配优先（避免 "type" 误匹配 "SiteStype" 等包含关系）
        for nkw in nkws:
            for i, h in enumerate(headers):
                if h and self._norm_col(h) == nkw:
                    return i
        # 2) 子串匹配兜底
        for nkw in nkws:
            for i, h in enumerate(headers):
                if h and nkw in self._norm_col(h):
                    return i
        return None

    def _resolve_step_output(self, step, file_name):
        """解析节点的输出文件为绝对路径（运行结果 → 输出文件模板 → 旧文字文件名）。
        file_name 指定时取匹配文件，否则取第一个；找不到返回 None。"""
        if not isinstance(step, dict):
            return None
        out = []
        output = step.get("output") or {}
        for r in (output.get("results") or []):
            if r and r.get("name"):
                out.append({"name": r["name"], "rel": r.get("rel") or r["name"]})
        for f in (output.get("files") or []):
            if f and f.get("name"):
                out.append({"name": f["name"], "rel": f.get("path") or f["name"]})
        for f in (step.get("files") or []):
            if f and f.get("name"):
                out.append({"name": f["name"], "rel": f.get("path") or f["name"]})
        seen = set()
        dedup = []
        for x in out:
            if x["name"] not in seen:
                seen.add(x["name"])
                dedup.append(x)
        if not dedup:
            return None
        if file_name:
            hit = next((x for x in dedup if x["name"] == file_name), None)
            if not hit:
                return None
            return self._abs_path(hit["rel"])
        return self._abs_path(dedup[0]["rel"])

    def get_anchor_data(self, flow_id, work_item_id):
        """读取任务锚点节点的输出文件（物料清单长表），过滤当前 DU 的物料，合并已保存状态，返回物料列表。"""
        try:
            data = self.load_data()
            flow = next((f for f in data.get("flows", []) if f.get("id") == flow_id), None)
            if not flow:
                return {"ok": False, "error": "流程不存在"}
            w = next((x for x in flow.get("workItems", []) if x.get("id") == work_item_id), None)
            if not w:
                return {"ok": False, "error": "作业项不存在"}
            duid = (w.get("duid") or "").strip()
            if not duid:
                return {"ok": False, "error": "当前作业无 DUID"}
            anchor_step = next((s for s in flow.get("steps", []) if s.get("anchor")), None)
            if not anchor_step:
                return {"ok": False, "notConfigured": True, "error": "未配置任务锚点"}
            anchor_cfg = anchor_step.get("anchor") or {}
            anchor_file = anchor_cfg.get("file") or ""
            target = self._resolve_step_output(anchor_step, anchor_file)

            # 完整物料清单（itemcode/desc/qty/type/sitestype/project）优先来自锚点输出文件；
            # 若输出文件丢失/不可读，则从独立持久化文件 du_state.json 恢复，保证标注不丢。
            materials = []
            all_columns = []
            if target:
                ext = os.path.splitext(target)[1].lower()
                headers, rows = self._read_table(target, ext)
                if headers is not None:
                    for _h in headers:
                        if _h not in all_columns:
                            all_columns.append(_h)
                    col_duid = self._find_col(headers, ["duid", "du_id", "duid2", "du", "du编号"])
                    col_item = self._find_col(headers, ["itemcode", "item_code", "item code", "code", "物料编码"])
                    if col_duid is not None and col_item is not None:
                        col_desc = self._find_col(headers, ["描述", "desc", "description", "物料描述"])
                        col_qty = self._find_col(headers, ["qty", "数量", "quantity", "new数量", "旧数量"])
                        col_type = self._find_col(headers, ["type", "类型", "物料类型"])
                        col_st = self._find_col(headers, ["sitestype", "site_type", "站点类型", "站点"])
                        col_proj = self._find_col(headers, ["可用项目合同", "项目类型", "project", "projectcode"])
                        seen = set()
                        for d in rows:
                            if (d.get(headers[col_duid]) or "").strip() != duid:
                                continue
                            code = (d.get(headers[col_item]) or "").strip()
                            if not code or code in seen:
                                continue
                            seen.add(code)
                            extra = {}
                            for _h in headers:
                                _v = d.get(_h)
                                if _v is not None:
                                    extra[_h] = str(_v).strip()
                            materials.append({
                                "itemcode": code,
                                "desc": (d.get(headers[col_desc]) if col_desc is not None else "") or "",
                                "qty": (d.get(headers[col_qty]) if col_qty is not None else "") or "",
                                "type": (d.get(headers[col_type]) if col_type is not None else "") or "",
                                "sitestype": (d.get(headers[col_st]) if col_st is not None else "") or "",
                                "project": (d.get(headers[col_proj]) if col_proj is not None else "") or "",
                                "extra": extra,
                            })
            if not materials:
                cached = (self._load_du_state().get(flow_id + "|" + work_item_id) or {})
                for m in (cached.get("materials") or []):
                    if m.get("itemcode"):
                        materials.append({
                            "itemcode": m.get("itemcode"),
                            "desc": m.get("desc", "") or "",
                            "qty": m.get("qty", "") or "",
                            "type": m.get("type", "") or "",
                            "sitestype": m.get("sitestype", "") or "",
                            "project": m.get("project", "") or "",
                            "fromItemcode": m.get("fromItemcode", "") or "",
                            "fromType": m.get("fromType", "") or "",
                            "extra": m.get("extra") or {},
                        })
            if not materials:
                return {"ok": False, "error": "任务锚点输出文件不存在（节点「" + (anchor_step.get("name") or "") + "」），且暂无本地缓存，请先运行一次锚点节点"}
            # 应用物料替换（重跑节点时会清除 replacements）
            replacements = [r for r in (w.get("replacements") or []) if isinstance(r, dict) and r.get("from")]
            if replacements:
                repl_map = {r.get("from"): r for r in replacements}
                for m in materials:
                    r = repl_map.get(m["itemcode"])
                    if r:
                        m["fromItemcode"] = m["itemcode"]
                        m["fromType"] = m.get("type", "") or ""
                        m["itemcode"] = (r.get("to") or m["itemcode"]).strip() or m["itemcode"]
                        if r.get("toQty"):
                            m["qty"] = r.get("toQty")
            # 合并已保存状态
            saved = {m.get("itemcode"): m for m in (w.get("materials") or []) if isinstance(m, dict) and m.get("itemcode")}
            for m in materials:
                sm = saved.get(m["itemcode"]) or {}
                m["status"] = sm.get("status", "") or ""
                m["remark"] = sm.get("remark", "") or ""
                m["done"] = bool(sm.get("done"))
            # 缓存完整清单（含状态）到 du_state.json，供下次打开/输出文件丢失时恢复
            self._cache_du_state(flow_id, work_item_id, duid, materials)
            # 可配置列：文件原始列 + 系统虚拟列（替换前 itemcode / 替换前 type）
            avail_cols = list(all_columns)
            for _v in ["fromItemcode", "fromType"]:
                if _v not in avail_cols:
                    avail_cols.append(_v)
            return {"ok": True, "duid": duid, "materials": materials,
                    "allColumns": avail_cols,
                    "statusOptions": flow.get("materialStatusOptions") or ["缺料", "缺余量", "等CRF", "无料", "等货", "其他"],
                    "anchorName": (anchor_step.get("name") or "") or ""}
        except Exception as e:
            return {"ok": False, "error": str(e)}

    def refresh_material_status(self, flow_id, work_item_id):
        """刷新物料状态（占位实现）。后续接入物料状态查询脚本后，在此拉取最新状态并写回 work item 的 materials。"""
        try:
            data = self.load_data()
            flow = next((f for f in data.get("flows", []) if f.get("id") == flow_id), None)
            if not flow:
                return {"ok": False, "message": "流程不存在"}
            w = next((x for x in flow.get("workItems", []) if x.get("id") == work_item_id), None)
            if not w:
                return {"ok": False, "message": "作业项不存在"}
            # TODO: 接入物料状态查询脚本，更新 w["materials"] 各物料的状态后调用 self.save_data(data)
            return {"ok": True, "message": "物料状态已刷新（占位：尚未接入查询脚本，未实际变更）"}
        except Exception as e:
            return {"ok": False, "message": str(e)}

    def get_anchor_data_batch(self, flow_id, batch_id):
        """读取任务锚点节点输出文件（物料长表），过滤当前批次所有 DU 的物料，返回物料列表（每项标注 duid）。"""
        try:
            data = self.load_data()
            flow = next((f for f in data.get("flows", []) if f.get("id") == flow_id), None)
            if not flow:
                return {"ok": False, "error": "流程不存在"}
            items = [x for x in flow.get("workItems", []) if x.get("batchId") == batch_id]
            if not items:
                return {"ok": False, "error": "批次无作业项"}
            duid_set = {(w.get("duid") or "").strip() for w in items if (w.get("duid") or "").strip()}
            if not duid_set:
                return {"ok": False, "error": "批次作业项无 DUID"}
            anchor_step = next((s for s in flow.get("steps", []) if s.get("anchor")), None)
            if not anchor_step:
                return {"ok": False, "notConfigured": True, "error": "未配置任务锚点"}
            anchor_cfg = anchor_step.get("anchor") or {}
            anchor_file = anchor_cfg.get("file") or ""
            target = self._resolve_step_output(anchor_step, anchor_file)

            # 批量下每个 DU 的物料清单优先从独立缓存 du_state.json 恢复
            # （锚点输出文件只有最近一次运行的单个 DU 结果，无法覆盖整批 DU）。
            state = self._load_du_state()
            duid_to_wid = {w.get("duid"): w.get("id") for w in items}
            materials = []
            cached_duids = set()
            for w in items:
                duid = (w.get("duid") or "").strip()
                cached = state.get(flow_id + "|" + w.get("id")) or {}
                cmats = cached.get("materials") or []
                if not cmats:
                    continue
                cached_duids.add(duid)
                for m in cmats:
                    if not m.get("itemcode"):
                        continue
                    materials.append({
                        "wid": w.get("id"),
                        "duid": duid,
                        "itemcode": m.get("itemcode"),
                        "desc": m.get("desc", "") or "",
                        "qty": m.get("qty", "") or "",
                        "type": m.get("type", "") or "",
                        "sitestype": m.get("sitestype", "") or "",
                        "project": m.get("project", "") or "",
                        "fromItemcode": m.get("fromItemcode", "") or "",
                        "fromType": m.get("fromType", "") or "",
                        "status": m.get("status", "") or "",
                        "remark": m.get("remark", "") or "",
                        "done": bool(m.get("done")),
                        "extra": m.get("extra") or {},
                    })

            # 对没有缓存的 DU，从锚点输出文件读取补充（也用于收集可选列）
            all_columns = []
            if target:
                ext = os.path.splitext(target)[1].lower()
                headers, rows = self._read_table(target, ext)
                if headers is not None:
                    for _h in headers:
                        if _h not in all_columns:
                            all_columns.append(_h)
                    col_duid = self._find_col(headers, ["duid", "du_id", "duid2", "du", "du编号"])
                    col_item = self._find_col(headers, ["itemcode", "item_code", "item code", "code", "物料编码"])
                    if col_duid is not None and col_item is not None:
                        col_desc = self._find_col(headers, ["描述", "desc", "description", "物料描述"])
                        col_qty = self._find_col(headers, ["qty", "数量", "quantity", "new数量", "旧数量"])
                        col_type = self._find_col(headers, ["type", "类型", "物料类型"])
                        col_st = self._find_col(headers, ["sitestype", "site_type", "站点类型", "站点"])
                        col_proj = self._find_col(headers, ["可用项目合同", "项目类型", "project", "projectcode"])
                        seen = set()
                        for d in rows:
                            duid = (d.get(headers[col_duid]) or "").strip()
                            if duid not in duid_set or duid in cached_duids:
                                continue
                            code = (d.get(headers[col_item]) or "").strip()
                            if not code or (duid, code) in seen:
                                continue
                            seen.add((duid, code))
                            extra = {}
                            for ci, h in enumerate(headers):
                                if ci in (col_duid, col_item, col_desc, col_qty, col_type, col_st, col_proj):
                                    continue
                                extra[h] = (d.get(h) or "") if ci < len(d) else ""
                            materials.append({
                                "wid": duid_to_wid.get(duid),
                                "duid": duid,
                                "itemcode": code,
                                "desc": (d.get(headers[col_desc]) if col_desc is not None else "") or "",
                                "qty": (d.get(headers[col_qty]) if col_qty is not None else "") or "",
                                "type": (d.get(headers[col_type]) if col_type is not None else "") or "",
                                "sitestype": (d.get(headers[col_st]) if col_st is not None else "") or "",
                                "project": (d.get(headers[col_proj]) if col_proj is not None else "") or "",
                                "extra": extra,
                            })

            if not materials:
                return {"ok": False, "error": "任务锚点输出文件不存在（节点「" + (anchor_step.get("name") or "") + "」），且暂无本地缓存，请先运行一次锚点节点"}

            # 应用物料替换（批量下每个 DU 的 replacements：from→to，与单 DU 逻辑一致）
            wid_repl = {}
            for w in items:
                repls = [r for r in (w.get("replacements") or []) if isinstance(r, dict) and r.get("from")]
                if repls:
                    wid_repl[w.get("id")] = {r.get("from"): r for r in repls}
            if wid_repl:
                for m in materials:
                    r = (wid_repl.get(m.get("wid")) or {}).get(m.get("itemcode"))
                    if r:
                        m["fromItemcode"] = m.get("itemcode")
                        m["fromType"] = m.get("type", "") or ""
                        m["itemcode"] = (r.get("to") or m.get("itemcode")).strip() or m.get("itemcode")
                        if r.get("toQty"):
                            m["qty"] = r.get("toQty")
                        # 替换后状态以作业项保存的为准，清掉 from 的旧状态避免残留
                        m["status"] = ""
                        m["remark"] = ""
                        m["done"] = False

            # 合并作业项最新已保存状态（workItems.materials 为用户最近标注，优先覆盖缓存）
            saved_by_duid = {}
            for w in items:
                duid = (w.get("duid") or "").strip()
                saved_by_duid[duid] = {m.get("itemcode"): m for m in (w.get("materials") or []) if isinstance(m, dict) and m.get("itemcode")}
            for m in materials:
                sm = (saved_by_duid.get(m.get("duid") or "") or {}).get(m.get("itemcode")) or {}
                if sm:
                    m["status"] = sm.get("status", "") or m.get("status", "")
                    m["remark"] = sm.get("remark", "") or m.get("remark", "")
                    m["done"] = bool(sm.get("done"))
            avail_cols = list(all_columns)
            for _v in ["fromItemcode", "fromType"]:
                if _v not in avail_cols:
                    avail_cols.append(_v)
            return {"ok": True, "batch": True, "materials": materials,
                    "allColumns": avail_cols,
                    "statusOptions": flow.get("materialStatusOptions") or ["缺料", "缺余量", "等CRF", "无料", "等货", "其他"],
                    "anchorName": (anchor_step.get("name") or "") or ""}
        except Exception as e:
            return {"ok": False, "error": str(e)}

    def get_shortage_data(self, flow_id):
        """缺料清单：统计所有未完成 DU 下未完成的物料，按 itemcode 聚合（跨 DU 维度）。"""
        try:
            data = self.load_data()
            flow = next((f for f in data.get("flows", []) if f.get("id") == flow_id), None)
            if not flow:
                return {"ok": False, "error": "流程不存在"}
            steps = flow.get("steps", []) or []
            step_ids = [s.get("id") for s in steps if s.get("id")]

            def _unfinished(w):
                if not step_ids:
                    return False
                done = set(w.get("doneStepIds") or [])
                return any(sid not in done for sid in step_ids)

            items = [x for x in (flow.get("workItems", []) or []) if _unfinished(x)]
            if not items:
                return {"ok": True, "rows": []}
            duid_set = {(w.get("duid") or "").strip() for w in items if (w.get("duid") or "").strip()}
            anchor_step = next((s for s in steps if s.get("anchor")), None)
            if not anchor_step:
                return {"ok": False, "notConfigured": True, "error": "未配置任务锚点"}
            anchor_cfg = anchor_step.get("anchor") or {}
            anchor_file = anchor_cfg.get("file") or ""
            target = self._resolve_step_output(anchor_step, anchor_file)

            state = self._load_du_state()
            duid_to_wid = {w.get("duid"): w.get("id") for w in items}
            materials = []
            cached_duids = set()
            for w in items:
                duid = (w.get("duid") or "").strip()
                cached = state.get(flow_id + "|" + w.get("id")) or {}
                cmats = cached.get("materials") or []
                if not cmats:
                    continue
                cached_duids.add(duid)
                for m in cmats:
                    if not m.get("itemcode"):
                        continue
                    materials.append({
                        "wid": w.get("id"),
                        "duid": duid,
                        "itemcode": m.get("itemcode"),
                        "desc": m.get("desc", "") or "",
                        "qty": m.get("qty", "") or "",
                        "type": m.get("type", "") or "",
                        "status": m.get("status", "") or "",
                        "remark": m.get("remark", "") or "",
                        "done": bool(m.get("done")),
                    })

            # 锚点输出文件补充未缓存的 DU
            if target:
                ext = os.path.splitext(target)[1].lower()
                headers, rows = self._read_table(target, ext)
                if headers is not None:
                    col_duid = self._find_col(headers, ["duid", "du_id", "duid2", "du", "du编号"])
                    col_item = self._find_col(headers, ["itemcode", "item_code", "item code", "code", "物料编码"])
                    if col_duid is not None and col_item is not None:
                        col_desc = self._find_col(headers, ["描述", "desc", "description", "物料描述"])
                        col_qty = self._find_col(headers, ["qty", "数量", "quantity", "new数量", "旧数量"])
                        col_type = self._find_col(headers, ["type", "类型", "物料类型"])
                        seen = set()
                        for d in rows:
                            duid = (d.get(headers[col_duid]) or "").strip()
                            if duid not in duid_set or duid in cached_duids:
                                continue
                            code = (d.get(headers[col_item]) or "").strip()
                            if not code or (duid, code) in seen:
                                continue
                            seen.add((duid, code))
                            materials.append({
                                "wid": duid_to_wid.get(duid),
                                "duid": duid,
                                "itemcode": code,
                                "desc": (d.get(headers[col_desc]) if col_desc is not None else "") or "",
                                "qty": (d.get(headers[col_qty]) if col_qty is not None else "") or "",
                                "type": (d.get(headers[col_type]) if col_type is not None else "") or "",
                                "status": "",
                                "remark": "",
                                "done": False,
                            })
            if not materials:
                return {"ok": True, "rows": []}

            # 合并已保存状态（w.materials 是最新的 done/status 来源；用替换前 itemcode 匹配，避免替换后匹配不到）
            saved_by_duid = {}
            for w in items:
                duid = (w.get("duid") or "").strip()
                saved_by_duid[duid] = {m.get("itemcode"): m for m in (w.get("materials") or []) if isinstance(m, dict) and m.get("itemcode")}
            for m in materials:
                sm = (saved_by_duid.get(m.get("duid") or "") or {}).get(m.get("itemcode")) or {}
                if sm:
                    m["status"] = sm.get("status", "") or m.get("status", "")
                    m["remark"] = sm.get("remark", "") or m.get("remark", "")
                    m["done"] = bool(sm.get("done"))

            # 应用物料替换（替换后 itemcode/qty 换成替换值，done/status/remark 随之继承）
            wid_repl = {}
            for w in items:
                repls = [r for r in (w.get("replacements") or []) if isinstance(r, dict) and r.get("from")]
                if repls:
                    wid_repl[w.get("id")] = {r.get("from"): r for r in repls}
            if wid_repl:
                for m in materials:
                    r = (wid_repl.get(m.get("wid")) or {}).get(m.get("itemcode"))
                    if r:
                        m["itemcode"] = (r.get("to") or m.get("itemcode")).strip() or m.get("itemcode")
                        if r.get("toQty"):
                            m["qty"] = r.get("toQty")

            # 筛选未完成物料，按 itemcode 聚合
            by_item = {}
            order = []
            for m in materials:
                if m.get("done"):
                    continue
                code = m.get("itemcode") or ""
                if not code:
                    continue
                if code not in by_item:
                    by_item[code] = {"itemcode": code, "desc": m.get("desc") or "", "type": m.get("type") or "", "duids": []}
                    order.append(code)
                by_item[code]["duids"].append({
                    "duid": m.get("duid") or "",
                    "workItemId": m.get("wid") or "",
                    "qty": m.get("qty") or "",
                    "status": m.get("status") or "",
                    "remark": m.get("remark") or "",
                })
            rows = [by_item[c] for c in order]
            for it in rows:
                total = 0.0
                for d in it["duids"]:
                    try:
                        total += float(d["qty"])
                    except (TypeError, ValueError):
                        pass
                it["totalQty"] = total
            return {"ok": True, "rows": rows}
        except Exception as e:
            return {"ok": False, "error": str(e)}

    def save_material_status(self, flow_id, work_item_id, items):
        """批量保存物料状态：items = [{itemcode, status?, remark?, done?}]，按 itemcode 合并更新。"""
        try:
            data = self.load_data()
            flow = next((f for f in data.get("flows", []) if f.get("id") == flow_id), None)
            if not flow:
                return {"ok": False, "error": "流程不存在"}
            w = next((x for x in flow.get("workItems", []) if x.get("id") == work_item_id), None)
            if not w:
                return {"ok": False, "error": "作业项不存在"}
            if not isinstance(w.get("materials"), list):
                w["materials"] = []
            saved = {m.get("itemcode"): m for m in w["materials"] if isinstance(m, dict) and m.get("itemcode")}
            for it in (items or []):
                if not isinstance(it, dict):
                    continue
                code = (it.get("itemcode") or "").strip()
                if not code:
                    continue
                if code not in saved:
                    saved[code] = {"itemcode": code}
                    w["materials"].append(saved[code])
                m = saved[code]
                if "status" in it:
                    m["status"] = (it.get("status") or "").strip()
                if "remark" in it:
                    m["remark"] = it.get("remark") or ""
                if "done" in it:
                    m["done"] = bool(it.get("done"))
            self.save_data(data)
            return {"ok": True}
        except Exception as e:
            return {"ok": False, "error": str(e)}

    def save_material_status_batch(self, flow_id, items):
        """批量保存物料状态（批次作业视图）：items = [{wid, itemcode, status?, remark?, done?}]，按 wid 分组更新对应 DU。"""
        try:
            data = self.load_data()
            flow = next((f for f in data.get("flows", []) if f.get("id") == flow_id), None)
            if not flow:
                return {"ok": False, "error": "流程不存在"}
            by_wid = {}
            for it in (items or []):
                if not isinstance(it, dict):
                    continue
                wid = it.get("wid")
                code = (it.get("itemcode") or "").strip()
                if not wid or not code:
                    continue
                by_wid.setdefault(wid, []).append(it)
            for wid, its in by_wid.items():
                w = next((x for x in flow.get("workItems", []) if x.get("id") == wid), None)
                if not w:
                    continue
                if not isinstance(w.get("materials"), list):
                    w["materials"] = []
                saved = {m.get("itemcode"): m for m in w["materials"] if isinstance(m, dict) and m.get("itemcode")}
                for it in its:
                    code = (it.get("itemcode") or "").strip()
                    if not code:
                        continue
                    if code not in saved:
                        saved[code] = {"itemcode": code}
                        w["materials"].append(saved[code])
                    m = saved[code]
                    if "status" in it:
                        m["status"] = (it.get("status") or "").strip()
                    if "remark" in it:
                        m["remark"] = it.get("remark") or ""
                    if "done" in it:
                        m["done"] = bool(it.get("done"))
            self.save_data(data)
            return {"ok": True}
        except Exception as e:
            return {"ok": False, "error": str(e)}

    def save_material_replacement(self, flow_id, work_item_id, from_code, to_code, to_qty):
        """记录物料替换：from_code 被替换成 to_code，数量 to_qty。
        原 itemcode 与数量写入备注（to_code 的 remark = "from_code * from_qty"）。"""
        try:
            data = self.load_data()
            flow = next((f for f in data.get("flows", []) if f.get("id") == flow_id), None)
            if not flow:
                return {"ok": False, "error": "流程不存在"}
            w = next((x for x in flow.get("workItems", []) if x.get("id") == work_item_id), None)
            if not w:
                return {"ok": False, "error": "作业项不存在"}
            from_code = (from_code or "").strip()
            to_code = (to_code or "").strip()
            to_qty = (to_qty or "").strip()
            if not from_code or not to_code:
                return {"ok": False, "error": "替换前后的 itemcode 都不能为空"}
            # 从锚点清单读原数量（from_qty）
            anchor_data = self.get_anchor_data(flow_id, work_item_id)
            from_qty = ""
            if anchor_data.get("ok"):
                for m in (anchor_data.get("materials") or []):
                    if m.get("itemcode") == from_code:
                        from_qty = str(m.get("qty") or "").strip()
                        break
            # 记录替换关系（同一 from 重复替换则更新）
            if not isinstance(w.get("replacements"), list):
                w["replacements"] = []
            repl = next((r for r in w["replacements"] if isinstance(r, dict) and r.get("from") == from_code), None)
            if repl:
                repl["to"] = to_code
                repl["toQty"] = to_qty
                repl["fromQty"] = from_qty
            else:
                w["replacements"].append({"from": from_code, "fromQty": from_qty, "to": to_code, "toQty": to_qty})
            # 备注写到替换后的 itemcode
            if not isinstance(w.get("materials"), list):
                w["materials"] = []
            saved = {m.get("itemcode"): m for m in w["materials"] if isinstance(m, dict) and m.get("itemcode")}
            if to_code not in saved:
                saved[to_code] = {"itemcode": to_code}
                w["materials"].append(saved[to_code])
            remark = from_code + (" * " + from_qty if from_qty else "")
            saved[to_code]["remark"] = remark
            self.save_data(data)
            return {"ok": True, "fromQty": from_qty, "remark": remark}
        except Exception as e:
            return {"ok": False, "error": str(e)}

    def save_material_status_options(self, flow_id, options):
        """保存物料状态枚举值（流程级，缺料/物料明细的状态下拉使用）。options = [str, ...]，去重后按序保存。"""
        try:
            data = self.load_data()
            flow = next((f for f in data.get("flows", []) if f.get("id") == flow_id), None)
            if not flow:
                return {"ok": False, "error": "流程不存在"}
            opts = []
            for o in (options or []):
                v = str(o).strip()
                if v and v not in opts:
                    opts.append(v)
            flow["materialStatusOptions"] = opts
            self.save_data(data)
            return {"ok": True, "options": opts}
        except Exception as e:
            return {"ok": False, "error": str(e)}

    def parse_duids(self, text):
        """解析粘贴的 DUID 文本（一行一个，可带表头/多余空白），返回去重后的 DUID 列表。"""
        try:
            text = text or ""
            duids = []
            seen = set()
            header_norm = {"duid", "du", "duid2", "duidlist", "id", "duid号", "du编号"}
            for raw in text.splitlines():
                line = raw.strip()
                if not line:
                    continue
                # 逗号/分号分隔 → 视为「一行多个 DUID」，全取
                if ',' in line or ';' in line:
                    tokens = [p.strip() for p in re.split(r'[,;]+', line) if p.strip()]
                else:
                    # tab/空白分隔 → 视为「多列」，取第一列 DUID
                    parts = [p for p in re.split(r'[\s\t]+', line) if p]
                    tokens = [parts[0]] if parts else []
                for token in tokens:
                    norm = re.sub(r'[\s_\-]', '', token).lower()
                    if norm in header_norm:
                        continue
                    if token not in seen:
                        seen.add(token)
                        duids.append(token)
            return {"ok": True, "duids": duids}
        except Exception as e:
            return {"ok": False, "error": str(e)}

    def parse_curl(self, curl_text):
        """解析「copy as cURL(bash)」文本，提取 url / method / headers（含 Cookie）。"""
        try:
            lines = [l.rstrip() for l in (curl_text or "").strip().splitlines()]
            merged = []
            for l in lines:
                if merged and merged[-1].endswith('\\'):
                    merged[-1] = merged[-1][:-1] + ' ' + l.strip()
                else:
                    merged.append(l.strip())
            text = ' '.join(merged).strip()
            if text.lower().startswith('curl'):
                text = text[4:].strip()
            toks = shlex.split(text, posix=True)
            url = None
            method = None
            headers = {}
            i, n = 0, len(toks)
            while i < n:
                t = toks[i]
                if t in ('-H', '--header') and i + 1 < n:
                    hv = toks[i + 1]
                    if ':' in hv:
                        k, v = hv.split(':', 1)
                        headers[k.strip()] = v.strip()
                    i += 2
                    continue
                if t in ('-X', '--request') and i + 1 < n:
                    method = toks[i + 1].upper()
                    i += 2
                    continue
                if t in ('-d', '--data', '--data-raw', '--data-binary') and i + 1 < n:
                    i += 2
                    continue
                if t.startswith('-'):
                    i += 1
                    continue
                if url is None and (t.startswith('http://') or t.startswith('https://')):
                    url = t
                i += 1
            if not url:
                for t in toks:
                    if t.startswith('http://') or t.startswith('https://'):
                        url = t
                        break
            if not url:
                return {"ok": False, "error": "未能从 curl 解析出 URL，请确认已复制完整内容"}
            method = method or "GET"
            cookie = next((v for k, v in headers.items() if k.lower() == 'cookie'), "")
            return {"ok": True, "url": url, "method": method, "headers": headers, "cookie": cookie}
        except ValueError as e:
            return {"ok": False, "error": "curl 解析失败（引号未配对？）：" + str(e)}
        except Exception as e:
            return {"ok": False, "error": "解析 curl 失败：" + str(e)}

    def _find_platform(self, platform_id):
        data = self.load_data()
        for p in data.get("platforms", []):
            if p.get("id") == platform_id:
                return p
        return None

    def make_platform_config_file(self, platform_id):
        """把某个平台配置（url/method/headers/cookie）写成 JSON 文件，供工具一个参数引用。"""
        try:
            p = self._find_platform(platform_id)
            if not p:
                return {"ok": False, "error": "平台配置不存在，请先保存"}
            os.makedirs(INPUTS_DIR, exist_ok=True)
            fn = "_platformcfg_" + self._safe_name(p.get("name") or "platform") + "_" + datetime.datetime.now().strftime("%Y%m%d_%H%M%S_%f") + ".json"
            path = os.path.join(INPUTS_DIR, fn)
            headers = {k: v for k, v in (p.get("headers") or {}).items() if k}
            payload = {
                "name": p.get("name"),
                "url": p.get("url"),
                "method": p.get("method") or "GET",
                "headers": headers,
                "cookie": headers.get("Cookie", ""),
                "updatedAt": p.get("updatedAt"),
            }
            with open(path, "w", encoding="utf-8") as f:
                json.dump(payload, f, ensure_ascii=False, indent=2)
            return {"ok": True, "path": path, "name": p.get("name")}
        except Exception as e:
            return {"ok": False, "error": str(e)}

    def _guess_download_name(self, url, resp_headers):
        """从 Content-Disposition 或 URL 推断下载文件名，推断不出返回 None。"""
        try:
            cd = (resp_headers or {}).get("Content-Disposition", "") if isinstance(resp_headers, dict) else ""
            if cd:
                m = re.search(r'filename\*?=(?:UTF-8\'\')?"?([^";]+)"?', cd, re.IGNORECASE)
                if m:
                    return os.path.basename(m.group(1).strip())
            from urllib.parse import urlparse, unquote
            name = os.path.basename(unquote(urlparse(url).path))
            if name and "." in name:
                return name
        except Exception:
            pass
        return None

    def download_checklist_file(self, flow_id, item_id, url, method, platform_id, save_name):
        """任务初始清单「从接口下载」：先清空该项固定目录旧文件，再下载最新保存。"""
        try:
            url = (url or "").strip()
            if not url:
                return {"ok": False, "error": "未配置下载 URL"}
            method = (method or "GET").upper()
            flow_id = str(flow_id or "")
            item_id = str(item_id or "")
            if not flow_id or not item_id:
                return {"ok": False, "error": "缺少流程或清单项标识"}

            item_dir = os.path.join(DATA_DIR, "checklist_files", flow_id, item_id)
            os.makedirs(item_dir, exist_ok=True)
            # 先删除旧文档，避免拿到旧文件
            for old in os.listdir(item_dir):
                try:
                    os.remove(os.path.join(item_dir, old))
                except Exception:
                    pass

            headers = {"User-Agent": "MRDN-System"}
            if platform_id:
                p = self._find_platform(platform_id)
                if p:
                    for k, v in (p.get("headers") or {}).items():
                        if k and v:
                            headers[k] = v

            try:
                import requests
                import urllib3
                urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)
            except Exception:
                return {"ok": False, "error": "缺少 requests 库"}

            kwargs = {"headers": headers, "timeout": 120, "allow_redirects": True, "verify": False}
            resp = requests.request(method, url, **kwargs)
            if resp.status_code >= 400:
                return {"ok": False, "error": "下载失败（HTTP " + str(resp.status_code) + "）"}
            content = resp.content
            if not content:
                return {"ok": False, "error": "接口返回为空"}

            safe = os.path.basename((save_name or "").strip())
            if not safe:
                safe = self._guess_download_name(url, resp.headers) or \
                       ("download_" + datetime.datetime.now().strftime("%Y%m%d_%H%M%S") + ".xlsx")
            path = os.path.join(item_dir, safe)
            with open(path, "wb") as f:
                f.write(content)
            rel = "checklist_files/" + flow_id + "/" + item_id + "/" + safe
            return {"ok": True, "name": safe, "path": rel, "size": len(content)}
        except Exception as e:
            return {"ok": False, "error": str(e)}

    def _first_platform_input_id(self, manifest):
        """取工具声明里第一个 type=platform 的输入 id（平台凭证参数名）。"""
        try:
            for inp in (manifest or {}).get("inputs", []):
                if inp.get("type") == "platform":
                    return inp.get("id")
        except Exception:
            pass
        return None

    def run_checklist_tool(self, flow_id, item_id, tool, platform_id, save_name, url=None):
        """任务初始清单「从工具获取」：运行工具 → 取第一个输出 → 复制到该项固定目录。"""
        try:
            tool = (tool or "").strip()
            if not tool:
                return {"ok": False, "error": "未选择工具"}
            flow_id = str(flow_id or "")
            item_id = str(item_id or "")
            if not flow_id or not item_id:
                return {"ok": False, "error": "缺少流程或清单项标识"}

            args = []
            if url and str(url).strip():
                args += ["--url", str(url).strip()]
            if platform_id:
                r = self.make_platform_config_file(platform_id)
                if not r or not r.get("ok"):
                    err = (r or {}).get("error") or "平台配置不存在"
                    return {"ok": False, "error": "平台凭证生成失败：" + err}
                pid = self._first_platform_input_id(self.get_tool_manifest(tool))
                if not pid:
                    return {"ok": False, "error": "工具未声明平台凭证输入，无法传递凭证"}
                args += ["--" + pid, r["path"]]

            res = self.run_tool(tool, args)
            if not res.get("ok"):
                return {"ok": False, "error": res.get("log") or "工具运行失败"}
            outputs = res.get("outputs") or []
            if not outputs:
                log = (res.get("log") or "").strip()
                err = "工具未产生输出文件"
                if log and log != "（无输出）":
                    tail = log[-500:].replace("\n", " ").replace("\r", " ")
                    err += "：" + tail
                return {"ok": False, "error": err}
            target = outputs[0]
            src_abs = None
            if target.get("rel"):
                src_abs = os.path.join(DATA_DIR, target["rel"])
            if not src_abs or not os.path.isfile(src_abs):
                src_abs = os.path.join(res.get("run_dir") or "", target.get("name") or "")
            if not os.path.isfile(src_abs):
                return {"ok": False, "error": "输出文件不存在：" + (target.get("name") or "")}

            item_dir = os.path.join(DATA_DIR, "checklist_files", flow_id, item_id)
            os.makedirs(item_dir, exist_ok=True)
            for old in os.listdir(item_dir):
                try:
                    os.remove(os.path.join(item_dir, old))
                except Exception:
                    pass
            safe = os.path.basename((save_name or "").strip()) or target.get("name") or \
                   ("tool_out_" + datetime.datetime.now().strftime("%Y%m%d_%H%M%S") + ".xlsx")
            dst = os.path.join(item_dir, safe)
            shutil.copy2(src_abs, dst)
            rel = "checklist_files/" + flow_id + "/" + item_id + "/" + safe
            return {"ok": True, "name": safe, "path": rel}
        except Exception as e:
            return {"ok": False, "error": str(e)}

    def get_tool_rule(self, name):
        """读取工具的规则说明 txt（tools/<stem>.txt），返回纯文本。
        供任务锚点「查询」操作展示所选工具版本（规则）的说明。"""
        try:
            name = (name or "").strip()
            if not name:
                return {"ok": False, "error": "未指定工具"}
            stem = os.path.splitext(name)[0]
            cands = [os.path.join(TOOLS_DIR, stem + ".txt")]
            if not os.path.isfile(cands[0]) and os.path.isdir(TOOLS_DIR):
                ns = self._norm_col(stem)
                for f in sorted(os.listdir(TOOLS_DIR)):
                    if f.lower().endswith(".txt") and self._norm_col(os.path.splitext(f)[0]) == ns:
                        cands[0] = os.path.join(TOOLS_DIR, f)
                        break
            for p in cands:
                if os.path.isfile(p):
                    with open(p, "rb") as fh:
                        raw = fh.read()
                    text = None
                    for enc in ("utf-8-sig", "gbk"):
                        try:
                            text = raw.decode(enc)
                            break
                        except Exception:
                            continue
                    if text is None:
                        text = raw.decode("utf-8", "replace")
                    return {"ok": True, "text": text}
            return {"ok": False, "error": "该工具没有规则说明 txt"}
        except Exception as e:
            return {"ok": False, "error": str(e)}

    def read_table(self, name):
        """读取表格文件完整内容，返回 {headers, rows}（rows 为 [{列名:值}, ...]，不截断）。
        供任务锚点「查询」操作读工具输出的查询结果表。"""
        try:
            name = (name or "").strip()
            if not name:
                return {"ok": False, "error": "空文件名"}
            target = self._abs_path(name)
            if not target:
                return {"ok": False, "error": "文件不存在：" + name}
            ext = os.path.splitext(target)[1].lower()
            headers, rows = self._read_table(target, ext)
            if headers is None:
                return {"ok": False, "error": "不是表格文件"}
            return {"ok": True, "headers": headers, "rows": rows}
        except Exception as e:
            return {"ok": False, "error": str(e)}

    def gen_anchor_cpq(self, flow_id, work_item_id, selections):
        """生成任务锚点 CPQ Spart 导入表。

        selections = [{itemcode, contractNumber, financialCode, qty, type}, ...]
        只生成主设备（type=Type 1），辅料跳过；所有主设备写进同一个 CPQ 模板。
        读固定 CPQ 模板（inputs/CPQ Spart导入模板.xlsx），填数据，保存到作业输出目录。"""
        try:
            data = self.load_data()
            flow = next((f for f in data.get("flows", []) if f.get("id") == flow_id), None)
            if not flow:
                return {"ok": False, "error": "流程不存在"}
            w = next((x for x in flow.get("workItems", []) if x.get("id") == work_item_id), None)
            if not w:
                return {"ok": False, "error": "作业项不存在"}
            duid = (w.get("duid") or "").strip()
            if not duid:
                return {"ok": False, "error": "当前作业无 DUID"}

            def _is_type1(t):
                if not t:
                    return False
                n = self._norm_col(t)
                return n in ("type1", "1", "类型1")

            main_rows = [s for s in (selections or [])
                         if isinstance(s, dict) and _is_type1(s.get("type"))]
            if not main_rows:
                return {"ok": False, "error": "没有可生成的主设备（type=Type 1）物料，请确认勾选的是主设备"}

            tpl = os.path.join(INPUTS_DIR, "CPQ Spart导入模板.xlsx")
            if not os.path.isfile(tpl) and os.path.isdir(INPUTS_DIR):
                alt = None
                for f in sorted(os.listdir(INPUTS_DIR)):
                    if "CPQ" in f.upper() and f.lower().endswith(".xlsx"):
                        alt = os.path.join(INPUTS_DIR, f)
                        break
                tpl = alt
            if not tpl or not os.path.isfile(tpl):
                return {"ok": False, "error": "找不到 CPQ Spart 导入模板（inputs/CPQ Spart导入模板.xlsx）"}

            import openpyxl
            wb = openpyxl.load_workbook(tpl)
            sheet_target = "Spart Configuration"
            if sheet_target not in wb.sheetnames:
                return {"ok": False, "error": "CPQ 模板缺少 sheet：Spart Configuration"}
            ws = wb[sheet_target]
            headers = [str(c.value).strip() if c.value is not None else "" for c in ws[1]]
            col_contract = self._find_col(headers, ["HUAWEI CONTRACT NO."])
            col_product = self._find_col(headers, ["Product Type"])
            col_puno = self._find_col(headers, ["PU NO."])
            col_duid = self._find_col(headers, ["DU ID"])
            col_fin = self._find_col(headers, ["FINANCIAL PRODUCT CODE"])
            col_spart = self._find_col(headers, ["SPART NUMBER"])
            col_qty = self._find_col(headers, ["SPART QTY."])
            if col_spart is None or col_qty is None:
                return {"ok": False, "error": "CPQ 模板缺少 SPART NUMBER / SPART QTY. 列"}

            if ws.max_row >= 2:
                ws.delete_rows(2, ws.max_row - 1)

            for i, s in enumerate(main_rows):
                r = 2 + i
                if col_contract is not None:
                    ws.cell(row=r, column=col_contract + 1, value=s.get("contractNumber") or "")
                if col_product is not None:
                    ws.cell(row=r, column=col_product + 1, value="Equipment")
                if col_puno is not None:
                    c = ws.cell(row=r, column=col_puno + 1, value="001")
                    c.number_format = "@"
                if col_duid is not None:
                    ws.cell(row=r, column=col_duid + 1, value=duid)
                if col_fin is not None:
                    ws.cell(row=r, column=col_fin + 1, value=s.get("financialCode") or "")
                if col_spart is not None:
                    ws.cell(row=r, column=col_spart + 1, value=s.get("itemcode") or "")
                if col_qty is not None:
                    ws.cell(row=r, column=col_qty + 1, value=s.get("qty") or "")

            run_dir = os.path.join(OUTPUTS_DIR, "任务锚点CPQ_" + datetime.datetime.now().strftime("%Y%m%d_%H%M%S_%f"))
            os.makedirs(run_dir, exist_ok=True)
            out_name = "CPQ导入表.xlsx"
            out_path = os.path.join(run_dir, out_name)
            wb.save(out_path)
            rel = "outputs/" + os.path.basename(run_dir) + "/" + out_name
            return {"ok": True, "path": out_path, "rel": rel, "count": len(main_rows), "duid": duid}
        except Exception as e:
            return {"ok": False, "error": str(e)}

    def gen_shortage_cpq(self, flow_id, selections):
        """缺料清单生成 CPQ Spart 导入表（跨 DU）。

        selections = [{itemcode, contractNumber, financialCode, qty, type, duid}, ...]
        每个 selection 自带目标 duid，DU ID 列按行填写各自的 duid。
        只生成主设备（type=Type 1），辅料跳过。"""
        try:
            data = self.load_data()
            flow = next((f for f in data.get("flows", []) if f.get("id") == flow_id), None)
            if not flow:
                return {"ok": False, "error": "流程不存在"}

            def _is_type1(t):
                if not t:
                    return False
                n = self._norm_col(t)
                return n in ("type1", "1", "类型1")

            main_rows = [s for s in (selections or [])
                         if isinstance(s, dict) and _is_type1(s.get("type"))]
            if not main_rows:
                return {"ok": False, "error": "没有可生成的主设备（type=Type 1）物料，请确认勾选的是主设备"}

            tpl = os.path.join(INPUTS_DIR, "CPQ Spart导入模板.xlsx")
            if not os.path.isfile(tpl) and os.path.isdir(INPUTS_DIR):
                alt = None
                for f in sorted(os.listdir(INPUTS_DIR)):
                    if "CPQ" in f.upper() and f.lower().endswith(".xlsx"):
                        alt = os.path.join(INPUTS_DIR, f)
                        break
                tpl = alt
            if not tpl or not os.path.isfile(tpl):
                return {"ok": False, "error": "找不到 CPQ Spart 导入模板（inputs/CPQ Spart导入模板.xlsx）"}

            import openpyxl
            wb = openpyxl.load_workbook(tpl)
            sheet_target = "Spart Configuration"
            if sheet_target not in wb.sheetnames:
                return {"ok": False, "error": "CPQ 模板缺少 sheet：Spart Configuration"}
            ws = wb[sheet_target]
            headers = [str(c.value).strip() if c.value is not None else "" for c in ws[1]]
            col_contract = self._find_col(headers, ["HUAWEI CONTRACT NO."])
            col_product = self._find_col(headers, ["Product Type"])
            col_puno = self._find_col(headers, ["PU NO."])
            col_duid = self._find_col(headers, ["DU ID"])
            col_fin = self._find_col(headers, ["FINANCIAL PRODUCT CODE"])
            col_spart = self._find_col(headers, ["SPART NUMBER"])
            col_qty = self._find_col(headers, ["SPART QTY."])
            if col_spart is None or col_qty is None:
                return {"ok": False, "error": "CPQ 模板缺少 SPART NUMBER / SPART QTY. 列"}

            if ws.max_row >= 2:
                ws.delete_rows(2, ws.max_row - 1)

            for i, s in enumerate(main_rows):
                r = 2 + i
                if col_contract is not None:
                    ws.cell(row=r, column=col_contract + 1, value=s.get("contractNumber") or "")
                if col_product is not None:
                    ws.cell(row=r, column=col_product + 1, value="Equipment")
                if col_puno is not None:
                    c = ws.cell(row=r, column=col_puno + 1, value="001")
                    c.number_format = "@"
                if col_duid is not None:
                    ws.cell(row=r, column=col_duid + 1, value=s.get("duid") or "")
                if col_fin is not None:
                    ws.cell(row=r, column=col_fin + 1, value=s.get("financialCode") or "")
                if col_spart is not None:
                    ws.cell(row=r, column=col_spart + 1, value=s.get("itemcode") or "")
                if col_qty is not None:
                    ws.cell(row=r, column=col_qty + 1, value=s.get("qty") or "")

            run_dir = os.path.join(OUTPUTS_DIR, "缺料清单CPQ_" + datetime.datetime.now().strftime("%Y%m%d_%H%M%S_%f"))
            os.makedirs(run_dir, exist_ok=True)
            out_name = "CPQ导入表.xlsx"
            out_path = os.path.join(run_dir, out_name)
            wb.save(out_path)
            rel = "outputs/" + os.path.basename(run_dir) + "/" + out_name

            # 生成 CPQ 后给涉及 DU 的物料加备注（不自动完成物料状态）
            try:
                remark_set = set()
                for s in main_rows:
                    duid = (s.get("duid") or "").strip()
                    code = (s.get("itemcode") or "").strip()
                    if not duid or not code:
                        continue
                    w2 = next((x for x in flow.get("workItems", []) if (x.get("duid") or "").strip() == duid), None)
                    if not w2:
                        continue
                    # selection.itemcode 是替换后，反向映射回原始 itemcode
                    orig = code
                    for r in (w2.get("replacements") or []):
                        if isinstance(r, dict) and (r.get("to") or "").strip() == code:
                            orig = (r.get("from") or "").strip() or code
                            break
                    key = (w2.get("id"), orig)
                    if key in remark_set:
                        continue
                    remark_set.add(key)
                    if not isinstance(w2.get("materials"), list):
                        w2["materials"] = []
                    m0 = next((x for x in w2["materials"] if isinstance(x, dict) and (x.get("itemcode") or "").strip() == orig), None)
                    if m0 is None:
                        m0 = {"itemcode": orig}
                        w2["materials"].append(m0)
                    m0["remark"] = "已缺料处理生成文件"
                if remark_set:
                    self.save_data(data)
            except Exception:
                pass  # 备注失败不阻断生成

            return {"ok": True, "path": out_path, "rel": rel, "count": len(main_rows)}
        except Exception as e:
            return {"ok": False, "error": str(e)}

    def test_platform_connection(self, platform_id):
        """用保存的 url + headers（含 cookie）发起一次请求，返回连接成功/失败状态（MySQL 测试式）。"""
        try:
            p = self._find_platform(platform_id)
            if not p:
                return {"ok": False, "connected": False, "message": "平台配置不存在"}
            url = (p.get("url") or "").strip()
            if not url:
                return {"ok": False, "connected": False, "message": "未配置 URL"}
            headers = {k: v for k, v in (p.get("headers") or {}).items() if k}
            method = (p.get("method") or "GET").upper()
            body = p.get("body") or ""
            return self._do_test_connection(url, method, headers, body)
        except Exception as e:
            return {"ok": False, "connected": False, "status_code": None, "message": "连接失败：" + str(e)}

    def test_platform_connection_raw(self, url, method, headers, cookie, body):
        """用原始参数直接测试连接（编辑期未保存也可测试）。"""
        try:
            url = (url or "").strip()
            if not url:
                return {"ok": False, "connected": False, "message": "未配置 URL"}
            h = {k: v for k, v in (headers or {}).items() if k and v}
            if cookie:
                h["Cookie"] = cookie
            return self._do_test_connection(url, (method or "GET").upper(), h, body or "")
        except Exception as e:
            return {"ok": False, "connected": False, "status_code": None, "message": "连接失败：" + str(e)}

    def _do_test_connection(self, url, method, headers, body):
        try:
            import requests
            import urllib3
            urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)
        except Exception:
            pass
        kwargs = {"headers": headers, "timeout": 20, "allow_redirects": True, "verify": False}
        if method in ("POST", "PUT", "PATCH") and body:
            kwargs["data"] = body
        resp = requests.request(method, url, **kwargs)
        status = resp.status_code
        if 200 <= status < 400:
            return {"ok": True, "connected": True, "status_code": status, "message": "连接成功（HTTP " + str(status) + "）", "url": url}
        if status in (401, 403):
            return {"ok": True, "connected": False, "status_code": status, "message": "已连通但登录态失效（HTTP " + str(status) + "，Cookie 可能已过期）", "url": url}
        return {"ok": True, "connected": False, "status_code": status, "message": "连接失败（HTTP " + str(status) + "）", "url": url}

    def _patch_playwright_driver(self):
        """PyInstaller 冻结后，playwright 驱动(node.exe + package/*.js) 位于 _MEIPASS 下，
        但 playwright 默认按 inspect.getfile(playwright).parent 找 driver（指到 PYZ 内不存在路径）。
        仅在冻结态重定向 compute_driver_executable 到 _MEIPASS/playwright/driver；
        开发态保持 playwright 默认（site-packages 下）。"""
        try:
            meipass = getattr(sys, '_MEIPASS', None)
            if not meipass:
                return
            import playwright._impl._transport as _pt
            import playwright._impl._driver as _pd
            base = os.path.join(meipass, 'playwright', 'driver')

            def _compute():
                node = os.environ.get('PLAYWRIGHT_NODEJS_PATH') or os.path.join(base, 'node.exe')
                cli = os.path.join(base, 'package', 'cli.js')
                return node, cli
            _pt.compute_driver_executable = _compute
            _pd.compute_driver_executable = _compute
        except Exception:
            pass

    def start_platform_login(self, url):
        """启动一个可见的 Edge 窗口打开 url 供用户登录，返回 session_id。
        用户登录完成后前端调用 finish_platform_login 抓取 cookie（含 HttpOnly）。"""
        url = (url or "").strip()
        if not url:
            return {"ok": False, "error": "请先填写要登录的网址"}
        sid = "login_" + uuid.uuid4().hex[:10]
        evt = threading.Event()
        state = {"result": None, "cancelled": False}
        self._patch_playwright_driver()

        def worker():
            result = {"ok": False, "error": "登录流程未完成"}
            try:
                result = self._login_sync(url, state, evt)
            except Exception as e:
                result = {"ok": False, "error": "启动浏览器失败：" + str(e)}
            state["result"] = result

        t = threading.Thread(target=worker, daemon=True)
        self._login_sessions[sid] = {"thread": t, "event": evt, "state": state}
        t.start()
        return {"ok": True, "session_id": sid}

    def _login_sync(self, url, state, evt):
        from playwright.sync_api import sync_playwright
        with sync_playwright() as pw:
            browser = pw.chromium.launch(channel="msedge", headless=False)
            context = browser.new_context()
            page = context.new_page()
            captured = {}

            def on_request(req):
                h = req.headers
                for k in ("authorization", "x-auth-token", "x-csrf-token", "x-xsrf-token", "x-api-key"):
                    if h.get(k) and k not in captured:
                        captured[k] = h[k]
                if h.get("user-agent") and "user-agent" not in captured:
                    captured["user-agent"] = h["user-agent"]
            page.on("request", on_request)

            goto_error = None
            try:
                page.goto(url, wait_until="domcontentloaded", timeout=60000)
            except Exception as e:
                goto_error = str(e)

            def cookie_sig():
                try:
                    cs = context.cookies()
                    return tuple(sorted((c.get("name") or "", c.get("value") or "") for c in cs if c.get("name")))
                except Exception:
                    return ()

            # 等页面初始加载稳定（3 秒），让登录页初始 cookie 完整建立，
            # 避免把页面加载自身产生的 cookie 变化误判成「登录完成」而提前退出
            evt.wait(3.0)
            baseline = cookie_sig()
            prev_sig = baseline
            stable_since = None
            auto_detected = False
            deadline = time.time() + 180

            while not evt.is_set():
                if state.get("cancelled"):
                    break
                sig = cookie_sig()
                # 只以 cookie 变化作为登录依据（去掉 URL 变化：登录页重定向会误判，导致用户还没输完账号密码就退出）
                changed = bool(sig) and sig != baseline
                if changed:
                    if sig == prev_sig:
                        if stable_since is None:
                            stable_since = time.time()
                        elif time.time() - stable_since >= 1.2:
                            auto_detected = True
                            break
                    else:
                        stable_since = None
                else:
                    stable_since = None
                prev_sig = sig
                if time.time() > deadline:
                    break
                # 未检测到 cookie 变化（用户还在输入账号密码）时，等待 20 秒再检测，给用户充足输入时间；
                # 已检测到变化则用短间隔快速确认稳定
                evt.wait(20.0 if not changed else 0.4)

            if state.get("cancelled"):
                state["result"] = {"ok": False, "cancelled": True}
            else:
                try:
                    cookies = context.cookies()
                    cookie_str = "; ".join(f"{c.get('name')}={c.get('value')}" for c in cookies if c.get('name'))
                    try:
                        ua = page.evaluate("() => navigator.userAgent")
                    except Exception:
                        ua = captured.get("user-agent", "")
                    headers = {}
                    if cookie_str:
                        headers["Cookie"] = cookie_str
                    if ua:
                        headers["User-Agent"] = ua
                    for k, v in captured.items():
                        if k != "user-agent":
                            headers[k] = v
                    state["result"] = {
                        "ok": True, "url": page.url, "headers": headers, "cookie": cookie_str,
                        "goto_error": goto_error, "auto_detected": auto_detected,
                    }
                except Exception as e:
                    state["result"] = {"ok": False, "error": "读取 cookie 失败：" + str(e)}
            try:
                context.close()
                browser.close()
            except Exception:
                pass
            return state.get("result")

    def finish_platform_login(self, session_id):
        sess = self._login_sessions.get(session_id)
        if not sess:
            return {"ok": False, "error": "登录会话不存在或已结束"}
        sess["event"].set()
        sess["thread"].join(timeout=40)
        self._login_sessions.pop(session_id, None)
        return sess["state"].get("result") or {"ok": False, "error": "登录流程超时或未完成"}

    def check_platform_login(self, session_id):
        sess = self._login_sessions.get(session_id)
        if not sess:
            return {"done": True, "result": {"ok": False, "error": "登录会话不存在或已结束"}}
        if sess["thread"].is_alive():
            return {"done": False}
        res = sess["state"].get("result")
        self._login_sessions.pop(session_id, None)
        return {"done": True, "result": res or {"ok": False, "error": "登录流程未完成"}}

    def cancel_platform_login(self, session_id):
        sess = self._login_sessions.get(session_id)
        if not sess:
            return {"ok": False, "error": "登录会话不存在"}
        sess["state"]["cancelled"] = True
        sess["event"].set()
        sess["thread"].join(timeout=15)
        self._login_sessions.pop(session_id, None)
        return {"ok": True}

    def create_work_folder(self, duid):
        """为某个 DUID 建立作业文件夹，返回相对 data 的路径（works/<duid>/）。"""
        try:
            safe = self._safe_name(duid)
            if not safe or safe == "未命名":
                return {"ok": False, "error": "DUID 为空或非法"}
            folder = os.path.join(WORKS_DIR, safe)
            os.makedirs(folder, exist_ok=True)
            return {"ok": True, "folder": "works/" + safe, "abs": folder}
        except Exception as e:
            return {"ok": False, "error": str(e)}

    def remove_work_folder(self, duid):
        """删除 DUID 作业项时，把文件夹移到回收站 works/_trash/（可恢复）。"""
        try:
            safe = self._safe_name(duid)
            folder = os.path.join(WORKS_DIR, safe)
            if not os.path.isdir(folder):
                return {"ok": True}
            trash = os.path.join(WORKS_DIR, "_trash")
            os.makedirs(trash, exist_ok=True)
            dst = os.path.join(trash, safe + "_" + datetime.datetime.now().strftime("%Y%m%d_%H%M%S_%f"))
            shutil.move(folder, dst)
            return {"ok": True}
        except Exception as e:
            return {"ok": False, "error": str(e)}

    def _trash_path(self, path):
        """把 path（文件或目录）移到 works/_trash/（可恢复），不存在则忽略。返回是否移动。"""
        try:
            if not path or not os.path.exists(path):
                return False
            trash = os.path.join(WORKS_DIR, "_trash")
            os.makedirs(trash, exist_ok=True)
            base = os.path.basename(os.path.normpath(path))
            dst = os.path.join(trash, base + "_" + datetime.datetime.now().strftime("%Y%m%d_%H%M%S_%f"))
            shutil.move(path, dst)
            return True
        except Exception:
            return False

    def clear_step_archive(self, steps, batch_id=None, duid=None):
        """重跑节点时，删除选中节点对应的归档目录（只清选中节点，其他保留）。
        steps: [{seq, name}]；batch_id 给批次时清批次级归档 + 批次内各 DU 归档；
        duid 给单 DU 时只清该 DU 的节点归档。"""
        try:
            seq_names = []
            for s in (steps or []):
                try:
                    seq_str = str(int(s.get("seq") or 0)).zfill(2)
                except Exception:
                    seq_str = "00"
                sname = self._safe_name(s.get("name")) or "节点"
                seq_names.append((seq_str, sname))
            if not seq_names:
                return {"ok": True, "cleared": 0}
            cleared = []
            if batch_id:
                safe = self._safe_name(self._batch_name(batch_id)) or "批次"
                base = os.path.join(WORKS_DIR, "_batch", safe)
                for seq_str, sname in seq_names:
                    if self._trash_path(os.path.join(base, seq_str + "_" + sname)):
                        cleared.append(seq_str + "_" + sname)
                data = self.load_data()
                for fl in data.get("flows", []):
                    for w in fl.get("workItems", []):
                        if w.get("batchId") == batch_id and w.get("duid"):
                            dfolder = os.path.join(WORKS_DIR, self._safe_name(w.get("duid")))
                            for seq_str, sname in seq_names:
                                if self._trash_path(os.path.join(dfolder, seq_str + "_" + sname)):
                                    cleared.append(seq_str + "_" + sname)
            elif duid:
                dfolder = os.path.join(WORKS_DIR, self._safe_name(duid))
                for seq_str, sname in seq_names:
                    if self._trash_path(os.path.join(dfolder, seq_str + "_" + sname)):
                        cleared.append(seq_str + "_" + sname)
            return {"ok": True, "cleared": len(cleared)}
        except Exception as e:
            return {"ok": False, "error": str(e)}

    def archive_work_files(self, duid, seq, step_name, files):
        """把 files（[{name,rel}] 或 rel 字符串列表）复制归档到 DUID 作业文件夹的节点子目录。
        返回归档后的文件列表（name/rel）。"""
        try:
            f = self.create_work_folder(duid)
            if not f.get("ok"):
                return f
            seq_str = str(int(seq or 0)).zfill(2)
            sname = self._safe_name(step_name) or "节点"
            step_dir = os.path.join(f["abs"], seq_str + "_" + sname)
            os.makedirs(step_dir, exist_ok=True)
            base_rel = f["folder"] + "/" + seq_str + "_" + sname
            archived = []
            for it in (files or []):
                if isinstance(it, str):
                    rel = it
                    name = os.path.basename(it)
                else:
                    rel = (it.get("rel") or it.get("path") or "").strip()
                    name = (it.get("name") or "").strip() or os.path.basename(rel)
                if not rel:
                    continue
                src = self._abs_path(rel)
                if not src:
                    continue
                dst = os.path.join(step_dir, name)
                try:
                    if os.path.abspath(src) != os.path.abspath(dst):
                        shutil.copy2(src, dst)
                    archived.append({"name": name, "rel": base_rel + "/" + name})
                except Exception:
                    continue
            return {"ok": True, "files": archived}
        except Exception as e:
            return {"ok": False, "error": str(e)}

    def save_evidence(self, duids, seq, step_name, evidence):
        """把操作举证保存到各 DUID 的节点归档目录。

        evidence 支持两种：
          - 字符串 data_url（图片，向后兼容）→ 存「图片举证.<ext>」
          - dict {kind, name, dataUrl}（图片或文件）
            · kind='image' → 存「图片举证.<ext>」
            · kind='file'  → 存「举证_<原文件名>」
        duids: 单DU时 [duid]；批次时批次内所有 DUID。返回各 DU 保存后的相对路径。"""
        try:
            if isinstance(evidence, dict):
                kind = evidence.get("kind", "")
                name = evidence.get("name") or ""
                data_url = evidence.get("dataUrl") or ""
                is_image = (kind == "image")
            else:
                data_url = evidence or ""
                name = ""
                is_image = True

            if not data_url or "," not in str(data_url):
                return {"ok": False, "error": "空举证"}

            header, b64 = str(data_url).split(",", 1)
            raw = base64.b64decode(b64)
            mime = (header.split(":")[1].split(";")[0]) if ":" in header else ""

            if is_image or not name:
                ext = {"image/png": "png", "image/jpeg": "jpg", "image/gif": "gif",
                       "image/webp": "webp", "image/bmp": "bmp"}.get(mime, "jpg")
                fn = "图片举证." + ext
            else:
                fn = "举证_" + self._safe_name(name)

            seq_str = str(int(seq or 0)).zfill(2)
            sname = self._safe_name(step_name) or "节点"
            saved = []
            for duid in (duids or []):
                f = self.create_work_folder(duid)
                if not f.get("ok"):
                    continue
                step_dir = os.path.join(f["abs"], seq_str + "_" + sname)
                os.makedirs(step_dir, exist_ok=True)
                with open(os.path.join(step_dir, fn), "wb") as fh:
                    fh.write(raw)
                saved.append({"duid": duid, "rel": f["folder"] + "/" + seq_str + "_" + sname + "/" + fn})
            return {"ok": True, "saved": saved}
        except Exception as e:
            return {"ok": False, "error": str(e)}

    def clear_du_evidence(self, duid):
        """重置单 DUID 进度时，删除该 DU 归档目录下所有举证文件（可恢复）：
        「图片举证.*」截图 + 「举证_*」附件。只清举证，不动其他归档文件。"""
        try:
            safe = self._safe_name(duid)
            folder = os.path.join(WORKS_DIR, safe)
            if not os.path.isdir(folder):
                return {"ok": True, "cleared": 0}
            cleared = 0
            for root, _dirs, files in os.walk(folder):
                for fn in files:
                    if fn.startswith("图片举证") or fn.startswith("举证_"):
                        if self._trash_path(os.path.join(root, fn)):
                            cleared += 1
            return {"ok": True, "cleared": cleared}
        except Exception as e:
            return {"ok": False, "error": str(e)}

    def _file_rel_name(self, it):
        """从 files 元素（字符串 rel 或 {name,rel}）提取 (rel, name)。"""
        if isinstance(it, str):
            return it, os.path.basename(it)
        rel = (it.get("rel") or it.get("path") or "").strip()
        name = (it.get("name") or "").strip() or os.path.basename(rel)
        return rel, name

    def archive_batch_files(self, batch_id, seq, step_name, files):
        """把批次「合表结果」整体归档到 works/_batch/<批次名>/<seq>_<节点>/。
        只保留批次运行结果（含批次全体的 Excel / 截图 / 输入文件）。"""
        try:
            safe = self._safe_name(self._batch_name(batch_id)) or "批次"
            folder = os.path.join(WORKS_DIR, "_batch", safe)
            os.makedirs(folder, exist_ok=True)
            seq_str = str(int(seq or 0)).zfill(2)
            sname = self._safe_name(step_name) or "节点"
            step_dir = os.path.join(folder, seq_str + "_" + sname)
            os.makedirs(step_dir, exist_ok=True)
            base_rel = "works/_batch/" + safe + "/" + seq_str + "_" + sname
            archived = []
            for it in (files or []):
                rel, name = self._file_rel_name(it)
                if not rel:
                    continue
                src = self._abs_path(rel)
                if not src:
                    continue
                dst = os.path.join(step_dir, name)
                try:
                    if os.path.abspath(src) != os.path.abspath(dst):
                        shutil.copy2(src, dst)
                    archived.append({"name": name, "rel": base_rel + "/" + name})
                except Exception:
                    continue
            return {"ok": True, "files": archived}
        except Exception as e:
            return {"ok": False, "error": str(e)}

    def _split_excel_for_du(self, src_path, duid):
        """把一个「合表」Excel 按 DUID 列拆出指定 DU 的行，返回新的 openpyxl 工作簿；
        找不到 DUID 列或该 DU 无数据时返回 None。"""
        import openpyxl
        duid = (duid or "").strip()
        if not duid:
            return None
        wb = openpyxl.load_workbook(src_path)
        ws = wb.active
        rows = list(ws.iter_rows())
        if len(rows) < 2:
            return None
        duid_headers = {"DUID", "DU ID", "DU_ID", "DU-ID"}
        du_col = None
        hdr_idx = None
        for idx in range(min(8, len(rows))):
            for ci, cell in enumerate(rows[idx]):
                v = str(cell.value or "").strip().upper()
                if v in duid_headers:
                    du_col = ci
                    hdr_idx = idx
                    break
            if du_col is not None:
                break
        if du_col is None:
            return None

        def _parse(v):
            if v is None:
                return []
            return [d.strip() for d in re.split(r"[、,，\s;；]+", str(v)) if d.strip()]

        # 从后往前删除表头之下不含该 DU 的数据行，保留表头及格式
        matched = 0
        for idx in range(len(rows) - 1, hdr_idx, -1):
            row = rows[idx]
            val = row[du_col].value if du_col < len(row) else None
            if duid in _parse(val):
                matched += 1
            else:
                try:
                    ws.delete_rows(idx + 1, 1)
                except Exception:
                    pass
        if matched == 0:
            return None
        return wb

    def archive_batch_du_files(self, batch_id, duids, seq, step_name, files):
        """批次完成后，把「合表」中的 Excel 按 DUID 拆出，分别归档到各 DU 的
        works/<duid>/<seq>_<节点>/（只含该 DU 数据行）；截图/整批输出件不进 DU 归档。"""
        result = {"ok": True, "per_du": {}}
        for duid in (duids or []):
            f = self.create_work_folder(duid)
            if not f.get("ok"):
                result["per_du"][duid] = []
                continue
            seq_str = str(int(seq or 0)).zfill(2)
            sname = self._safe_name(step_name) or "节点"
            step_dir = os.path.join(f["abs"], seq_str + "_" + sname)
            os.makedirs(step_dir, exist_ok=True)
            base_rel = f["folder"] + "/" + seq_str + "_" + sname
            archived = []
            for it in (files or []):
                rel, name = self._file_rel_name(it)
                if not rel:
                    continue
                ext = os.path.splitext(name)[1].lower()
                if ext not in (".xlsx", ".xlsm"):
                    continue
                src = self._abs_path(rel)
                if not src:
                    continue
                wb = self._split_excel_for_du(src, duid)
                if wb is None:
                    continue
                dst = os.path.join(step_dir, name)
                try:
                    wb.save(dst)
                    archived.append({"name": name, "rel": base_rel + "/" + name})
                except Exception:
                    continue
            result["per_du"][duid] = archived
        return result

    def open_archive_folder(self, type, id):
        """打开 DU 或批次的归档文件夹。type: 'du' | 'batch'。"""
        data = self.load_data()
        try:
            if type == "du":
                w = next((x for fl in data.get("flows", []) for x in fl.get("workItems", [])
                          if x.get("id") == id), None)
                if not w:
                    return {"ok": False, "error": "作业项不存在"}
                folder = os.path.join(WORKS_DIR, self._safe_name(w.get("duid") or ""))
            elif type == "batch":
                folder = os.path.join(WORKS_DIR, "_batch", self._safe_name(self._batch_name(id)))
            else:
                return {"ok": False, "error": "type 参数错误"}
            os.makedirs(folder, exist_ok=True)
            os.startfile(folder)
            return {"ok": True}
        except Exception as e:
            return {"ok": False, "error": str(e)}

    def export_work_record(self, flow_id, work_item_id):
        """导出某个 DUID 的作业记录文档（Excel），返回生成的文件绝对路径。"""
        try:
            import openpyxl
            from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
            from openpyxl.utils import get_column_letter

            data = self.load_data()
            flow = next((f for f in data.get("flows", []) if f.get("id") == flow_id), None)
            if not flow:
                return {"ok": False, "error": "流程不存在"}
            w = next((x for x in flow.get("workItems", []) if x.get("id") == work_item_id), None)
            if not w:
                return {"ok": False, "error": "作业项不存在"}

            steps = flow.get("steps", [])
            done_ids = set(w.get("doneStepIds") or [])
            records = w.get("records") or []
            rec_by_step = {r.get("stepId"): r for r in records if r.get("stepId")}
            # 当前节点 = 第一个未完成节点
            current_id = next((s.get("id") for s in steps if s.get("id") not in done_ids), None)

            HDR_FILL = PatternFill("solid", fgColor="5B6CFF")
            HDR_FONT = Font(bold=True, color="FFFFFF")
            HDR_ALIGN = Alignment(horizontal="center", vertical="center", wrap_text=True)
            WRAP = Alignment(vertical="top", wrap_text=True)
            THIN = Side(style="thin", color="D9DEE8")
            BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)

            wb = openpyxl.Workbook()
            ws = wb.active
            ws.title = "作业记录"

            # ---- 头部信息区 ----
            detached = bool(not w.get("batchId") and w.get("sourceBatchId"))
            status = "已完成" if (steps and len(done_ids) >= len(steps)) else ("进行中" if current_id else "未开始")
            head = [
                ("DUID", w.get("duid") or ""),
                ("所属批次", w.get("batchId") or ("（拆出，来源 " + str(w.get("sourceBatchId")) + "）" if detached else "（无）")),
                ("状态", status),
                ("进度", "%d / %d 节点" % (len(done_ids), len(steps))),
                ("创建时间", w.get("createdAt") or ""),
                ("最后更新", w.get("updatedAt") or ""),
                ("备注", w.get("remark") or ""),
            ]
            ws.merge_cells("A1:H1")
            c = ws["A1"]
            c.value = "MRDN 系统 · 作业记录"
            c.font = Font(bold=True, size=16, color="5B6CFF")
            c.alignment = Alignment(horizontal="left", vertical="center")
            ws.row_dimensions[1].height = 28

            r = 3
            for k, v in head:
                ws.cell(row=r, column=1, value=k).font = Font(bold=True, color="5B6CFF")
                ws.cell(row=r, column=1).alignment = Alignment(vertical="center")
                ws.cell(row=r, column=2, value=v).alignment = WRAP
                ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=8)
                r += 1
            r += 1

            # ---- 节点明细表 ----
            ws.cell(row=r, column=1, value="节点作业明细").font = Font(bold=True, size=12)
            r += 1
            headers = ["序号", "节点名称", "状态", "完成时间", "归档文件"]
            for i, h in enumerate(headers, 1):
                cell = ws.cell(row=r, column=i, value=h)
                cell.font = HDR_FONT
                cell.fill = HDR_FILL
                cell.alignment = HDR_ALIGN
                cell.border = BORDER
            r += 1
            for s in steps:
                sid = s.get("id")
                rec = rec_by_step.get(sid)
                done = sid in done_ids
                if done:
                    st = "已完成"
                elif sid == current_id:
                    st = "进行中"
                else:
                    st = "未完成"
                done_at = (rec.get("doneAt") or "") if rec else ""
                files = []
                if rec:
                    files = [f.get("name") or "" for f in (rec.get("files") or []) if f.get("name")]
                file_txt = "\n".join(files)
                row_vals = [s.get("seq"), s.get("name") or sid, st, done_at, file_txt]
                for i, v in enumerate(row_vals, 1):
                    cell = ws.cell(row=r, column=i, value=v)
                    cell.alignment = WRAP
                    cell.border = BORDER
                r += 1

            # 列宽
            widths = [8, 26, 10, 20, 40]
            for i, wd in enumerate(widths, 1):
                ws.column_dimensions[get_column_letter(i)].width = wd

            fn = "作业记录_" + self._safe_name(w.get("duid") or "DUID") + "_" + \
                datetime.datetime.now().strftime("%Y%m%d_%H%M%S") + ".xlsx"
            path = os.path.join(EXPORT_DIR, fn)
            wb.save(path)
            return {"ok": True, "path": path, "name": fn}
        except Exception as e:
            return {"ok": False, "error": str(e)}


if __name__ == "__main__":
    ensure_webview_runtimes()
    api = Api()
    with open(os.path.join(resource_dir(), "index.html"), encoding="utf-8") as _f:
        _index_html = _f.read()
    window = webview.create_window(
        "MRDN 系统 · 业务操作手册",
        html=_index_html,
        js_api=api,
        width=1320,
        height=860,
        min_size=(1000, 680),
    )
    webview.start(debug=os.environ.get("MRDN_DEBUG") == "1")
