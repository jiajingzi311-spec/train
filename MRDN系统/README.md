# MRDN 系统（业务操作 SOP 工作台）

一个 **单机版业务流程工作台**：把一条 13 节点的业务流（MRDN：需求分析 → 物料清单 → 发包/余量 →
挑料 → 完结）拆成一个个「操作节点」，每个节点记录 **输入 / 输出**，节点之间可「引用上一节点的输出」
形成链路，节点上可挂一个外部工具（`tools/` 下的 exe），在 UI 里一键运行并查看日志。

> **本仓库是「可公开」版本**：源码完整，但所有业务数据都是 `mock_data/` 里的**同 schema 虚构数据**。
> 真实的 Cookie、合同、项目码、站点、人员信息一律不在仓库中（安全红线见下文）。

## 目录结构

```
MRDN系统/
├── main.py                  # 后端：Python + pywebview，暴露 js_api（Api 类），全部业务逻辑
├── index.html               # 前端：单文件 UI（流程条视图 / 操作台 / 平台登录抓 Cookie / 文件预览）
├── 使用说明.md              # 面向使用者的操作手册
├── MRDN系统.spec            # PyInstaller 打包配置（exe + _internal）
├── 打包.py / 打包.bat        # 一键打包（自动备份/恢复 data，失败可抢救）
├── update_config.example.json  # 可选：版本更新服务器地址示例（运行时放 update_config.json）
├── tool_src/
│   ├── mrdn_download.py     # 通用接口下载工具源码（--url/--save_dir/--referer/--out_dir）
│   └── mrdn_download.spec   # 其 PyInstaller 配置
├── tools/
│   ├── *.tool.json          # 每个工具 exe 的「输入/输出契约」声明（app 靠它做参数绑定）
│   ├── 工具开发指导.md       # 工具开发规范（凭证 JSON、HTTP 带 Cookie、输出约定）
│   ├── MRDN物料余量查询.txt   # 一个工具的业务规则文档（库存/余量筛选条件示例）
│   └── onebox_config.example.json  # 网盘下载工具的账号配置示例（真实文件不入库）
├── mock_data/               # ⚠️ data/ 的完整 mock 副本（结构与真实 data/ 一致，内容全假）
│   ├── data.json           # 应用主状态：flows + platforms + checklist + workItems + batches
│   ├── du_state.json        # DU 锚点物料缓存
│   ├── 操作手册.xlsx         # 应用启动时读取的 SOP 手册（6 个固定 sheet）
│   ├── inputs/  outputs/  works/  shots/  checklist_files/  导出/
│   └── README.md           # 每个 mock 文件 ↔ 真实文件的对照表
├── _inspect*.py             # 早期调试脚本（读 Excel 结构用，可选保留）
└── .gitignore               # 安全红线：data/、build/、dist/、tools/*.exe 永不入库
```

## 怎么运行

**开发模式**（需要 Python + WebView2 运行时）：

```bat
pip install pywebview openpyxl playwright requests
set MRDN_DEBUG=1        :: 可选，开调试
python main.py
```

**发布模式**：`python 打包.py`（或双击 打包.bat）→ 生成 `dist/MRDN系统/`（exe + _internal + data + tools），
整个文件夹拷到别的 Win10/11 机器双击 exe 即用，零环境依赖。

> 打包会**自动备份再恢复** `dist/.../data`，保证升级不丢已配置数据（见 打包.py 头部注释）。

## 架构要点（重建必读）

1. **前后端**：`main.py` 用 pywebview 加载 `index.html`，`Api` 类的每个 public 方法都是前端可调的
   `window.pywebview.api.<method>`。前端是单 HTML（无构建步骤），截图用 Ctrl+V 粘贴 base64 存 `data/shots/`。
2. **状态文件**：
   - `data/data.json` — 全部应用状态（流程/节点/清单项/批次/作业项/平台凭证）。节点 id 形如 `N009`，
     流程 id 形如 `fMOCK001`，清单项 id 形如 `C212526`，作业项 id 形如 `W031`，批次 id 形如 `B030`。
   - `data/du_state.json` — 独立持久化「每个作业项的完整物料锚点清单」（key = `flow_id|work_item_id`，
     value = `{duid, materials:[{itemcode,desc,qty,type,sitestype,project,fromItemcode,fromType,status,remark,done}]}`），
     输出文件丢失时用于恢复标注。
   - `data/操作手册.xlsx` — 启动时读取的 SOP 存档（sheet：节点总览/操作指导/输入项明细/工具配置/输出文件/初始清单）。
3. **工具调用契约**（`run_tool`）：
   - app 找到 `tools/<工具名>.exe`，读取 `tools/<同名>.tool.json` 得到输入/输出契约；
   - 实际执行：`exe <前端按契约绑定的参数> --out_dir data/outputs/<工具名>_<时间戳>/`（cwd=tools/，600s 超时）；
   - stdout+stderr 合并当运行日志显示；`--out_dir` 里新生成的文件自动登记为该次运行的输出（可在线预览）。
   - 契约里的 `type: "platform"` 输入项绑定的是**平台凭证 JSON**（`data/inputs/_platformcfg_<平台>_<时间戳>.json`，
     结构 `{name,url,method,headers:{Cookie,User-Agent},cookie,updatedAt}`）——应用内「平台登录」会打开
     Playwright 浏览器，用户登录后自动抓取 Cookie 存入该 JSON，工具 exe 再带着它去调平台接口。
4. **平台凭证**：`mock_data/data.json` 的 `platforms[]` 有两个平台（CPQ / ISC）。真实 Cookie 形如
   `JSESSIONID=...; x_sso_token=...; userId=...`，**仓库里只有 MOCK 假值**，重建时按此结构实现
   「一键登录抓 Cookie + 测试链接」即可。
5. **物料设计表下载**：`tool_src/mrdn_download.py` 是独立 exe，从预签名 URL / 接口下载文件到
   `data/ML_design/<exp|ftk>/`（先清空目录），清单项的 `download` 字段记录分享链接与落盘文件名
   （`data/checklist_files/<flow_id>/<checklist_id>/`）。

## MRDN 13 节点流程（`mock_data/data.json` 里已内置完整定义）

| seq | 节点 | 干什么 | 关键输入 | 关键输出 |
|---|---|---|---|---|
| 1 | N009 需求分析 | 从需求邮件/DUID 清单判定项目、合同类型、仓库 | DUID 清单、需求邮件 | DU 信息表（DU分析，3 列） |
| 2 | N010 物料清单制作 | 物料设计宽表透视转置 + 类型转换 + 老料替换 | 物料设计文档（checklist 引用） | 物料清单（13 列） |
| 3 | N014 辅料包配置 | 筛选需发包辅料 → 导入 CPQ → S2C 转换 → 发布 | 物料清单（引用 N010） | CPQ 导入表、余量结果 |
| 4 | N012 辅料配置推送 | 生成 SmartMR 模板（只装 Type2/3）上传，推送辅料 | 引用 N010 | SmartMR 文件 |
| 5 | N023 辅料不满足检查 | SmartMR 不满足的辅料在锚点清单标注，发供应经理调配 | — | 不满足处理方案 |
| 6 | N016 辅料 MOS 冻结 | 生成通知邮件 + welink 通知一线冻结 | 项目号/DUID/实例号 | 邮件与截图 |
| 7 | N024 物料表一次检查更新 | 核对在线文档版本一致性 | 首次物料清单（引用 N010） | 检查结论 |
| 8 | N017 辅料挑料 | 辅料挑料操作 | — | 挑料完成状态 |
| 9 | N018 主设备配置 | 主设备**库存查询**（ISC）+ **余量查询**（CPQ） | itemcode 清单、合同信息表、CPQ/ISC 凭证 | 库存余量查询结果(9列)、余量查询结果(14列) |
| 10 | N019 主设备 MOS 挂载 | 主设备 MOS 挂载 | — | 挂载完成截图 |
| 11 | N020 主设备挑料 | FIFO 挑料、核对合同一致性、提交 MR、导出 DN 单 | — | DN 头信息+详细单据 |
| 12 | N025 物料清单检查 | 导入最新物料表做差异对比 | 最新物料设计表、物料类型表 | 对比差异表(8列) |
| 13 | N021 DN 完结邮件 | 复制挑完料的 DU，发邮件 + welink 通知 | 引用 N020 | 完结通知 |

另有 N022 空实例配置（辅料包发布后主设备 1:1 对应的空实例处理，引用 `空实例列表.xlsx`）。

## mock 数据 ↔ 真实数据

`mock_data/` 与真实 `data/` **逐文件对应**（详见 `mock_data/README.md`）。虚构值约定：

- DUID：`MOCKDU01_WL` / `MOCKDU02_WL` / `MOCKDU03_L1800_WL`（真实格式 `SLxxxx_..._WL`）
- itemcode：`99000001`~`99000008`（真实为 8 位字母数字混合，如 Type 1/2/3 三类物料）
- 合同号：`00099000000010A`（真实为 14 位数字 + 1 位字母）
- 项目码：`99Z0AAA` / `99Z0BBB` / `99Z0CCC` / `99Z0DDD` / `99Z0EEE`
- PO#：`9990000001` 起；仓库：`MOCK_WH_A/B/C`；实体：`Mock Int'l` / `Mock Local`
- 平台域名：`*.mock-platform.example`；Cookie：`MOCK_JSESSIONID=...` 等假值
- 站点：`NS-MOCK24-A1`；空实例：`A900001 | AS999900001`；人员：`Mock User A`

## 安全红线（贡献者必读）

1. **`data/` 永不入库**（.gitignore 已排除）。真实运行数据含：平台登录 Cookie（= 身份凭证）、
   合同/PO 数据、客户站点与项目码、人员姓名——任何一项都不许进 Git。
2. **`tools/onebox_config.json`（网盘明文账号密码）永不入库**，仓库只放 `.example`。
3. 提交前自查：仓库全文不得出现 公司内部 intranet 域名后缀、真实项目码（7 位项目码）、真实合同号
   （0006 开头 14 位）、真实 DUID（SL 开头）、真实 itemcode、`JSESSIONID=<40位hex>`、
   `x_sso_token=` 后的真实长串、内部人员姓名。
4. `tools/*.exe`、`build/`、`dist/`、`_backups/` 均为构建产物，不进仓库（.gitignore 已排除）。

## 重建说明（给接手的 agent / 开发者）

- 要重建的对象 = **主应用**（main.py + index.html + spec + 打包链路，源码已全）+
  **tools 下的各工具 exe**（本仓库只有它们的 `.tool.json` 契约、规则文档与下载工具源码；
  各工具的内部实现逻辑以 `tools/MRDN物料余量查询.txt` 这类规则文档 + `工具开发指导.md` 为准）。
- 工具输入/输出 Excel 的**列结构**全部可从 `mock_data/` 直接读到（与真实完全一致）。
- 平台接口细节（ISC 库存 `findOnhandDetailList`、CPQ 余量 `getContractRemainInfo` 等）
  见 `tools/MRDN物料余量查询.txt` 第 6~18 条（真实系统里的接口名/参数保留，项目码与仓库名已 mock）。
- 运行验证：`python main.py` 启动后，UI 应能加载 `mock_data` 同构的 data.json（首次运行会按
  `DEFAULT_DATA` 生成演示数据；把 `mock_data/` 拷到 exe 同级 `data/` 即可复现示例状态）。
