# mock_data 说明

本目录是 **`data/` 的完整 mock 副本**：目录结构、文件名、Excel 列、JSON schema 与真实系统完全一致，
但所有业务内容（DUID、itemcode、合同号、项目码、站点、人员、平台 Cookie）均为虚构。

| 路径 | 对应真实文件 | 说明 |
|---|---|---|
| data.json | data/data.json | 应用主状态：flows(13 节点 MRDN 流程) + platforms(CPQ/ISC 平台凭证) + checklist + workItems + batches + prep |
| du_state.json | data/du_state.json | DU 锚点物料清单缓存（flow_id\|work_item_id → {duid, materials[]}） |
| 操作手册.xlsx | data/操作手册.xlsx | 应用启动时读取的 SOP 手册（节点总览/操作指导/输入项明细/工具配置/输出文件/初始清单 6 个 sheet） |
| inputs/合同信息表.xlsx | 同名 | 项目码/PO#/合同号/实体/贸易条款 对照表（库存与余量查询的筛选依据） |
| inputs/物料类型表.xlsx | 同名 | ITEM编码 → 中文描述 → 设备类型(Type 1/2/3) |
| inputs/02-物料类型PHItemType_MOCK.xlsx | 02-物料类型PHItemType_<时间戳>.xlsx | 平台导出的物料类型宽表（11 列，带提交人/时间） |
| inputs/物料设计文档.xlsx | 同名 | 上游设计的物料宽表（3 行表头：合计行/itemcode 行/表头行 + 数据行） |
| inputs/2025 EXP Material for configuration_MOCK.xlsx | 2025 EXP Material for configuration_<版本>.xlsx | 客户物料主表：Material summary / Site list / Item replace / Label Release / Site lookup |
| inputs/输入模板.xlsx | 同名 | 物料转置输入模板（12 列，含发包/替换规则列） |
| inputs/物料转置输出模板.xlsx | 同名 | 物料转置输出模板（13 列，多一列 Date ready） |
| inputs/输出模板.xlsx | 同名 | CPQ Spart 输出模板（16 列 + sheetParam 服务产品列表 + Instructions） |
| inputs/CPQ Spart导入模板.xlsx | 同名 | CPQ 平台 Spart 导入文件模板（16 列） |
| inputs/SmartMR模板.xlsx | 同名 | SmartMR 上传模板（Du Material 18 列 + ref_* 下拉引用表） |
| inputs/SmartMR输出模板.xlsx | 同名 | SmartMR 输出模板 |
| inputs/需求分析输出模板.xlsx | 同名 | 需求分析 17 列模板（第 3 行为各列来源说明） |
| inputs/DU任务清单.xlsx | 同名 | 4 sheet：DU 清单/需求分析输出表/工具规则/物料转置表（实际作业清单） |
| inputs/catchtask.xlsx | 同名 | 合同号 → DUID 对照 |
| inputs/空实例列表.xlsx | 同名 | 空实例ID + 合同号 |
| inputs/_platformcfg_CPQ_*.json | 同名 | CPQ 平台凭证（真实=登录 Cookie，mock=假值） |
| inputs/_platformcfg_ISC_*.json | 同名 | ISC 平台凭证 |
| inputs/example_duidlist.txt | _duidlist_*_<时间戳>.txt | DUID 清单（每行一个） |
| inputs/example_itemcodelist.txt | _itemcodelist_*_<时间戳>.txt | itemcode 清单（每行一个） |
| outputs/MOCKRUN_*/ | outputs/<工具名>_<时间戳>/ | 各工具运行输出：DU分析(3列)/库存余量(9列)/余量(14列)/CPQ导入表(16列)/对比差异(8列) |
| works/MOCKDU01_WL/ | works/<DUID>/ | 单 DU 作业归档：01_需求分析/02_物料清单制作/03_辅料包配置 下的输出件 |
| checklist_files/fMOCK001/CMOCK*/ | checklist_files/<flow_id>/<checklist_id>/ | 清单项文件（物料设计表 exp/ftk，下载工具落盘处） |
| shots/ | shots/ | UI 截图目录 |
| 导出/ | 导出/ | 手动导出的 Excel（操作手册_日期.xlsx） |
