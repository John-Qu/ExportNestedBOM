ExportNestedBOM

概述
- ExportNestedBOM 是一个面向 SolidWorks 与 Excel 的装配件 BOM 自动化处理工具集，支持在 Excel 中对从 SolidWorks 导出的子装配清单进行合并、格式化、生成总 BOM 清单，并按工作表分别导出 PDF。
- 兼容 macOS 与 Windows，已统一采用 Excel VBA 的 Application.PathSeparator 进行路径拼接，避免路径不兼容问题。

核心特性
- 合并子装配：将多个子装配清单合并到当前工作簿，自动解除共享以避免复制受限。
- 单表格式化：统一列标题与列序；将布尔列（组/购/加/钣）图标化显示；设置字体与对齐；打印设置；清理末尾无预览行；在“零件名称”前确保存在“备注”列。
- 总 BOM 生成：基于“汇总”工作表，直接在当前合并后的工作簿的各子装配工作表中查找数据（不再遍历文件夹）。
- 分表导出 PDF：每个工作表单独导出 PDF，文件名仅使用工作表名，保存到工作簿所在文件夹。
- 分表打印：每张工作表作为独立打印作业，页码 (&P) 从 1 开始计数。
- 日志记录：关键步骤输出日志，便于定位问题。

快速开始
1) 打开 Excel 工作簿（建议先将子装配清单合并到一个工作簿）。
2) 在宏安全允许的情况下运行以下入口宏（见下文“宏入口与说明”）：
   - Run_Merge_SubBOMs_Into_CurrentWorkbook（合并子装配）
   - Run_Format_AllVisibleSheets（批量格式化当前工作簿内所有可见工作表）
   - Run_Generate_TotalBOM_FromSummary（基于汇总表生成总 BOM）
   - Run_Export_AllSheets_ToPDF（分表导出 PDF）或 Run_Print_AllSheets_Separately（分表打印）
3) 在工作簿所在文件夹中查看生成的 PDF 文件与日志输出。

宏入口与说明（excel_src）
- Run_Format_CurrentSheet：格式化当前活动工作表，包含列重命名与列序、布尔图标化、字体与对齐、打印设置、清理末尾无预览行、确保“备注”位于“零件名称”前。
- Run_Format_AllVisibleSheets：格式化当前工作簿中所有可见工作表（同上）。
- Run_FormatAndExport_CurrentWorkbook：格式化并导出当前工作簿（组合流程）。
- Run_Generate_TotalBOM_FromSummary：从“汇总”工作表读取关键字段，在当前工作簿的子装配工作表中逐行查找并生成“总 BOM 清单”。
- Run_Merge_SubBOMs_Into_CurrentWorkbook：将目标文件夹内的子装配清单合并到当前工作簿（已修正路径分隔符，兼容 macOS/Windows）。
- Run_Print_AllSheets_Separately：按工作表分表打印，确保每张工作表的页码从 1 开始计数。
- Run_Export_AllSheets_ToPDF：按工作表分表导出 PDF，文件名仅使用工作表名，保存到工作簿所在文件夹。

SolidWorks 侧宏入口（SW_src）
- Run_AddCustomProps：向模型/工程图批量添加自定义属性，配合 Excel 侧流程生成规范的 BOM 数据。

项目结构
- excel_src/：Excel 侧 VBA 模块（Main、SingleSheetFormatter、SummaryProcessor、PdfExport、Utils、Config、Logger）。
- SW_src/：SolidWorks 侧 VBA 模块（Main、RecursiveProcessor、UpdateProperties、Utils、Logger、Configuration、Constants、窗体）。
- doc/：文档与规格说明（本次重写新增用户/开发者/架构/变更文档）。
- test/：示例数据与测试模型。

关键实现与约定
- 路径兼容：统一使用 Application.PathSeparator，禁止硬编码 “\” 或 “/”。
- 列序规范：最终列序为 [零件号, 文档预览, 序号, 代号, 名称, 数量, 材料, 处理, 渠道, 型号, 组, 购, 加, 钣, 备注, 零件名称, 规格, 标准]。
- 备注列保障：若“备注”缺失且存在“零件名称”，在其前插入空白“备注”列。
- 预览判定：“文档预览”列中，单元格文本非空或单元格范围内存在形状/OLEObject（中心点落在单元格内）视为“有预览”。末尾连续的“无预览”行会被清理。
- 总 BOM 查找：从当前工作簿内可见工作表中查找匹配项，排除“汇总”与“总 BOM 清单”两张表。

常见问题
- 子装配未复制：通常由路径分隔符不兼容导致枚举失败。已统一为 Application.PathSeparator。
- 页码连续：整本打印时 &P 为全作业连续；使用“分表打印”可让页码在每张工作表从 1 开始。
- PDF 文件名：仅使用工作表名，非法字符需自行处理（如替换 \ / : * ? " < > |）。

许可证
- MIT License（详见 LICENSE）。

更多文档
- 请查看 doc/ 用户使用指南、开发者指南、架构设计与变更日志。