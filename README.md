# SolidWorks BOM自动化处理系统

**当前稳定版本：v1.0（发布日期：2025-09-02）**

## 项目简介

基于SolidWorks 2019 VBA开发的BOM清单递归导出、零件数量递归统计及汇总表生成工具，满足生产环境中对装配体层级结构拆解、零件用量核算的自动化需求。

## 主要功能

1. **BOM递归导出**：自动识别装配体层级关系，递归导出各级BOM表为Excel文件（含缩略图）
2. **数量统计**：精确计算底层零件在整个装配体中的总数量（支持多层嵌套乘积）
3. **汇总表生成**：生成包含所有底层零件总数量和计算过程的汇总表
4. **自定义属性批量添加/更新（前置工具）**：对零件/装配体的“自定义属性（文档级）”与“配置特定（当前配置）”批量写入常用字段。支持单文件和文件夹批量处理，带可编辑确认界面与逐项勾选导入。

## 系统要求

- Windows 10专业版
- SolidWorks 2019
- 支持未激活的Excel 2016（使用SolidWorks内置导出，避免Excel COM依赖）

## 安装与使用

### 方法一：SolidWorks宏编辑器
1. 打开SolidWorks 2019
2. 菜单栏：工具 → 宏 → 编辑
3. 在VBA编辑器中：
   - 插入 → 模块，分别创建 Main、RecursiveProcessor、Utils、Logger 四个模块
   - 将对应的`.bas`文件内容复制粘贴到各模块中
4. 按F5运行`Main.RunExportNestedBOM`

### 方法二：导入宏文件
1. 将整个`src`文件夹复制到SolidWorks可访问的位置
2. 在SolidWorks中：工具 → 宏 → 运行
3. 选择`Main.bas`并运行`RunExportNestedBOM`函数

### 方法三：前置“自定义属性批量添加/更新”宏
该宏用于在导出BOM/统计前，对模型文件补齐或修正自定义属性。  
- 源码文件：
  - 窗体：src/frmProps.frm
  - 宏：src/UpdateProperties.bas（入口过程：Run_AddCustomProps）
- 运行方式（SolidWorks宏编辑器中）：
  1) 打开SolidWorks → 工具 → 宏 → 编辑；在VBA编辑器中导入窗体“frmProps”和模块“UpdateProperties.bas”
  2) 运行 Run_AddCustomProps
  3) 在弹窗中选择：
     - 是：处理当前打开文件
     - 否：选择文件夹并批量处理 *.sldprt、*.sldasm
- 确认窗体功能：
  - 每项“值”可直接修改，右侧勾选控制是否导入；默认全选
  - 点击“Add/Update”写入后自动保存（批量模式），“Skip”跳过，“Cancel”中止批量
- 写入策略与范围：
  - 同步写入 文档级（自定义） 与 配置特定（当前配置）
  - 采用“先Delete再Add2（文本类型）”的强力写入，避免新增失败

### 与BOM导出/汇总的关系
- 建议先运行属性批量添加宏，确保模型属性完整一致，再执行BOM导出/数量统计/汇总。

## 使用说明

### 文件结构要求
- 所有相关的工程图文件(.slddrw)必须在同一文件夹
- 子装配体工程图命名规则：`[子装配体代号].slddrw`
- 例如：代号为"TS180-01Z-02Z"的子装配体，对应工程图为"TS180-01Z-02Z.slddrw"

### BOM表格式要求
程序会自动识别以下列：
- **数量列**：支持"数量"、"QTY"、"Qty"等
- **代号列**：支持"代号"、"PART NUMBER"、"Part Number"等
- **名称列**：支持"名称"、"PART NAME"、"Name"等  
- **是否组装列**：支持"是否组装"、"Is Assembly"、"组装"等，值为"是"/"Yes"/"Y"/"True"/"1"表示子装配体

### 运行流程
1. 运行宏后会弹出文件选择对话框
2. 选择顶层装配体工程图(.slddrw文件)
3. 程序自动：
   - 导出各级BOM表为Excel文件
   - 计算底层零件总数量
   - 生成汇总表：`[顶层装配体名称]_汇总.xls`

### 输出文件
- **各级BOM**: `[装配体名称].xls`（包含缩略图）
- **汇总表**: `[顶层装配体名称]_汇总.xls`（包含计算过程）
- **日志文件**: `[顶层装配体名称]_run.log`

## 测试案例

项目包含基于实际装配体的测试案例：
- 顶层：TS180-01Z-01Z 四柱平台
- 子装配体：TS180-01Z-02Z 立柱平台  
- 底层零件：各种标准件和自制件

## 技术特点

- **兼容性**：避免Excel COM依赖，使用SolidWorks内置导出
- **健壮性**：循环引用检测、递归深度控制（最大10层）
- **可追溯性**：详细的计算过程记录和日志输出
- **容错性**：智能列识别，缺失列时提供默认处理

## 故障排除

### 常见问题
1. **"未找到BOM表"**: 确保工程图包含BOM表且表格类型为"BomFeat"
2. **"未找到子装配体工程图"**: 检查子装配体工程图文件名是否与BOM中代号一致
3. **"循环引用"**: 检查装配体结构，避免A包含B，B又包含A的情况
4. **"列识别失败"**: 检查日志文件，程序会尝试默认列位置

### 日志查看
程序运行过程中会生成详细日志，位置：`[工程图文件夹]\[装配体名称]_run.log`

## 项目结构（补充：excel/ 目录）

```
ExportNestedBOM/
├── src/                           # VBA源码文件
│   ├── Main.bas                  # 主入口模块
│   ├── RecursiveProcessor.bas    # 递归处理核心逻辑
│   ├── Utils.bas                 # 工具函数
│   ├── Logger.bas                # 日志记录
│   ├── Constants.bas             # 系统常量定义
│   ├── Configuration.bas         # 配置管理模块
│   └── ExportNestedBOM.swp       # SolidWorks宏项目文件
├── test/                          # 测试相关文件
│   ├── 测试指南.md               # 详细测试说明
│   └── 单元测试模拟.bas           # 单元测试代码
├── doc/                           # 项目文档
│   └── SolidWorks BOM自动化处理系统需求规格说明书.md
├── deploy/                        # 部署相关文档
│   └── 部署说明.md               # 安装和部署指南
├── examples/                      # 示例和使用场景
│   └── 示例使用场景.md           # 详细使用示例
└── README.md                      # 项目说明文档
├── excel/                        # Excel 2016 运行的模块（BOM表整理与PDF导出）
│   ├── Config.bas
│   ├── Logger.bas
│   ├── Utils.bas
│   ├── SingleSheetFormatter.bas
│   ├── SummaryProcessor.bas
│   ├── PdfExport.bas
│   └── Main.bas
```

## 开发者信息

本项目采用软件工程最佳实践：
- 模块化设计，职责分离
- 完整的错误处理和日志记录
- 详细的代码注释和文档
- 基于实际测试数据的验证

## 最近更新

### v1.0 (2024-07-28)
- 修复生产环境汇总表中文乱码问题（HTML输出编码改为gb2312）
- 支持"代号 名称.slddrw"格式的工程图文件查找
- BOM列识别支持"零件号"作为代号列关键词
- 优化错误处理和日志记录

## 贡献与反馈

欢迎通过GitHub Issues提交问题反馈或功能建议。如需贡献代码，请提交Pull Request。

## 许可证

本项目采用MIT许可证。详见LICENSE文件。

本项目仅供学习和研究使用。

### 自定义属性功能常见问题
1. “写入失败或新增属性未生效”
   - 该宏内部采用“Delete + Add2（文本类型）”策略，同时对文档级与配置特定写入；如仍失败，请确认文件是否可写、是否被只读锁定。
2. “批量模式中断”
   - 若单个文件点击“Cancel”，将中止批量；日志或提示框会统计“已更新/已跳过/用户中止”数量。
3. “质量/是否钣金为空”
   - 质量需要模型有材料与质量特性；是否钣金仅对零件类型判断，装配体恒为否。

## 子装配参与性确认（生产安全开关）

为避免子装配漏计或静默跳过，导出/统计前会弹出“参与性确认”列表，显示每个子装配的：
- 是否装配属性值（是否为“是”）
- 是否存在工程图
- 工程图中BOM表数量
- 参与状态（Included/Skipped-PropertyMissing/Skipped-NoDrawing/Skipped-NoBOMTable/Skipped-ExportError）

可执行：
- 导出CSV检查表（给项目经理/设计人员批量核对）
- 选择“继续执行”或“中止修复”
- 按需在配置中设置阻断策略与告警行为

提示：
- “覆盖率”与“统计完整性差距”指标会在汇总中展示，便于快速评估导出质量。

## Excel 2016 模块（BOM表整理与PDF导出）

说明：
- 本模块在 Excel 2016 内运行，无需依赖 SolidWorks 运行时，完成“单表规范化、PDF 导出、总汇总与分类生成”。
- 模块位置：excel/ 目录。采用 Late Binding，无需在 Excel 勾选外部引用。

模块清单与职责：
- Config.bas：全局配置（映射表路径、布尔真值、字体、打印/PDF、是否启用 PDF 合并等）
- Logger.bas：日志输出到工作簿同级 logs/ 目录
- Utils.bas：通用函数（标题别名、列重排、打印/页眉页脚、目录/文件工具）
- SingleSheetFormatter.bas：单工作表 S1-S8 规范化（映射替换、标题重命名、列重排、布尔图标化、字体与对齐、打印设置）
- SummaryProcessor.bas：以“汇总”驱动生成“总 BOM 清单”和分类表（外购件/钣金件/机箱模型），并规范化格式
- PdfExport.bas：按工作表导出为 PDF（默认使用 ExportAsFixedFormat），预留 PDFCreator 合并占位
- Main.bas：入口过程（Run_FormatAndExport_CurrentWorkbook / Run_BuildSummary_And_Export / Run_FullPipeline）

在 Excel 导入与运行：
1) 打开 Excel 2016，Alt+F11 打开 VBA 编辑器
2) 工程右键 → 导入文件 → 依次导入 excel 目录下 .bas 文件
3) Excel 菜单“开发工具” → 宏 → 运行以下入口之一：
   - Main.Run_FormatAndExport_CurrentWorkbook：仅规范化当前工作簿的所有数据表并导出各表 PDF
   - Main.Run_BuildSummary_And_Export：基于当前工作簿“汇总”表，生成“总 BOM 清单”和分类表，并导出 PDF
   - Main.Run_FullPipeline：先规范化，再生成总汇总与分类，最后导出 PDF

配置要点（可在 Config.bas 中修改）：
- CFG_MAPPING_WORKBOOK_PATH/CFG_MAPPING_SHEET：Toolbox 名映射表路径与工作表名（默认 TS180/lists.xlsm、ToolboxNames）
- CFG_BooleanTrueValues/CFG_Icon_True/CFG_Icon_False：布尔真值集合与展示图标（○/×）
- 字体与回退：CFG_Font_Primary/CFG_Font_Fallback
- 输出目录：CFG_PDF_OutputDir（默认“PDF”）
- 打印设置：A4 横向、100% 缩放、B:O 打印区域与页眉页脚（左=目录名，中=文件名，右=最后修改日期）
- 合并策略：CFG_Enable_PDFCreator_Merge（默认 True；未检测到 PDFCreator COM 时自动降级为仅导出）

注意：
- 默认优先使用 ExportAsFixedFormat 导出 PDF；PDFCreator 合并因版本差异保留为占位提示（可按本机版本补全）。
- 列标题容错：已内置别名与“材 料”→“材料”的清洗。