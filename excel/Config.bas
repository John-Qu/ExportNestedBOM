' Attribute VB_Name = "Config"
Option Explicit

' 基本路径与外部资源
Public Const CFG_MAPPING_WORKBOOK_PATH As String = "FomatBOM_ExportPDF.xlsm"
Public Const CFG_MAPPING_SHEET As String = "ToolboxNames"
' 图标字符
Public Const CFG_Icon_True As String = "◉"
Public Const CFG_Icon_False As String = "✕"

' 字体配置（不可用时 Utils 里会回退）
Public Const CFG_Font_Primary As String = "汉仪长仿宋体"
Public Const CFG_Font_Fallback As String = "宋体"

' PDF 输出目录名（相对工作簿目录）
Public Const CFG_PDF_OutputDir As String = "PDF"

' 打印设置
Public Const CFG_Page_PaperSizeA4 As Integer = 9  ' xlPaperA4
Public Const CFG_Page_OrientationLandscape As Integer = 2 ' xlLandscape
Public Const CFG_Page_Zoom As Integer = 100

' 是否启用 PDFCreator 合并（若机器未安装则在运行时自动降级为 False）
Public Const CFG_Enable_PDFCreator_Merge As Boolean = True

' 关键词（机箱模型分类）
Public Function CFG_Keywords_Enclosure() As Variant
    CFG_Keywords_Enclosure = Array("机箱", "箱体")
End Function

' 布尔真值集合（大小写不敏感、去空白后比较）
Public Function CFG_BooleanTrueValues() As Variant
    CFG_BooleanTrueValues = Array("是", "yes", "y", "j", "shi", "要")
End Function

' 列别名与最终列序（规范）
' 最终列序 From A to R:
' A 零件号, B 文档预览, C 序号, D 代号, E 名称, F 数量, G 材料, H 处理,
' I 渠道, J 型号, K 组, L 购, M 加, N 钣, O 备注, P 零件名称, Q 规格, R 标准
Public Function CFG_FinalColumnOrder() As Variant
    CFG_FinalColumnOrder = Array( _
        "零件号", "文档预览", "序号", "代号", "名称", "数量", "材料", "处理", _
        "渠道", "型号", "组", "购", "加", "钣", "备注", "零件名称", "规格", "标准" _
    )
End Function

' 列标题重命名规则（字典：原标题->新标题）
Public Function CFG_HeaderRenamePairs() As Object
    Dim dict As Object: Set dict = CreateObject("Scripting.Dictionary")
    dict.CompareMode = 1 ' vbTextCompare
    dict("是否组装") = "组"
    dict("是否外购") = "购"
    dict("是否机加") = "加"
    dict("是否钣金") = "钣"
    dict("SUPPLIER") = "渠道"
    dict("材     料") = "材料" ' 去掉中间空格
    Set CFG_HeaderRenamePairs = dict
End Function

' 标题别名（用于兼容不同导出模板）
Public Function CFG_HeaderAliases() As Object
    Dim m As Object: Set m = CreateObject("Scripting.Dictionary")
    m.CompareMode = 1

    ' 每个标准名 -> 别名数组
    m("文档预览") = Array("预览", "Document Preview", "Preview")
    m("序号") = Array("序号", "Index", "No.", "项目号")
    m("零件号") = Array("零件号", "PartNo", "Part No", "Part Number", "PART NUMBER", "零件编号")
    m("代号") = Array("代号", "Code", "Key")
    m("名称") = Array("名称", "Name", "Description")
    m("数量") = Array("数量", "Qty", "Quantity")
    m("材料") = Array("材料", "材     料", "Material")
    m("处理") = Array("处理", "处理方式", "Treatment")
    m("渠道") = Array("渠道", "Supplier", "SUPPLIER")
    m("型号") = Array("型号", "Model", "Type")
    m("组") = Array("组", "是否组装", "Assemble")
    m("购") = Array("购", "是否外购", "Buy")
    m("加") = Array("加", "是否机加", "Machine")
    m("钣") = Array("钣", "是否钣金", "Sheet")
    m("备注") = Array("备注", "Remark", "Remarks", "Note")
    m("零件名称") = Array("零件名称", "PART NAME", "Component Name")
    m("规格") = Array("规格", "Spec", "SPECIFICATION")
    m("标准") = Array("标准", "Standard", "STANDARD")
    CFG_HeaderAliases = m
End Function

' 格式设置中的对齐策略
Public Function CFG_AlignLeftHeaders() As Variant
    CFG_AlignLeftHeaders = Array("代号", "型号", "处理", "备注")
End Function
Public Function CFG_AlignRightHeaders() As Variant
    CFG_AlignRightHeaders = Array("数量")
End Function

