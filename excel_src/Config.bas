Option Explicit

Public Const CFG_MAPPING_WORKBOOK_NAME As String = "FomatBOM_ExportPDF.xlsm"
Public Const CFG_MAPPING_SHEET_NAME As String = "ToolboxNames"
Public Const CFG_LOG_FOLDER As String = "logs"
Public Const CFG_HEADER_SCAN_MAX_ROWS As Long = 10
Public Const CFG_HEADER_DUMP_ROWS As Long = 6
Public Const CFG_HEADER_DUMP_COLS As Long = 20

' === PDF 导出相关配置 ===
Public Const CFG_PDF_OutputDir As String = "PDF" ' 相对于工作簿目录
Public Const CFG_Enable_PDFCreator_Merge As Boolean = True ' 预留：若检测不到 PDFCreator COM，将自动降级为单表导出

Public Function CFG_TRUE_SET() As Variant
    CFG_TRUE_SET = Array("是", "yes", "y", "j", "shi", "要")
End Function

Public Property Get ICON_TRUE() As String
    ICON_TRUE = "●"
End Property

Public Property Get ICON_FALSE() As String
    ICON_FALSE = "X"
End Property