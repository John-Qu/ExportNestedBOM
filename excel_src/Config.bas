Option Explicit

Public Const CFG_MAPPING_WORKBOOK_NAME As String = "FomatBOM_ExportPDF.xlsm"
Public Const CFG_MAPPING_SHEET_NAME As String = "ToolboxNames"
Public Const CFG_LOG_FOLDER As String = "logs"
Public Const CFG_HEADER_SCAN_MAX_ROWS As Long = 10
Public Const CFG_HEADER_DUMP_ROWS As Long = 6
Public Const CFG_HEADER_DUMP_COLS As Long = 20

Public Function CFG_TRUE_SET() As Variant
    CFG_TRUE_SET = Array("是", "yes", "y", "j", "shi", "要")
End Function

Public Property Get ICON_TRUE() As String
    ICON_TRUE = "●"
End Property

Public Property Get ICON_FALSE() As String
    ICON_FALSE = "X"
End Property