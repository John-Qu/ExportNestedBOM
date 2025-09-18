Attribute VB_Name = "SingleSheetFormatter"
Option Explicit

' S1-S8：格式化当前工作簿的所有数据工作表（可按需过滤）
Public Sub FormatActiveWorkbookBOMSheets()
    On Error GoTo EH
    LogInit ActiveWbDir()
    Dim map As Object: Set map = LoadToolboxMapping()

    Dim wb As Workbook: Set wb = Application.ActiveWorkbook
    Dim ws As Worksheet
    For Each ws In wb.Worksheets
        ' 如需跳过“汇总”等表，可在此定制
        If LCase$(ws.Name) <> "汇总" And ws.Visible = xlSheetVisible Then
            FormatSingleBOMSheet ws, map
            LogInfo "已格式化工作表: " & ws.Name
        End If
    Next ws

    LogInfo "FormatActiveWorkbookBOMSheets 完成"
    GoTo DONE
EH:
    LogError "FormatActiveWorkbookBOMSheets 失败: " & Err.Description
DONE:
    LogClose
End Sub

' 对单个工作表执行 S1-S8
Public Sub FormatSingleBOMSheet(ByVal ws As Worksheet, ByVal toolboxMap As Object)
    On Error GoTo EH

    ' S2 Toolbox 名称替换与回填：K 零件名称命中映射 -> E 名称=中文名；L->J，M->I
    ApplyToolboxNameReplacement ws, toolboxMap

    ' S3 标题重命名
    ApplyHeaderRenameRules ws

    ' S4 列位置调整（整表 -> 目标列序）
    ReorderColumnsToSpec ws

    ' S5 布尔图标化（K/L/M/N）
    ApplyBooleanIconization ws

    ' S6 字体与对齐
    ApplyFontAndAlignment ws

    ' S7 打印设置
    ApplyPrintSetup ws

    ' S8 PDF 导出由 PdfExport 模块统一进行（此处只格式化）
    Exit Sub
EH:
    LogError "FormatSingleBOMSheet(" & ws.Name & ") 失败: " & Err.Description
End Sub

Private Sub ApplyToolboxNameReplacement(ByVal ws As Worksheet, ByVal toolboxMap As Object)
    On Error GoTo EH
    Dim lastRow As Long: lastRow = LastUsedRow(ws, 1)

    Dim colK As Long: colK = GetColumnIndex(ws, "零件名称")
    Dim colE As Long: colE = GetColumnIndex(ws, "名称")
    Dim colL As Long: colL = GetColumnIndex(ws, "规格")
    Dim colM As Long: colM = GetColumnIndex(ws, "标准")
    Dim colJ As Long: colJ = GetColumnIndex(ws, "型号")
    Dim colI As Long: colI = GetColumnIndex(ws, "渠道")

    If colK = 0 Then Exit Sub

    Dim r As Long
    For r = 2 To lastRow
        Dim partName As String: partName = Trim$(CStr(ws.Cells(r, colK).Value))
        If Len(partName) > 0 Then
            If toolboxMap.Exists(partName) Then
                If colE > 0 Then ws.Cells(r, colE).Value = toolboxMap(partName)
                If colL > 0 And colJ > 0 Then ws.Cells(r, colJ).Value = ws.Cells(r, colL).Value
                If colM > 0 And colI > 0 Then ws.Cells(r, colI).Value = ws.Cells(r, colM).Value
            End If
        End If
    Next r
    Exit Sub
EH:
    LogError "ApplyToolboxNameReplacement 失败: " & Err.Description
End Sub