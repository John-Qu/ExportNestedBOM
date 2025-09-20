Option Explicit

Public Sub T1_Run_ToolboxReplace_ActiveSheet()
    Dim wb As Workbook: Set wb = Utils.ResolveTargetWorkbook()
    If wb Is Nothing Then
        MsgBox "未找到可处理的目标工作簿（请打开要处理的 .xls 文件）", vbExclamation
        Exit Sub
    End If

    Dim dirPath As String: dirPath = Utils.WorkbookDir(wb)
    Logger.LogInit dirPath, "T1"
    Logger.LogInfo "Target Workbook=" & wb.Name & ", Dir=" & dirPath

    Dim mapping As Object
    Set mapping = Utils.LoadToolboxMapping(dirPath)
    Logger.LogInfo "Mapping entries=" & mapping.Count

    Dim replaced As Long, unmatched As Long
    On Error GoTo FIN
    ' 激活目标工作簿，以便获取其活动工作表
    wb.Activate
    Dim ws As Worksheet: Set ws = ActiveSheet
    Logger.LogInfo "Processing ActiveSheet=" & ws.Name
    ' 先执行列标题重命名与列顺序调整（T2）
    SingleSheetFormatter.RenameHeadersAndReorder ws
    ' 再执行用例T1的替换
    SingleSheetFormatter.ApplyToolboxNameReplacement ws, mapping, replaced, unmatched
    ' 执行用例T3：布尔图标化
    SingleSheetFormatter.IconizeBooleanFlags ws
    Logger.LogInfo "DONE: replaced=" & replaced & ", unmatched=" & unmatched & "; Log=" & Logger.LogPath
FIN:
    Logger.LogClose
End Sub

Public Sub T1_Run_ToolboxReplace_AllVisibleSheets()
    Dim wb As Workbook: Set wb = Utils.ResolveTargetWorkbook()
    If wb Is Nothing Then
        MsgBox "未找到可处理的目标工作簿（请打开要处理的 .xls 文件）", vbExclamation
        Exit Sub
    End If

    Dim dirPath As String: dirPath = Utils.WorkbookDir(wb)
    Logger.LogInit dirPath, "T1"
    Logger.LogInfo "Target Workbook=" & wb.Name & ", Dir=" & dirPath

    Dim ws As Worksheet
    For Each ws In wb.Worksheets
        If ws.Visible = xlSheetVisible Then
            On Error Resume Next
            SingleSheetFormatter.RenameHeadersAndReorder ws
            Dim mapping As Object: Set mapping = Utils.LoadToolboxMapping(dirPath)
            Dim replaced As Long, unmatched As Long
            SingleSheetFormatter.ApplyToolboxNameReplacement ws, mapping, replaced, unmatched
            SingleSheetFormatter.IconizeBooleanFlags ws
            On Error GoTo 0
        End If
    Next ws
    Logger.LogInfo "DONE: all visible sheets processed; Log=" & Logger.LogPath
    Logger.LogClose
End Sub

' 新增：运行格式与打印（S2/S3/S6/S7）
Public Sub Run_Format_CurrentSheet()
    Dim wb As Workbook: Set wb = Utils.ResolveTargetWorkbook()
    If wb Is Nothing Then
        MsgBox "未找到可处理的目标工作簿（请打开要处理的 .xls 文件）", vbExclamation
        Exit Sub
    End If

    Dim dirPath As String: dirPath = Utils.WorkbookDir(wb)
    Logger.LogInit dirPath, "T4"
    On Error GoTo FIN
    wb.Activate
    Dim ws As Worksheet: Set ws = ActiveSheet
    Logger.LogInfo "Formatting ActiveSheet=" & ws.Name
    SingleSheetFormatter.FormatSingleBOMSheet ws
    Logger.LogInfo "DONE: formatted sheet=" & ws.Name
FIN:
    Logger.LogClose
End Sub

Public Sub Run_Format_AllVisibleSheets()
    Dim wb As Workbook: Set wb = Utils.ResolveTargetWorkbook()
    If wb Is Nothing Then
        MsgBox "未找到可处理的目标工作簿（请打开要处理的 .xls 文件）", vbExclamation
        Exit Sub
    End If

    Dim dirPath As String: dirPath = Utils.WorkbookDir(wb)
    Logger.LogInit dirPath, "T4"
    On Error GoTo FIN
    Dim ws As Worksheet
    For Each ws In wb.Worksheets
        If ws.Visible = xlSheetVisible Then
            Logger.LogInfo "Formatting sheet=" & ws.Name
            SingleSheetFormatter.FormatSingleBOMSheet ws
        End If
    Next ws
    Logger.LogInfo "DONE: formatted all visible sheets"
FIN:
    Logger.LogClose
End Sub

' 新增：格式化并导出当前工作簿所有可见工作表为 PDF（T5/E4）
Public Sub Run_FormatAndExport_CurrentWorkbook()
    Dim wb As Workbook: Set wb = Utils.ResolveTargetWorkbook()
    If wb Is Nothing Then
        MsgBox "未找到可处理的目标工作簿（请打开要处理的 .xls 文件）", vbExclamation
        Exit Sub
    End If

    Dim dirPath As String: dirPath = Utils.WorkbookDir(wb)
    Logger.LogInit dirPath, "T5"
    On Error GoTo FIN
    Dim ws As Worksheet
    Dim pdfPath As String
    For Each ws In wb.Worksheets
        If ws.Visible = xlSheetVisible Then
            Logger.LogInfo "Formatting+Exporting sheet=" & ws.Name
            SingleSheetFormatter.FormatSingleBOMSheet ws
            pdfPath = PdfExport.ExportWorksheetToPdf(ws)
            If Len(pdfPath) > 0 Then
                Logger.LogInfo "PDF generated: " & pdfPath
            End If
        End If
    Next ws
    Logger.LogInfo "DONE: All visible sheets formatted and exported to PDF/"
FIN:
    Logger.LogClose
End Sub