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
    Logger.LogInit dirPath, "T1-all"
    Logger.LogInfo "Target Workbook=" & wb.Name & ", Dir=" & dirPath

    Dim mapping As Object
    Set mapping = Utils.LoadToolboxMapping(dirPath)
    Logger.LogInfo "Mapping entries=" & mapping.Count

    Dim ws As Worksheet
    Dim replaced As Long, unmatched As Long
    For Each ws In wb.Worksheets
        If ws.Visible = xlSheetVisible Then
            replaced = 0: unmatched = 0
            ' 先执行列标题重命名与列顺序调整（T2）
            SingleSheetFormatter.RenameHeadersAndReorder ws
            ' 再执行用例T1的替换
            SingleSheetFormatter.ApplyToolboxNameReplacement ws, mapping, replaced, unmatched
        End If
    Next ws

    Logger.LogInfo "Finished all visible sheets. Log=" & Logger.LogPath
    Logger.LogClose
End Sub