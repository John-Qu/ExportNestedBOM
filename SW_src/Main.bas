' Main.bas - Entry point for ExportNestedBOM
Option Explicit

Public Sub RunExportNestedBOM()
    On Error GoTo EH
    
    ' 初始化配置
    InitializeConfig
    
    ' 验证系统环境
    Dim envIssues As String
    envIssues = ValidateEnvironment()
    If Len(envIssues) > 0 Then
        MsgBox "系统环境检查发现问题：" & vbCrLf & envIssues & vbCrLf & _
               "是否继续运行？", vbExclamation + vbYesNo, "环境检查"
        If vbNo = MsgBox("", vbYesNo) Then Exit Sub
    End If
    
    Dim swApp As Object ' SldWorks.SldWorks
    Set swApp = Application.SldWorks
    
    If swApp Is Nothing Then
        MsgBox "无法连接到SolidWorks应用程序，请确保SolidWorks已正常启动。", vbCritical
        Exit Sub
    End If
    
    Dim drawingPath As String
    drawingPath = GetTopLevelDrawingPath(swApp)
    If Len(drawingPath) = 0 Then
        MsgBox "未选择工程图文件，操作已取消。", vbExclamation
        Exit Sub
    End If
    
    ' 验证文件存在性
    If Not FileExists(drawingPath) Then
        MsgBox "选择的文件不存在：" & vbCrLf & drawingPath, vbCritical
        Exit Sub
    End If
    
    Dim logPath As String
    logPath = GetLogPathFromDrawing(drawingPath)
    Logger_Init logPath
    Logger_Info "=== 开始处理 ===" 
    Logger_Info "顶层工程图：" & drawingPath
    Logger_Info "SolidWorks版本：" & swApp.RevisionNumber
    
    Dim summary As Object ' Scripting.Dictionary -> partKey -> item dict
    Set summary = CreateObject("Scripting.Dictionary")
    
    Dim visited As Object: Set visited = CreateObject("Scripting.Dictionary")
    
    Dim topAsmName As String
    topAsmName = GetFileNameNoExt(drawingPath)
    
    Dim startTime As Double: startTime = Timer

    ' 新增：导出前参与性确认（可阻断）
    If CONFIRM_BEFORE_EXPORT Then
        Dim okConfirm As Boolean
        okConfirm = ConfirmSubAssemblyParticipation(swApp, drawingPath)
        If Not okConfirm Then
            Logger_Warn "用户在参与性确认阶段取消或存在阻断性问题，流程中止。"
            Exit Sub
        End If
    End If

    ProcessDrawingRecursive swApp, drawingPath, 1, 0, visited, summary, topAsmName, ""
    
    Dim endTime As Double: endTime = Timer
    Logger_Info "递归处理耗时：" & Format(endTime - startTime, "0.00") & " 秒"
    
    ' 检查是否有汇总数据
    If summary.Count = 0 Then
        Logger_Warn "未发现任何底层零件，请检查BOM表结构和'是否组装'列设置"
        MsgBox "处理完成，但未发现底层零件。请检查：" & vbCrLf & _
               "1. BOM表格式是否正确" & vbCrLf & _
               "2. '是否组装'列是否正确标记", vbExclamation
        Exit Sub
    End If
    
    ' 输出汇总表为xls(HTML)
    Dim outFolder As String: outFolder = GetFileFolder(drawingPath)
    Dim summaryXls As String
    summaryXls = outFolder & "\" & topAsmName & "_汇总.xls"
    
    On Error GoTo EH_Summary
    WriteSummaryHtmlXls summary, summaryXls
    Logger_Info "汇总输出：" & summaryXls & " (包含 " & summary.Count & " 种底层零件)"
    
    Logger_Info "=== 处理完成 ==="
    MsgBox "处理完成：" & vbCrLf & _
           "顶层工程图：" & drawingPath & vbCrLf & _
           "汇总表：" & summaryXls & vbCrLf & _
           "底层零件种类：" & summary.Count & vbCrLf & _
           "详细日志：" & logPath, vbInformation
    Exit Sub
    
EH_Summary:
    Logger_Error "汇总表生成失败：" & Err.Number & ": " & Err.Description
    MsgBox "汇总表生成失败，但BOM导出可能已完成。" & vbCrLf & _
           "错误：" & Err.Description & vbCrLf & _
           "请检查输出目录权限。", vbExclamation
    Exit Sub
    
EH:
    Dim errMsg As String
    errMsg = "程序执行出错：" & vbCrLf & _
             "错误代码：" & Err.Number & vbCrLf & _
             "错误描述：" & Err.Description & vbCrLf & _
             "错误来源：" & Err.Source
    
    Logger_Error "RunExportNestedBOM 出错：" & Err.Number & ": " & Err.Description & " (来源:" & Err.Source & ")"
    MsgBox errMsg, vbCritical
End Sub

Private Function GetTopLevelDrawingPath(swApp As Object) As String
    Dim filters As String
    filters = "工程图 (*.slddrw)|*.slddrw|所有文件 (*.*)|*.*"
    Dim opts As Long, cfg As String, disp As String
    On Error Resume Next
    GetTopLevelDrawingPath = swApp.GetOpenFileName("选择顶层装配体工程图", "", filters, opts, cfg, disp)
    On Error GoTo 0
    If Len(GetTopLevelDrawingPath) = 0 Then
        ' 如果用户未选择，则尝试使用当前活动文档
        Dim act As Object
        Set act = swApp.ActiveDoc
        If Not act Is Nothing Then
            If act.GetType = 3 Then ' swDocDRAWING
                GetTopLevelDrawingPath = act.GetPathName
            End If
        End If
    End If
End Function