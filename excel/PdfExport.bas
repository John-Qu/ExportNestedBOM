Attribute VB_Name = "PdfExport"
Option Explicit

' 将某个工作表导出为 PDF 到默认目录
Public Sub ExportWorksheetToPdf(ByVal ws As Worksheet, Optional ByVal outputDir As String = "")
    On Error GoTo EH
    If Len(outputDir) = 0 Then
        outputDir = ActiveWbDir() & Application.PathSeparator & CFG_PDF_OutputDir
    End If
    Call EnsureDirExists(outputDir)

    Dim fname As String
    fname = CleanFileName(ws.Parent.Name & "_" & ws.Name & ".pdf")

    Dim fullPath As String
    fullPath = outputDir & Application.PathSeparator & fname

    ws.ExportAsFixedFormat Type:=xlTypePDF, Filename:=fullPath, _
        Quality:=xlQualityStandard, IncludeDocProperties:=True, IgnorePrintAreas:=False, _
        OpenAfterPublish:=False

    LogInfo "已导出 PDF: " & fullPath
    Exit Sub
EH:
    LogError "ExportWorksheetToPdf 失败: " & Err.Description
End Sub

' 导出工作簿中所有可见工作表
Public Sub ExportWorkbookSheetsToPdf(ByVal wb As Workbook, Optional ByVal outputDir As String = "")
    On Error GoTo EH
    If Len(outputDir) = 0 Then outputDir = ActiveWbDir() & Application.PathSeparator & CFG_PDF_OutputDir
    EnsureDirExists outputDir

    Dim ws As Worksheet
    For Each ws In wb.Worksheets
        If ws.Visible = xlSheetVisible Then
            ExportWorksheetToPdf ws, outputDir
        End If
    Next ws
    Exit Sub
EH:
    LogError "ExportWorkbookSheetsToPdf 失败: " & Err.Description
End Sub

' 合并 PDF（优先 PDFCreator；若不可用则降级为记录提示）
' 注意：PDFCreator 需本机安装并启用 COM，接口版本存在差异，此处提供占位流程
Public Sub TryMergePdfs(ByVal outputMergedPdf As String, ByVal pdfList As Collection)
    On Error GoTo EH

    If pdfList Is Nothing Or pdfList.Count = 0 Then
        LogWarn "TryMergePdfs: 无需合并，列表为空"
        Exit Sub
    End If

    If Not CFG_Enable_PDFCreator_Merge Then
        LogWarn "TryMergePdfs: 未启用合并开关，跳过"
        Exit Sub
    End If

    ' 尝试后期绑定 PDFCreator 对象（v2+ JobQueue）
    Dim q As Object
    On Error Resume Next
    Set q = CreateObject("PDFCreator.JobQueue")
    On Error GoTo 0

    If q Is Nothing Then
        LogWarn "未检测到 PDFCreator COM 接口，跳过合并"
        Exit Sub
    End If

    ' 简化：使用 PDFCreator 合并队列（伪代码流程，部分环境需适配）
    ' 建议在生产环境中根据 PDFCreator 版本调整调用。
    Dim files() As String
    ReDim files(1 To pdfList.Count)
    Dim i As Long
    For i = 1 To pdfList.Count
        files(i) = CStr(pdfList(i))
    Next i

    ' 由于不同版本 API 差异，这里仅给出提示：
    LogWarn "检测到 PDFCreator，但由于版本差异，合并步骤未执行。请使用外部合并工具或根据本机版本完善接口调用。"
    ' 你可以在此处集成外部工具（如命令行 PDFtk 或 Ghostscript），或使用 PDFCreator 的具体 API 版本进行合并。

    Exit Sub
EH:
    LogError "TryMergePdfs 失败: " & Err.Description
End Sub

Private Function CleanFileName(ByVal s As String) As String
    Dim bad As Variant
    bad = Array("\", "/", ":", "*", "?", """", "<", ">", "|")
    Dim i As Long
    For i = LBound(bad) To UBound(bad)
        s = Replace(s, CStr(bad(i)), "_")
    Next i
    CleanFileName = s
End Function