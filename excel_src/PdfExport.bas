Option Explicit

' 导出单个工作表为 PDF，返回生成的完整文件路径
Public Function ExportWorksheetToPdf(ByVal ws As Worksheet) As String
    Dim wb As Workbook: Set wb = ws.Parent
    Dim dirPath As String: dirPath = Utils.WorkbookDir(wb)
    Dim outDir As String: outDir = dirPath & Application.PathSeparator & CFG_PDF_OutputDir

    On Error Resume Next
    If Dir(outDir, vbDirectory) = "" Then MkDir outDir
    On Error GoTo 0

    Dim baseName As String
    If InStrRev(wb.Name, ".") > 0 Then
        baseName = Left$(wb.Name, InStrRev(wb.Name, ".") - 1)
    Else
        baseName = wb.Name
    End If
    Dim pdfPath As String
    pdfPath = outDir & Application.PathSeparator & baseName & "_" & ws.Name & ".pdf"

    ' 计算文档属性
    Dim leafFolder As String
    leafFolder = Utils.GetLeafFolderName(dirPath)
    Dim titleVal As String
    titleVal = leafFolder & " " & baseName
    Dim authorVal As String
    authorVal = Environ$("USERNAME")
    If Len(authorVal) = 0 Then authorVal = Application.UserName
    Dim subjectVal As String
    subjectVal = leafFolder

    ' 备份并设置内置文档属性（用于 PDF 元数据）
    Dim props As Object
    Set props = wb.BuiltinDocumentProperties
    Dim prevTitle As Variant, prevAuthor As Variant, prevSubject As Variant
    On Error Resume Next
    prevTitle = props("Title").Value
    prevAuthor = props("Author").Value
    prevSubject = props("Subject").Value
    props("Title").Value = titleVal
    props("Author").Value = authorVal
    props("Subject").Value = subjectVal
    On Error GoTo 0

    Logger.LogInfo "Export PDF: " & pdfPath & "; Title=" & titleVal & "; Author=" & authorVal & "; Subject=" & subjectVal

    ' 使用 ExportAsFixedFormat 导出 PDF（Excel 2010+ 可用）
    On Error GoTo EH
    ws.ExportAsFixedFormat Type:=xlTypePDF, Filename:=pdfPath, _
        Quality:=xlQualityStandard, IncludeDocProperties:=True, _
        IgnorePrintAreas:=False, OpenAfterPublish:=False

    ExportWorksheetToPdf = pdfPath

CLEANUP:
    ' 恢复原先的文档属性
    On Error Resume Next
    props("Title").Value = prevTitle
    props("Author").Value = prevAuthor
    props("Subject").Value = prevSubject
    On Error GoTo 0
    Exit Function
EH:
    Logger.LogError "Export PDF failed for sheet=" & ws.Name & ": " & Err.Description
    ExportWorksheetToPdf = ""
    Resume CLEANUP
End Function

' 预留：尝试合并一组 PDF（PDFCreator 可用时）。当前版本仅记录提示并返回 False。
Public Function TryMergePdfs(ByVal pdfList As Variant, ByVal outputPath As String) As Boolean
    ' 占位：根据本机 PDFCreator 版本进行 COM 接口调用；默认返回 False 表示未合并
    Logger.LogWarn "PDF merge not implemented; exported single PDFs only. Target would be: " & outputPath
    TryMergePdfs = False
End Function