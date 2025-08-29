' Utils.bas - utilities
Option Explicit

Public Function GetLogPathFromDrawing(drwPath As String) As String
    GetLogPathFromDrawing = GetFileFolder(drwPath) & "\" & GetFileNameNoExt(drwPath) & "_run.log"
End Function

Public Sub WriteSummaryHtmlXls(ByRef summary As Object, ByVal xlsPath As String)
    On Error GoTo EH
    
    ' 检查输出目录权限
    Dim outputDir As String: outputDir = GetFileFolder(xlsPath)
    If Not CanWriteToDirectory(outputDir) Then
        Err.Raise 70, "WriteSummaryHtmlXls", "无法写入目录：" & outputDir
    End If
    
    Dim f As Integer: f = FreeFile
    Dim k As Variant
    Open xlsPath For Output As #f
    
    Print #f, "<html><head><meta charset='utf-8'><style>"
    Print #f, "table{border-collapse:collapse;width:100%;font-family:Arial,sans-serif}"
    Print #f, "td,th{border:1px solid #999;padding:8px;text-align:left}"
    Print #f, "th{background-color:#f2f2f2;font-weight:bold}"
    Print #f, "tr:nth-child(even){background-color:#f9f9f9}"
    Print #f, "</style></head><body>"
    
    Print #f, "<h2>BOM汇总表</h2>"
    Print #f, "<p>生成时间：" & Now & "</p>"
    Print #f, "<p>零件种类：" & summary.Count & "</p>"
    Print #f, "<table><tr><th>代号</th><th>名称</th><th>总数量</th><th>分解链</th></tr>"
    
    For Each k In summary.Keys
        Dim item As Object: Set item = summary(k)
        Print #f, "<tr><td>" & HtmlEncode(item("PartNo")) & "</td><td>" & HtmlEncode(item("PartName")) & _
                  "</td><td>" & CStr(item("TotalQty")) & "</td><td>" & HtmlEncode(item("Breakdown")) & "</td></tr>"
    Next
    
    Print #f, "</table></body></html>"
    Close #f
    
    ' 验证文件是否成功生成
    If Not FileExists(xlsPath) Then
        Err.Raise 53, "WriteSummaryHtmlXls", "汇总文件生成失败"
    End If
    Exit Sub
    
EH:
    On Error Resume Next
    Close #f
    On Error GoTo 0
    Err.Raise Err.Number, Err.Source, Err.Description
End Sub

Public Function CanWriteToDirectory(dirPath As String) As Boolean
    On Error Resume Next
    Dim testFile As String
    testFile = dirPath & "\test_write_" & Format(Now, "hhnnss") & ".tmp"
    Dim f As Integer: f = FreeFile
    Open testFile For Output As #f
    Print #f, "test"
    Close #f
    CanWriteToDirectory = FileExists(testFile)
    If CanWriteToDirectory Then Kill testFile
    On Error GoTo 0
End Function

Public Function HtmlEncode(ByVal s As String) As String
    s = Replace(s, "&", "&amp;")
    s = Replace(s, "<", "&lt;")
    s = Replace(s, ">", "&gt;")
    s = Replace(s, """", "&quot;")
    HtmlEncode = s
End Function

Public Function GetFileFolder(path As String) As String
    Dim i As Long
    For i = Len(path) To 1 Step -1
        If Mid$(path, i, 1) = "\" Or Mid$(path, i, 1) = "/" Then
            GetFileFolder = Left$(path, i - 1)
            Exit Function
        End If
    Next
    GetFileFolder = CurDir$()
End Function

Public Function GetFileNameNoExt(path As String) As String
    Dim i As Long, j As Long
    For i = Len(path) To 1 Step -1
        If Mid$(path, i, 1) = "." Then
            j = i - 1
            Exit For
        End If
    Next
    If j = 0 Then j = Len(path)
    For i = j To 1 Step -1
        If Mid$(path, i, 1) = "\" Or Mid$(path, i, 1) = "/" Then
            GetFileNameNoExt = Mid$(path, i + 1, j - i)
            Exit Function
        End If
    Next
    GetFileNameNoExt = Left$(path, j)
End Function

Public Function FileExists(path As String) As Boolean
    On Error Resume Next
    FileExists = (Len(Dir$(path)) > 0)
    On Error GoTo 0
End Function

Public Function ExtractPartCode(fileName As String) As String
    ' 从文件名中提取代号，以空格为分隔符取前半部分的字母和数字字符串
    Dim i As Long, result As String
    result = ""
    
    ' 查找第一个空格位置
    i = InStr(1, fileName, " ")
    
    ' 如果找到空格，取空格前的部分，否则使用整个文件名
    If i > 0 Then
        result = Left$(fileName, i - 1)
    Else
        result = fileName
    End If
    
    ' 只保留字母和数字字符
    Dim j As Long, ch As String, cleanResult As String
    cleanResult = ""
    
    For j = 1 To Len(result)
        ch = Mid$(result, j, 1)
        ' 如果是字母或数字，则保留
        If (ch >= "A" And ch <= "Z") Or (ch >= "a" And ch <= "z") Or (ch >= "0" And ch <= "9") Or ch = "-" Then
            cleanResult = cleanResult & ch
        End If
    Next j
    
    ExtractPartCode = cleanResult
End Function
