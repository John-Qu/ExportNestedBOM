' Attribute VB_Name = "Utils"
Option Explicit

Public Function EnsureDirExists(ByVal path As String) As Boolean
    On Error GoTo EH
    If Len(Dir(path, vbDirectory)) = 0 Then
        MkDir path
    End If
    EnsureDirExists = True
    Exit Function
EH:
    EnsureDirExists = (Len(Dir(path, vbDirectory)) > 0)
End Function

Public Function ActiveWbDir() As String
    If Not Application Is Nothing And Not Application.ActiveWorkbook Is Nothing Then
        ActiveWbDir = Application.ActiveWorkbook.Path
    Else
        ActiveWbDir = CurDir
    End If
End Function

Public Function ResolvePath(ByVal relativeOrAbsolute As String) As String
    If Len(relativeOrAbsolute) = 0 Then
        ResolvePath = relativeOrAbsolute
    ElseIf InStr(1, relativeOrAbsolute, ":", vbTextCompare) > 0 Or Left(relativeOrAbsolute, 2) = "\\" Then
        ResolvePath = relativeOrAbsolute
    Else
        ResolvePath = ActiveWbDir() & Application.PathSeparator & relativeOrAbsolute
    End If
End Function

Public Function FileExists(ByVal fullPath As String) As Boolean
    FileExists = (Len(Dir(fullPath, vbNormal)) > 0)
End Function

Public Function GetFileName(ByVal fullPath As String) As String
    GetFileName = Dir(fullPath)
End Function

Public Function GetFolderNameFromPath(ByVal fullPath As String) As String
    Dim p As Long: p = InStrRev(fullPath, Application.PathSeparator)
    If p > 0 Then
        GetFolderNameFromPath = Mid(fullPath, 1, p - 1)
    Else
        GetFolderNameFromPath = fullPath
    End If
End Function

Public Function GetLeafFolderName(ByVal folderPath As String) As String
    Dim p As Long: p = InStrRev(folderPath, Application.PathSeparator)
    If p > 0 Then
        GetLeafFolderName = Mid(folderPath, p + 1)
    Else
        GetLeafFolderName = folderPath
    End If
End Function

Public Function GetUserNameSafe() As String
    On Error Resume Next
    GetUserNameSafe = Environ$("USERNAME")
    If Len(GetUserNameSafe) = 0 Then GetUserNameSafe = "UnknownUser"
    On Error GoTo 0
End Function

' 载入对照表：ToolboxNames，B 列=英文名，C 列=中文名
Public Function LoadToolboxMapping() As Object
    Dim dict As Object: Set dict = CreateObject("Scripting.Dictionary")
    dict.CompareMode = 1

    Dim fullPath As String
    fullPath = ResolvePath(CFG_MAPPING_WORKBOOK_PATH)
    If Not FileExists(fullPath) Then
        LogWarn "映射表文件不存在，跳过替换: " & fullPath
        Set LoadToolboxMapping = dict
        Exit Function
    End If

    Dim app As Application: Set app = Application
    Dim wb As Workbook
    Dim wasVisible As Boolean: wasVisible = app.Visible
    Dim wasAlerts As Boolean: wasAlerts = app.DisplayAlerts
    app.DisplayAlerts = False

    On Error GoTo CLEANUP
    Set wb = app.Workbooks.Open(fullPath, ReadOnly:=True)
    Dim ws As Worksheet
    Set ws = wb.Worksheets(CFG_MAPPING_SHEET)

    Dim r As Long, lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, 2).End(xlUp).Row
    For r = 2 To lastRow
        Dim en As String, zh As String
        en = Trim(CStr(ws.Cells(r, 2).Value))
        zh = Trim(CStr(ws.Cells(r, 3).Value))
        If Len(en) > 0 And Len(zh) > 0 Then
            If Not dict.Exists(en) Then dict.Add en, zh
        End If
    Next r
    LogInfo "映射表加载完成：" & GetFileName(fullPath) & "，条目数=" & CStr(dict.Count)

CLEANUP:
    On Error Resume Next
    If Not wb Is Nothing Then wb.Close SaveChanges:=False
    app.DisplayAlerts = wasAlerts
    app.Visible = wasVisible
    On Error GoTo 0

    Set LoadToolboxMapping = dict
End Function

' 寻找行列
Public Function LastUsedRow(ByVal ws As Worksheet, Optional ByVal col As Long = 1) As Long
    Dim r As Long: r = ws.Cells(ws.Rows.Count, col).End(xlUp).Row
    If r < 1 Then r = 1
    LastUsedRow = r
End Function

Public Function FindHeaderColumn(ByVal ws As Worksheet, ByVal headerName As String, Optional ByVal aliases As Object = Nothing) As Long
    Dim maxCol As Long: maxCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    Dim h As Long
    For h = 1 To maxCol
        Dim t As String: t = NormalizeHeader(CStr(ws.Cells(1, h).Value))
        If StrComp(t, NormalizeHeader(headerName), vbTextCompare) = 0 Then
            FindHeaderColumn = h
            Exit Function
        End If
        If Not aliases Is Nothing Then
            Dim arr As Variant
            If aliases.Exists(headerName) Then
                arr = aliases(headerName)
                Dim i As Long
                For i = LBound(arr) To UBound(arr)
                    If StrComp(t, NormalizeHeader(CStr(arr(i))), vbTextCompare) = 0 Then
                        FindHeaderColumn = h
                        Exit Function
                    End If
                Next i
            End If
        End If
    Next h
    FindHeaderColumn = 0
End Function

Public Function NormalizeHeader(ByVal s As String) As String
    Dim t As String: t = Trim$(s)
    t = Replace(t, vbCr, "")
    t = Replace(t, vbLf, "")
    t = Replace(t, Chr(160), " ")
    t = Application.WorksheetFunction.Substitute(t, "  ", " ")
    t = Replace(t, "材 料", "材料")
    NormalizeHeader = t
End Function

' 标题重命名（S3）
Public Sub ApplyHeaderRenameRules(ByVal ws As Worksheet)
    Dim maxCol As Long: maxCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    Dim dict As Object: Set dict = CFG_HeaderRenamePairs()
    Dim c As Long
    For c = 1 To maxCol
        Dim h As String: h = NormalizeHeader(CStr(ws.Cells(1, c).Value))
        If dict.Exists(h) Then
            ws.Cells(1, c).Value = dict(h)
        Else
            ' 单独处理 "材 料"
            If StrComp(h, "材     料", vbTextCompare) = 0 Then ws.Cells(1, c).Value = "材料"
        End If
    Next c
End Sub

' 按最终列序重排（S4）
Public Sub ReorderColumnsToSpec(ByVal ws As Worksheet)
    Dim targetOrder As Variant: targetOrder = CFG_FinalColumnOrder()
    Dim aliases As Object: Set aliases = CFG_HeaderAliases()

    Dim colMap() As Long
    ReDim colMap(LBound(targetOrder) To UBound(targetOrder))
    Dim i As Long

    ' 找到每个目标标题目前所在列
    For i = LBound(targetOrder) To UBound(targetOrder)
        colMap(i) = FindHeaderColumn(ws, CStr(targetOrder(i)), aliases)
    Next i

    ' 建临时工作表复制列避免复杂剪切操作
    Dim wb As Workbook: Set wb = ws.Parent
    Dim tmp As Worksheet
    Set tmp = wb.Worksheets.Add(After:=ws)
    tmp.Name = ws.Name & "_tmp_" & Format(Now, "hhnnss")

    ' 复制标题
    Dim outCol As Long: outCol = 1
    For i = LBound(targetOrder) To UBound(targetOrder)
        If colMap(i) > 0 Then
            ws.Columns(colMap(i)).Copy tmp.Columns(outCol)
            tmp.Cells(1, outCol).Value = CStr(targetOrder(i)) ' 标题规范化
            outCol = outCol + 1
        Else
            ' 若缺失该列，创建空列并写标题
            tmp.Cells(1, outCol).Value = CStr(targetOrder(i))
            outCol = outCol + 1
        End If
    Next i

    ' 将 tmp 内容复制回原表（清空后粘贴值和格式）
    ws.Cells.Clear
    Dim lastCol As Long: lastCol = tmp.Cells(1, tmp.Columns.Count).End(xlToLeft).Column
    Dim lastRow As Long: lastRow = tmp.Cells(tmp.Rows.Count, 1).End(xlUp).Row
    tmp.Range(tmp.Cells(1, 1), tmp.Cells(IIf(lastRow < 1, 1, lastRow), lastCol)).Copy ws.Range("A1")

    Application.DisplayAlerts = False
    tmp.Delete
    Application.DisplayAlerts = True
End Sub

' K/L/M/N 布尔图标化（S5）
Public Sub ApplyBooleanIconization(ByVal ws As Worksheet)
    Dim tf As Variant: tf = CFG_BooleanTrueValues()
    Dim trueSet As Object: Set trueSet = CreateObject("Scripting.Dictionary")
    trueSet.CompareMode = 1
    Dim i As Long
    For i = LBound(tf) To UBound(tf)
        trueSet(Trim$(LCase$(CStr(tf(i))))) = True
    Next i

    Dim headerList As Variant
    headerList = Array("组", "购", "加", "钣")

    Dim aliases As Object: Set aliases = CFG_HeaderAliases()
    Dim cols(0 To 3) As Long
    For i = 0 To 3
        cols(i) = FindHeaderColumn(ws, CStr(headerList(i)), aliases)
    Next i

    Dim lastRow As Long: lastRow = LastUsedRow(ws, 1)
    Dim r As Long, c As Long
    For i = 0 To 3
        c = cols(i)
        If c > 0 Then
            For r = 2 To lastRow
                Dim raw As String
                raw = Trim$(CStr(ws.Cells(r, c).Value))
                Dim norm As String: norm = LCase$(Replace(raw, " ", ""))
                If Len(norm) = 0 Then
                    ws.Cells(r, c).Value = CFG_Icon_False
                ElseIf trueSet.Exists(norm) Then
                    ws.Cells(r, c).Value = CFG_Icon_True
                Else
                    ws.Cells(r, c).Value = CFG_Icon_False
                End If
            Next r
        End If
    Next i
End Sub

' S6 字体与对齐
Public Sub ApplyFontAndAlignment(ByVal ws As Worksheet)
    On Error Resume Next
    Dim used As Range: Set used = ws.UsedRange
    If used Is Nothing Then Exit Sub

    Dim fontName As String: fontName = CFG_Font_Primary
    used.Font.Name = fontName
    used.Font.Size = 14
    If Err.Number <> 0 Then
        Err.Clear
        used.Font.Name = CFG_Font_Fallback
        used.Font.Size = 12
    End If

    ' 标题行加粗、居中
    ws.Rows(1).Font.Bold = True
    ws.Rows(1).HorizontalAlignment = xlCenter
    ws.Rows(1).VerticalAlignment = xlCenter

    ' 对齐策略
    Dim lefts As Variant: lefts = CFG_AlignLeftHeaders()
    Dim rights As Variant: rights = CFG_AlignRightHeaders()
    Dim aliases As Object: Set aliases = CFG_HeaderAliases()

    Dim maxCol As Long: maxCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    Dim c As Long

    ' 默认中对齐
    used.HorizontalAlignment = xlCenter
    used.VerticalAlignment = xlCenter

    ' 左对齐
    For c = 1 To maxCol
        If IsHeaderIn(ws, c, lefts, aliases) Then
            ws.Range(ws.Cells(2, c), ws.Cells(LastUsedRow(ws, 1), c)).HorizontalAlignment = xlLeft
        End If
    Next c
    ' 右对齐
    For c = 1 To maxCol
        If IsHeaderIn(ws, c, rights, aliases) Then
            ws.Range(ws.Cells(2, c), ws.Cells(LastUsedRow(ws, 1), c)).HorizontalAlignment = xlRight
        End If
    Next c

    ' 去掉单元格边框
    used.Borders.LineStyle = xlLineStyleNone
    On Error GoTo 0
End Sub

Private Function IsHeaderIn(ByVal ws As Worksheet, ByVal col As Long, ByVal headerArr As Variant, ByVal aliases As Object) As Boolean
    Dim h As String: h = NormalizeHeader(CStr(ws.Cells(1, col).Value))
    Dim i As Long
    For i = LBound(headerArr) To UBound(headerArr)
        Dim target As String: target = NormalizeHeader(CStr(headerArr(i)))
        If StrComp(h, target, vbTextCompare) = 0 Then
            IsHeaderIn = True: Exit Function
        End If
        ' Check aliases of target
        If aliases.Exists(target) Then
            Dim al As Variant: al = aliases(target)
            Dim j As Long
            For j = LBound(al) To UBound(al)
                If StrComp(h, NormalizeHeader(CStr(al(j))), vbTextCompare) = 0 Then
                    IsHeaderIn = True: Exit Function
                End If
            Next j
        End If
    Next i
    IsHeaderIn = False
End Function

' S7 打印设置与页眉页脚
Public Sub ApplyPrintSetup(ByVal ws As Worksheet)
    Dim lastRow As Long: lastRow = LastUsedRow(ws, 1)
    With ws.PageSetup
        .PrintArea = ws.Range(ws.Cells(1, 2), ws.Cells(lastRow, 15)).Address  ' B:O
        .PrintTitleRows = "$1:$1"
        .Orientation = CFG_Page_OrientationLandscape
        .PaperSize = CFG_Page_PaperSizeA4
        .Zoom = CFG_Page_Zoom
        .LeftHeader = GetLeafFolderName(ActiveWbDir())
        .CenterHeader = ws.Parent.Name
        .RightHeader = Format(FileDateTime(ws.Parent.FullName), "yyyy-mm-dd")
        .CenterFooter = "第 &[页码] 页，共 &[总页数] 页"
        .LeftMargin = Application.CentimetersToPoints(1.5)
        .RightMargin = Application.CentimetersToPoints(1.5)
        .TopMargin = Application.CentimetersToPoints(1.5)
        .BottomMargin = Application.CentimetersToPoints(1.5)
        .HeaderMargin = Application.CentimetersToPoints(0.8)
        .FooterMargin = Application.CentimetersToPoints(0.8)
    End With
End Sub

' 列定位辅助
Public Function GetColumnIndex(ByVal ws As Worksheet, ByVal headerName As String) As Long
    GetColumnIndex = FindHeaderColumn(ws, headerName, CFG_HeaderAliases())
End Function

' 列复制（按标题）
Public Sub CopyRowByHeaders(ByVal srcWs As Worksheet, ByVal srcRow As Long, ByVal dstWs As Worksheet, ByVal dstRow As Long)
    Dim order As Variant: order = CFG_FinalColumnOrder()
    Dim i As Long
    For i = LBound(order) To UBound(order)
        Dim h As String: h = CStr(order(i))
        Dim srcCol As Long: srcCol = GetColumnIndex(srcWs, h)
        Dim dstCol As Long: dstCol = GetColumnIndex(dstWs, h)
        If dstCol = 0 Then dstCol = i + 1 ' fallback by layout
        If srcCol > 0 Then
            dstWs.Cells(dstRow, dstCol).Value = srcWs.Cells(srcRow, srcCol).Value
        Else
            ' 留空
        End If
    Next i
End Sub

' 列表目录下 BOM 文件（不包含“汇总”）
Public Function ListBOMFiles(ByVal folderPath As String) As Collection
    Dim col As New Collection
    Dim f As String
    f = Dir(folderPath & Application.PathSeparator & "*.xls*")
    Do While Len(f) > 0
        If InStr(1, f, "汇总", vbTextCompare) = 0 Then
            col.Add folderPath & Application.PathSeparator & f
        End If
        f = Dir
    Loop
    Set ListBOMFiles = col
End Function

' 在目标工作簿每个数据表创建副本并应用Toolbox替换
Public Sub ApplyToolboxReplacement_Direct(ByVal wb As Workbook, ByVal toolboxMap As Object)
    Dim ws As Worksheet
    Dim sheetName As String
    Dim replaced As Long
    Dim total As Long: total = 0
    For Each ws In wb.Worksheets
        sheetName = ws.Name
        If ws.Visible = xlSheetVisible Then
            replaced = 0
            Call SingleSheetFormatter.ApplyToolboxNameReplacement(ws, toolboxMap, replaced)
            LogInfo "工作表 [" & sheetName & "] 已替换：" & CStr(replaced) & " 条"
            total = total + replaced
        End If
    Next
    LogInfo "Toolbox名称替换完成，总计替换：" & CStr(total) & " 条"
End Sub

Public Sub ApplyToolboxReplacement_StepByStep_WPS(ByVal wb As Workbook, ByVal toolboxMap As Object)
    Dim ws As Worksheet
    Dim previewWs As Worksheet
    Dim sheetName As String
    Dim cell As Range
    For Each ws In wb.Worksheets
        sheetName = ws.Name
        If InStr(sheetName, "汇总") = 0 Then
            ' 新建副本工作表
            Set previewWs = wb.Worksheets.Add(After:=ws)
            previewWs.Name = sheetName & "_预览"
            ws.Cells.Copy previewWs.Cells
            previewWs.Activate
            ' 应用Toolbox替换
            Call SingleSheetFormatter.ApplyToolboxNameReplacement(previewWs, toolboxMap)
            ' 在副本A1写入提示
            Set cell = previewWs.Range("A1")
            cell.Value = "已应用Toolbox替换，请人工确认"
            ' 弹窗提示
            MsgBox "已生成副本：" & previewWs.Name & "，请确认替换效果。", vbInformation
        End If
    Next
End Sub

