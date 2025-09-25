Option Explicit

Public Function WorkbookDir(ByVal wb As Workbook) As String
    WorkbookDir = wb.Path
End Function

Public Function LastUsedRow(ByVal ws As Worksheet) As Long
    ' 更稳健的“最后行”探测：综合 公式/常量/UsedRange 三种方式，取最大值
    Dim lrF As Long, lrV As Long, lrU As Long
    On Error Resume Next
    lrF = ws.Cells.Find(What:="*", After:=ws.Cells(1, 1), LookIn:=xlFormulas, LookAt:=xlPart, _
                        SearchOrder:=xlByRows, SearchDirection:=xlPrevious).Row
    lrV = ws.Cells.Find(What:="*", After:=ws.Cells(1, 1), LookIn:=xlValues, LookAt:=xlPart, _
                        SearchOrder:=xlByRows, SearchDirection:=xlPrevious).Row
    lrU = ws.UsedRange.Row + ws.UsedRange.Rows.Count - 1
    On Error GoTo 0
    If lrF < 1 Then lrF = 1
    If lrV < 1 Then lrV = 1
    If lrU < 1 Then lrU = 1
    LastUsedRow = Application.WorksheetFunction.Max(lrF, lrV, lrU)
End Function

Public Function CollapseSpaces(ByVal s As String) As String
    Dim t As String
    t = Trim$(s)
    Do While InStr(t, "  ") > 0
        t = Replace$(t, "  ", " ")
    Loop
    CollapseSpaces = t
End Function

Public Function NormalizeName(ByVal s As String) As String
    Dim t As String
    t = CStr(s)
    ' 标准化行内与跨行空白
    t = Replace$(t, vbCr, " ")
    t = Replace$(t, vbLf, " ")
    t = Replace$(t, "\n", " ")
    ' 将全角空格与制表符等统一为半角空格
    t = Replace$(t, ChrW(12288), " ") ' 全角空格
    t = Replace$(t, vbTab, " ")
    ' 去除不可见零宽字符（常见 PDF/导出残留）
    t = Replace$(t, ChrW(&H200B), "") ' Zero Width Space
    t = Replace$(t, ChrW(&H200C), "") ' Zero Width Non-Joiner
    t = Replace$(t, ChrW(&H200D), "") ' Zero Width Joiner
    t = Replace$(t, ChrW(&HFEFF), "") ' BOM
    ' 归一化多空格
    t = CollapseSpaces(t)

    ' 额外鲁棒性：如果字符串中不包含任何 ASCII 字母或数字，视为纯中文/符号标题，移除所有空格
    ' 这可修复 A3 模板中列头内断行（如“是否外\n购”、“是否钣\n金”）导致匹配失败的问题
    Dim i As Long, ch As String, hasAsciiAlnum As Boolean
    hasAsciiAlnum = False
    For i = 1 To Len(t)
        ch = Mid$(t, i, 1)
        If (ch Like "[A-Za-z0-9]") Then
            hasAsciiAlnum = True
            Exit For
        End If
    Next i
    If Not hasAsciiAlnum Then
        t = Replace$(t, " ", "")
    End If

    NormalizeName = UCase$(t)
End Function

Public Function GetColumnIndex(ByVal ws As Worksheet, ByVal aliases As Variant) As Long
    Dim wanted As Object: Set wanted = CreateObject("Scripting.Dictionary")
    Dim i As Long, headerRow As Long
    For i = LBound(aliases) To UBound(aliases)
        wanted(NormalizeName(CStr(aliases(i)))) = True
    Next i

    Dim maxScan As Long: maxScan = IIf(CFG_HEADER_SCAN_MAX_ROWS > 0, CFG_HEADER_SCAN_MAX_ROWS, 1)
    For headerRow = 1 To maxScan
        ' 修正：针对每个候选表头行单独计算该行的实际最右列
        Dim lastCol As Long
        lastCol = ws.Cells(headerRow, ws.Columns.Count).End(xlToLeft).Column
        If lastCol < 1 Then lastCol = 1
        For i = 1 To lastCol
            Dim h As String
            h = NormalizeName(CStr(ws.Cells(headerRow, i).Value))
            If Len(h) > 0 Then
                If wanted.Exists(h) Then
                    GetColumnIndex = i
                    If headerRow <> 1 Then
                        Logger.LogInfo "Header '" & h & "' found at row=" & headerRow & ", col=" & i & " on sheet " & ws.Name
                    End If
                    Exit Function
                End If
            End If
        Next i
    Next headerRow

    GetColumnIndex = 0
End Function

Public Function LoadToolboxMapping(ByVal baseDir As String) As Object
    Dim map As Object: Set map = CreateObject("Scripting.Dictionary")
    Dim wb As Workbook
    Set wb = FindOpenWorkbookByName(CFG_MAPPING_WORKBOOK_NAME)
    Dim opened As Boolean: opened = False
    If wb Is Nothing Then
        Dim path As String
        path = baseDir & "\" & CFG_MAPPING_WORKBOOK_NAME
        On Error Resume Next
        Set wb = Application.Workbooks.Open(FileName:=path, ReadOnly:=True)
        opened = Not wb Is Nothing
        On Error GoTo 0
    End If
    If wb Is Nothing Then
        Logger.LogWarn "Mapping workbook not found: " & CFG_MAPPING_WORKBOOK_NAME & " under " & baseDir
        Set LoadToolboxMapping = map
        Exit Function
    End If
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = wb.Worksheets(CFG_MAPPING_SHEET_NAME)
    On Error GoTo 0
    If ws Is Nothing Then
        Logger.LogWarn "Mapping sheet not found: " & CFG_MAPPING_SHEET_NAME
        If opened Then wb.Close SaveChanges:=False
        Set LoadToolboxMapping = map
        Exit Function
    End If
    Dim lastRow As Long: lastRow = LastUsedRow(ws)
    Dim r As Long
    For r = 2 To lastRow
        Dim en As String, cn As String
        en = NormalizeName(ws.Cells(r, 1).Value)
        cn = CStr(ws.Cells(r, 2).Value)
        If Len(en) > 0 Then
            If Not map.Exists(en) Then map.Add en, cn Else map(en) = cn
        End If
    Next r
    Logger.LogInfo "Mapping loaded: entries=" & map.Count & " from " & CFG_MAPPING_WORKBOOK_NAME & "!" & CFG_MAPPING_SHEET_NAME
    If opened Then wb.Close SaveChanges:=False
    Set LoadToolboxMapping = map
End Function

Public Function FindOpenWorkbookByName(ByVal nameOnly As String) As Workbook
    Dim wb As Workbook
    For Each wb In Application.Workbooks
        If StrComp(UCase$(wb.Name), UCase$(nameOnly), vbTextCompare) = 0 Then
            Set FindOpenWorkbookByName = wb
            Exit Function
        End If
    Next wb
    Set FindOpenWorkbookByName = Nothing
End Function

' Return the workbook to process (target), preferring ActiveWorkbook if it is not the mapping/macro workbook.
Public Function ResolveTargetWorkbook() As Workbook
    Dim macroName As String: macroName = CFG_MAPPING_WORKBOOK_NAME
    Dim wb As Workbook

    ' Prefer the active workbook if it is not the macro/mapping workbook
    On Error Resume Next
    If Not ActiveWorkbook Is Nothing Then
        If StrComp(UCase$(ActiveWorkbook.Name), UCase$(macroName), vbTextCompare) <> 0 Then
            Set ResolveTargetWorkbook = ActiveWorkbook
            Exit Function
        End If
    End If
    On Error GoTo 0

    ' Otherwise, pick the first open workbook that is not the macro workbook
    For Each wb In Application.Workbooks
        If StrComp(UCase$(wb.Name), UCase$(macroName), vbTextCompare) <> 0 Then
            Set ResolveTargetWorkbook = wb
            Exit Function
        End If
    Next wb

    Set ResolveTargetWorkbook = Nothing
End Function


' 获取路径最后一级目录名（不含上级路径分隔符）
Public Function GetLeafFolderName(ByVal folderPath As String) As String
    Dim p As Long: p = InStrRev(folderPath, Application.PathSeparator)
    If p > 0 Then
        GetLeafFolderName = Mid$(folderPath, p + 1)
    Else
        GetLeafFolderName = folderPath
    End If
End Function