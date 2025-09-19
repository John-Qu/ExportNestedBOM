Option Explicit

Public Function WorkbookDir(ByVal wb As Workbook) As String
    WorkbookDir = wb.Path
End Function

Public Function LastUsedRow(ByVal ws As Worksheet) As Long
    On Error Resume Next
    LastUsedRow = ws.Cells.Find(What:="*", After:=ws.Cells(1, 1), LookIn:=xlFormulas, LookAt:=xlPart, _
        SearchOrder:=xlByRows, SearchDirection:=xlPrevious).Row
    If LastUsedRow = 0 Then LastUsedRow = 1
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
    t = Replace$(t, vbCr, " ")
    t = Replace$(t, vbLf, " ")
    t = Replace$(t, "\n", " ")
    t = CollapseSpaces(t)
    NormalizeName = UCase$(t)
End Function

Public Function GetColumnIndex(ByVal ws As Worksheet, ByVal aliases As Variant) As Long
    Dim headerRow As Long: headerRow = 1
    Dim lastCol As Long
    lastCol = ws.Cells(headerRow, ws.Columns.Count).End(xlToLeft).Column
    Dim wanted As Object: Set wanted = CreateObject("Scripting.Dictionary")
    Dim i As Long
    For i = LBound(aliases) To UBound(aliases)
        wanted(NormalizeName(CStr(aliases(i)))) = True
    Next i
    For i = 1 To lastCol
        Dim h As String
        h = NormalizeName(CStr(ws.Cells(headerRow, i).Value))
        If wanted.Exists(h) Then
            GetColumnIndex = i
            Exit Function
        End If
    Next i
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