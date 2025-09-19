Option Explicit

Public Sub ApplyToolboxNameReplacement(ByVal ws As Worksheet, ByVal mapping As Object, _
    ByRef replacedCount As Long, ByRef unmatchedCount As Long)

    Dim colE As Long, colI As Long, colJ As Long, colK As Long, colL As Long, colM As Long
    colE = Utils.GetColumnIndex(ws, Array("名称", "NAME"))
    colI = Utils.GetColumnIndex(ws, Array("SUPPLIER", "渠道", "SUPPLIER "))
    colJ = Utils.GetColumnIndex(ws, Array("型号", "MODEL"))
    colK = Utils.GetColumnIndex(ws, Array("零件名称", "PART NAME", "COMPONENT NAME", "零件名"))
    colL = Utils.GetColumnIndex(ws, Array("规格", "SPEC"))
    colM = Utils.GetColumnIndex(ws, Array("标准", "STANDARD"))

    If colE = 0 Or colI = 0 Or colJ = 0 Or colK = 0 Or colL = 0 Or colM = 0 Then
        Logger.LogWarn "Header columns not found in sheet: " & ws.Name
        Exit Sub
    End If

    Dim lastRow As Long: lastRow = Utils.LastUsedRow(ws)
    If lastRow < 2 Then Exit Sub

    Dim r As Long
    For r = 2 To lastRow
        Dim key As String
        key = Utils.NormalizeName(ws.Cells(r, colK).Value)
        If Len(key) > 0 Then
            If mapping.Exists(key) Then
                ws.Cells(r, colE).Value = mapping(key)
                ws.Cells(r, colJ).Value = ws.Cells(r, colL).Value
                ws.Cells(r, colI).Value = ws.Cells(r, colM).Value
                replacedCount = replacedCount + 1
            Else
                ws.Rows(r).Interior.Color = RGB(255, 255, 204)
                unmatchedCount = unmatchedCount + 1
            End If
        End If
    Next r

    Logger.LogInfo "Sheet [" & ws.Name & "] replaced=" & replacedCount & ", unmatched=" & unmatchedCount
End Sub