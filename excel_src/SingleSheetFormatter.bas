Option Explicit

Public Sub ApplyToolboxNameReplacement(ByVal ws As Worksheet, ByVal mapping As Object, _
    ByRef replacedCount As Long, ByRef unmatchedCount As Long)

    Dim colE As Long, colI As Long, colJ As Long, colK As Long, colL As Long, colM As Long
    colE = Utils.GetColumnIndex(ws, Array("名称", "Name", "NAME", "品名", "零件名称", "部件名称"))
    colI = Utils.GetColumnIndex(ws, Array("SUPPLIER", "Supplier", "渠道", "供应商", "供 应 商", "SUPPLIER ", "供应渠道"))
    colJ = Utils.GetColumnIndex(ws, Array("型号", "型 号", "MODEL", "Model", "规格型号"))
    ' K 列仅使用英文列名作为匹配键，避免误用中文“名称”导致全表高亮
    colK = Utils.GetColumnIndex(ws, Array("PART NAME", "Part Name", "PARTNAME", "COMPONENT NAME", "Component Name", "COMPONENT"))
    colL = Utils.GetColumnIndex(ws, Array("规格", "Spec", "SPEC", "SPECIFICATION", "规 格", "规格参数"))
    colM = Utils.GetColumnIndex(ws, Array("标准", "Standard", "STANDARD", "标 准", "执行标准"))

    If colE = 0 Or colI = 0 Or colJ = 0 Or colK = 0 Or colL = 0 Or colM = 0 Then
        Logger.LogWarn "Header columns not found in sheet: " & ws.Name
        Dim r As Long, c As Long, rMax As Long, cMax As Long
        rMax = Application.WorksheetFunction.Min(Utils.LastUsedRow(ws), CFG_HEADER_DUMP_ROWS)
        cMax = Application.WorksheetFunction.Min(ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column, CFG_HEADER_DUMP_COLS)
        For r = 1 To rMax
            Dim rowDump As String: rowDump = ""
            For c = 1 To cMax
                Dim v As String
                v = CStr(ws.Cells(r, c).Value)
                If c = 1 Then
                    rowDump = v
                Else
                    rowDump = rowDump & " | " & v
                End If
            Next c
            Logger.LogInfo "Row" & r & ": " & rowDump
        Next r
        Exit Sub
    End If

    ' 输出识别到的列索引，便于确认是否对齐到 K=PART NAME
    Logger.LogInfo "Detected columns on [" & ws.Name & "]: E(Name)=" & colE & ", I(Supplier)=" & colI & ", J(Model)=" & colJ & ", K(PART NAME)=" & colK & ", L(Spec)=" & colL & ", M(Standard)=" & colM

    Dim lastRow As Long: lastRow = Utils.LastUsedRow(ws)
    If lastRow < 2 Then Exit Sub

    For r = 2 To lastRow
        Dim key As String
        key = Utils.NormalizeName(ws.Cells(r, colK).Value)
        If Len(key) > 0 Then
            If mapping.Exists(key) Then
                ws.Cells(r, colE).Value = mapping(key)
                ws.Cells(r, colJ).Value = ws.Cells(r, colL).Value
                ws.Cells(r, colI).Value = ws.Cells(r, colM).Value
                ' 清除之前可能存在的标黄
                ws.Rows(r).Interior.Pattern = xlNone
                replacedCount = replacedCount + 1
            Else
                ws.Rows(r).Interior.Color = RGB(255, 255, 204)
                unmatchedCount = unmatchedCount + 1
            End If
        End If
    Next r

    Logger.LogInfo "Sheet [" & ws.Name & "] replaced=" & replacedCount & ", unmatched=" & unmatchedCount
End Sub