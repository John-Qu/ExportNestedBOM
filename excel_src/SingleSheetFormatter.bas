Option Explicit

Public Sub ApplyToolboxNameReplacement(ByVal ws As Worksheet, ByVal mapping As Object, _
    ByRef replacedCount As Long, ByRef unmatchedCount As Long)

    Dim colE As Long, colI As Long, colJ As Long, colK As Long, colL As Long, colM As Long, colBuy As Long
    colE = Utils.GetColumnIndex(ws, Array("名称", "Name", "NAME", "品名", "零件名称", "部件名称"))
    colI = Utils.GetColumnIndex(ws, Array("SUPPLIER", "Supplier", "渠道", "供应商", "供 应 商", "SUPPLIER ", "供应渠道"))
    colJ = Utils.GetColumnIndex(ws, Array("型号", "型 号", "MODEL", "Model", "规格型号"))
    ' K 列：用于映射匹配的“零件名称”，允许标准名与常见英文导出名，避免与 E 列“名称”混淆
    colK = Utils.GetColumnIndex(ws, Array("零件名称", "PART NAME", "Part Name", "PARTNAME", "COMPONENT NAME", "Component Name", "COMPONENT"))
    colL = Utils.GetColumnIndex(ws, Array("规格", "Spec", "SPEC", "SPECIFICATION", "规 格", "规格参数"))
    colM = Utils.GetColumnIndex(ws, Array("标准", "Standard", "STANDARD", "标 准", "执行标准"))
    ' 购列（是否外购），用于在命中映射时将其标记为“是”，后续由 IconizeBooleanFlags 转为图标
    colBuy = Utils.GetColumnIndex(ws, Array("购", "是否外购", "外购", "Purchase", "Is Purchase"))

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

    ' 输出识别到的列索引，便于确认是否对齐到 K=零件名称/PART NAME
    Logger.LogInfo "Detected columns on [" & ws.Name & "]: E(Name)=" & colE & ", I(Supplier)=" & colI & ", J(Model)=" & colJ & ", K(PartName)=" & colK & ", L(Spec)=" & colL & ", M(Standard)=" & colM & ", Buy(col)=" & colBuy

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
                ' 命中映射即视为外购件：根据用例 T1 将“是否外购”列标记为“是”，后续会被图标化
                If colBuy > 0 Then ws.Cells(r, colBuy).Value = "是"
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

' ======================== 用例 T2：列标题重命名与列顺序调整 ========================

Public Sub RenameHeadersAndReorder(ByVal ws As Worksheet)
    On Error GoTo FAIL

    Dim headerRow As Long: headerRow = DetectHeaderRow(ws)
    If headerRow = 0 Then headerRow = 1
    If headerRow <> 1 Then
        Logger.LogInfo "Header row detected at row=" & headerRow & " on sheet " & ws.Name
    End If

    ' 定义各列的别名
    Dim aPreview, aSeq, aPartNo, aCode, aName, aQty, aMaterial, aProcess
    Dim aSupplier, aModel, aPartName, aSpec, aStd, aRemark
    Dim aAsm, aBuy, aMach, aSheet

    aPreview = Array("文档预览", "预览", "Preview", "Document Preview")
    aSeq = Array("序号", "Index", "No.", "NO", "编号")
    aPartNo = Array("零件号", "PART NUMBER", "Part Number", "零件编码")
    aCode = Array("代号", "代码", "图号", "Code")
    aName = Array("名称", "Name", "品名", "部件名称")
    aQty = Array("数量", "Qty", "QTY", "件数", "数量（个）")
    aMaterial = Array("材料", "材 料", "材     料", "Material", "MATERIAL")
    aProcess = Array("处理", "表面处理", "Finish", "Treatment", "处理方式")

    aSupplier = Array("SUPPLIER", "Supplier", "渠道", "供应商", "供应渠道")
    aModel = Array("型号", "MODEL", "Model", "规格型号")
    aPartName = Array("零件名称", "PART NAME", "Part Name", "PARTNAME", "COMPONENT NAME", "COMPONENT")
    aSpec = Array("规格", "SPEC", "Spec", "SPECIFICATION", "规格参数")
    aStd = Array("标准", "Standard", "STANDARD", "执行标准")
    aRemark = Array("备注", "Remark", "REMARK", "说明")

    aAsm = Array("是否组装", "组装", "Assembly", "Is Assembly", "组")
    aBuy = Array("是否外购", "外购", "Purchase", "Is Purchase", "购")
    aMach = Array("是否机加", "机加", "Machining", "Is Machining", "加")
    aSheet = Array("是否钣金", "钣金", "Sheet Metal", "Is Sheet Metal", "钣")

    ' 找列号
    Dim cPreview As Long, cSeq As Long, cPartNo As Long, cCode As Long, cName As Long, cQty As Long
    Dim cMaterial As Long, cProcess As Long, cSupplier As Long, cModel As Long, cPartName As Long
    Dim cSpec As Long, cStd As Long, cRemark As Long, cAsm As Long, cBuy As Long, cMach As Long, cSheet As Long

    cPreview = FindHeaderColInRow(ws, headerRow, aPreview)
    cSeq = FindHeaderColInRow(ws, headerRow, aSeq)
    cPartNo = FindHeaderColInRow(ws, headerRow, aPartNo)
    cCode = FindHeaderColInRow(ws, headerRow, aCode)
    cName = FindHeaderColInRow(ws, headerRow, aName)
    cQty = FindHeaderColInRow(ws, headerRow, aQty)
    cMaterial = FindHeaderColInRow(ws, headerRow, aMaterial)
    cProcess = FindHeaderColInRow(ws, headerRow, aProcess)
    cSupplier = FindHeaderColInRow(ws, headerRow, aSupplier)
    cModel = FindHeaderColInRow(ws, headerRow, aModel)
    cPartName = FindHeaderColInRow(ws, headerRow, aPartName)
    cSpec = FindHeaderColInRow(ws, headerRow, aSpec)
    cStd = FindHeaderColInRow(ws, headerRow, aStd)
    cRemark = FindHeaderColInRow(ws, headerRow, aRemark)
    cAsm = FindHeaderColInRow(ws, headerRow, aAsm)
    cBuy = FindHeaderColInRow(ws, headerRow, aBuy)
    cMach = FindHeaderColInRow(ws, headerRow, aMach)
    cSheet = FindHeaderColInRow(ws, headerRow, aSheet)

    ' 修复：若“数量”列标题为空（空格/换行导致），按模板约定它位于“名称”右侧一列，自动补齐为“数量”
    If cQty = 0 And cName > 0 Then
        Dim candidateQtyCol As Long: candidateQtyCol = cName + 1
        Dim maxCol As Long: maxCol = ws.Cells(headerRow, ws.Columns.Count).End(xlToLeft).Column
        If candidateQtyCol <= maxCol Then
            Dim rawHeader As String: rawHeader = CStr(ws.Cells(headerRow, candidateQtyCol).Value)
            If Len(Utils.NormalizeName(rawHeader)) = 0 Then
                ws.Cells(headerRow, candidateQtyCol).Value = "数量"
                cQty = candidateQtyCol
                Logger.LogInfo "Filled blank Qty header at row=" & headerRow & ", col=" & candidateQtyCol & " on sheet " & ws.Name
            End If
        End If
    End If

    ' 标题重命名（统一标准名）
    If cPartNo > 0 Then ws.Cells(headerRow, cPartNo).Value = "零件号"
    If cPreview > 0 Then ws.Cells(headerRow, cPreview).Value = "文档预览"
    If cSeq > 0 Then ws.Cells(headerRow, cSeq).Value = "序号"
    If cCode > 0 Then ws.Cells(headerRow, cCode).Value = "代号"
    If cName > 0 Then ws.Cells(headerRow, cName).Value = "名称"
    If cQty > 0 Then ws.Cells(headerRow, cQty).Value = "数量"
    If cMaterial > 0 Then ws.Cells(headerRow, cMaterial).Value = "材料" ' 规范化去空格
    If cProcess > 0 Then ws.Cells(headerRow, cProcess).Value = "处理"
    If cSupplier > 0 Then ws.Cells(headerRow, cSupplier).Value = "渠道" ' SUPPLIER -> 渠道
    If cModel > 0 Then ws.Cells(headerRow, cModel).Value = "型号"
    If cPartName > 0 Then ws.Cells(headerRow, cPartName).Value = "零件名称"
    If cSpec > 0 Then ws.Cells(headerRow, cSpec).Value = "规格"
    If cStd > 0 Then ws.Cells(headerRow, cStd).Value = "标准"
    If cRemark > 0 Then ws.Cells(headerRow, cRemark).Value = "备注"
    If cAsm > 0 Then ws.Cells(headerRow, cAsm).Value = "组"
    If cBuy > 0 Then ws.Cells(headerRow, cBuy).Value = "购"
    If cMach > 0 Then ws.Cells(headerRow, cMach).Value = "加"
    If cSheet > 0 Then ws.Cells(headerRow, cSheet).Value = "钣"

    Logger.LogInfo "Headers normalized on [" & ws.Name & "] at row=" & headerRow

    ' 若缺少“备注”列但存在“零件名称”列，则在其前插入空白“备注”列，保证列序
    If cRemark = 0 And cPartName > 0 Then
        ws.Columns(cPartName).Insert Shift:=xlToRight
        ws.Cells(headerRow, cPartName).Value = "备注"
        cRemark = cPartName
        cPartName = cPartName + 1
        Logger.LogInfo "Inserted missing [备注] column before [零件名称] on sheet " & ws.Name
    End If

    ' 列顺序调整到最终序
    Dim desired()
    desired = Array("零件号", "文档预览", "序号", "代号", "名称", "数量", _
                    "材料", "处理", "渠道", "型号", _
                    "组", "购", "加", "钣", _
                    "备注", "零件名称", "规格", "标准")

    Dim pos As Long
    For pos = LBound(desired) To UBound(desired)
        Dim targetName As String: targetName = CStr(desired(pos))
        Dim wantCol As Long: wantCol = pos + 1 ' 数组从0开始
        Dim curCol As Long
        curCol = FindHeaderColInRow(ws, headerRow, Array(targetName))
        If curCol > 0 And curCol <> wantCol Then
            ' 剪切并插入到目标位置
            ws.Columns(curCol).Cut
            ws.Columns(wantCol).Insert Shift:=xlToRight
        End If
    Next pos

    Logger.LogInfo "Reordered columns on [" & ws.Name & "] to final sequence A:R"
    Exit Sub
FAIL:
    Logger.LogError "RenameHeadersAndReorder failed on sheet " & ws.Name & ": " & Err.Description
End Sub

' ======================== 用例 T3：布尔显示图标化（组/购/加/钣 -> ●/X） ========================

Public Sub IconizeBooleanFlags(ByVal ws As Worksheet)
    On Error GoTo FAIL
    Dim headerRow As Long: headerRow = DetectHeaderRow(ws)
    If headerRow = 0 Then headerRow = 1

    ' 目标列别名（兼容重命名前/后）
    Dim aAsm, aBuy, aMach, aSheet
    aAsm = Array("组", "是否组装", "组装", "Assembly", "Is Assembly")
    aBuy = Array("购", "是否外购", "外购", "Purchase", "Is Purchase")
    aMach = Array("加", "是否机加", "机加", "Machining", "Is Machining")
    aSheet = Array("钣", "是否钣金", "钣金", "Sheet Metal", "Is Sheet Metal")

    Dim cAsm As Long, cBuy As Long, cMach As Long, cSheet As Long
    cAsm = FindHeaderColInRow(ws, headerRow, aAsm)
    cBuy = FindHeaderColInRow(ws, headerRow, aBuy)
    cMach = FindHeaderColInRow(ws, headerRow, aMach)
    cSheet = FindHeaderColInRow(ws, headerRow, aSheet)

    If cAsm = 0 And cBuy = 0 And cMach = 0 And cSheet = 0 Then
        Logger.LogWarn "IconizeBooleanFlags: boolean columns not found on sheet " & ws.Name
        Exit Sub
    End If

    ' 构建真值集合字典（规格化后对比）
    Dim trueDict As Object: Set trueDict = CreateObject("Scripting.Dictionary")
    Dim tvs As Variant: tvs = CFG_TRUE_SET()
    Dim i As Long
    For i = LBound(tvs) To UBound(tvs)
        trueDict(Utils.NormalizeName(CStr(tvs(i)))) = True
    Next i

    Dim lastRow As Long: lastRow = Utils.LastUsedRow(ws)
    Dim r As Long, cnt As Long: cnt = 0
    For r = headerRow + 1 To lastRow
        ' 针对每个目标列进行图标化
        If cAsm > 0 Then cnt = cnt + IconizeOneCell(ws, r, cAsm, trueDict)
        If cBuy > 0 Then cnt = cnt + IconizeOneCell(ws, r, cBuy, trueDict)
        If cMach > 0 Then cnt = cnt + IconizeOneCell(ws, r, cMach, trueDict)
        If cSheet > 0 Then cnt = cnt + IconizeOneCell(ws, r, cSheet, trueDict)
    Next r

    ' 在替换图标后，删除末尾无“文档预览”的行
    DeleteTrailingRowsWithoutPreview ws

    Logger.LogInfo "IconizeBooleanFlags done on [" & ws.Name & "]: changed cells=" & cnt
    Exit Sub
FAIL:
    Logger.LogError "IconizeBooleanFlags failed on sheet " & ws.Name & ": " & Err.Description
End Sub

Private Function IconizeOneCell(ByVal ws As Worksheet, ByVal rowIndex As Long, ByVal colIndex As Long, ByVal trueDict As Object) As Long
    Dim v As String
    v = CStr(ws.Cells(rowIndex, colIndex).Value)
    Dim norm As String: norm = Utils.NormalizeName(v)

    Dim isTrue As Boolean
    If Len(norm) = 0 Then
        isTrue = False
    ElseIf StrComp(v, ICON_TRUE, vbBinaryCompare) = 0 Then
        isTrue = True
    ElseIf StrComp(v, ICON_FALSE, vbBinaryCompare) = 0 Then
        isTrue = False
    Else
        isTrue = trueDict.Exists(norm)
    End If

    Dim newVal As String
    newVal = IIf(isTrue, ICON_TRUE, ICON_FALSE)
    Debug.Print "写入符号：" & newVal

    If StrComp(CStr(ws.Cells(rowIndex, colIndex).Value), newVal, vbBinaryCompare) <> 0 Then
        ws.Cells(rowIndex, colIndex).Value = newVal
        IconizeOneCell = 1
    Else
        IconizeOneCell = 0
    End If
End Function

Private Function DetectHeaderRow(ByVal ws As Worksheet) As Long
    Dim lastCol As Long
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column

    Dim maxScan As Long: maxScan = IIf(CFG_HEADER_SCAN_MAX_ROWS > 0, CFG_HEADER_SCAN_MAX_ROWS, 1)
    Dim row As Long, i As Long

    Dim vocab As Object: Set vocab = CreateObject("Scripting.Dictionary")
    Dim aliases
    aliases = Array( _
        "文档预览", "预览", "Preview", "Document Preview", _
        "序号", "Index", "No.", "NO", "编号", _
        "零件号", "PART NUMBER", "Part Number", "零件编码", _
        "代号", "代码", "图号", "Code", _
        "名称", "Name", "品名", "部件名称", _
        "数量", "Qty", "QTY", "件数", _
        "材料", "材 料", "材     料", "Material", "MATERIAL", _
        "处理", "表面处理", "Finish", "Treatment", "处理方式", _
        "SUPPLIER", "Supplier", "渠道", "供应商", "供应渠道", _
        "型号", "MODEL", "Model", "规格型号", _
        "零件名称", "PART NAME", "Part Name", "PARTNAME", "COMPONENT NAME", "COMPONENT", _
        "规格", "SPEC", "Spec", "SPECIFICATION", "规格参数", _
        "标准", "Standard", "STANDARD", "执行标准", _
        "备注", "Remark", "REMARK", "说明", _
        "是否组装", "组装", "Assembly", "Is Assembly", _
        "是否外购", "外购", "Purchase", "Is Purchase", _
        "是否机加", "机加", "Machining", "Is Machining", _
        "是否钣金", "钣金", "Sheet Metal", "Is Sheet Metal")

    For i = LBound(aliases) To UBound(aliases)
        vocab(Utils.NormalizeName(CStr(aliases(i)))) = True
    Next i

    Dim bestRow As Long: bestRow = 1
    Dim bestCount As Long: bestCount = -1

    For row = 1 To maxScan
        Dim cnt As Long: cnt = 0
        For i = 1 To lastCol
            Dim h As String
            h = Utils.NormalizeName(CStr(ws.Cells(row, i).Value))
            If Len(h) > 0 Then
                If vocab.Exists(h) Then cnt = cnt + 1
            End If
        Next i
        If cnt > bestCount Then
            bestCount = cnt
            bestRow = row
        End If
    Next row

    DetectHeaderRow = bestRow
End Function

Private Function FindHeaderColInRow(ByVal ws As Worksheet, ByVal headerRow As Long, ByVal aliases As Variant) As Long
    Dim lastCol As Long: lastCol = ws.Cells(headerRow, ws.Columns.Count).End(xlToLeft).Column
    Dim dict As Object: Set dict = CreateObject("Scripting.Dictionary")
    Dim i As Long
    For i = LBound(aliases) To UBound(aliases)
        dict(Utils.NormalizeName(CStr(aliases(i)))) = True
    Next i
    For i = 1 To lastCol
        Dim h As String
        h = Utils.NormalizeName(CStr(ws.Cells(headerRow, i).Value))
        If Len(h) > 0 Then
            If dict.Exists(h) Then
                FindHeaderColInRow = i
                Exit Function
            End If
        End If
    Next i
    FindHeaderColInRow = 0
End Function

' 统一格式化单表（S2/S3/S6/S7）：重命名与列序、布尔图标化、字体与对齐、打印设置
Public Sub FormatSingleBOMSheet(ByVal ws As Worksheet)
    On Error GoTo FAIL
    RenameHeadersAndReorder ws
    IconizeBooleanFlags ws
    ApplyFontAndAlignment ws
    ApplyPrintSetup ws
    Logger.LogInfo "FormatSingleBOMSheet done on [" & ws.Name & "]"
    Exit Sub
FAIL:
    Logger.LogError "FormatSingleBOMSheet failed on sheet " & ws.Name & ": " & Err.Description
End Sub

' S6 字体与对齐、列宽
Public Sub ApplyFontAndAlignment(ByVal ws As Worksheet)
    Dim lastRow As Long: lastRow = Utils.LastUsedRow(ws)
    Dim lastCol As Long: lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    Dim headerRow As Long: headerRow = 1
    Dim fontName As String
    Dim i As Long, r As Long, c As Long
    fontName = "汉仪长仿宋体"
    On Error Resume Next
    ws.Cells.Font.Name = fontName
    If ws.Cells.Font.Name <> fontName Then
        fontName = "宋体"
        ws.Cells.Font.Name = fontName
        If ws.Cells.Font.Name <> fontName Then
            fontName = "微软雅黑"
            ws.Cells.Font.Name = fontName
        End If
    End If
    ws.Cells.Font.Size = 14
    ws.Cells.Font.Bold = False
    ws.Rows(headerRow).Font.Bold = True
    ws.Rows(headerRow).HorizontalAlignment = xlCenter
    ws.Rows(headerRow).VerticalAlignment = xlCenter
    Dim colWidths As Variant '列宽
    colWidths = Array(22, 15, 4, 15, 8, 4, 8, 8, 10, 12, 2.5, 2.5, 2.5, 2.5, 27, 12, 12, 12)
    For i = 1 To Application.WorksheetFunction.Min(UBound(colWidths) + 1, lastCol)
        ws.Columns(i).ColumnWidth = colWidths(i - 1)
    Next i
    For r = headerRow + 1 To lastRow
        ws.Rows(r).Font.Bold = False
        ws.Rows(r).VerticalAlignment = xlCenter
        For c = 1 To lastCol
            Select Case c
                Case 4, 10, 8, 15 ' D 代号、J 型号、H 处理、O 备注
                    ws.Cells(r, c).HorizontalAlignment = xlLeft
                Case 6 ' F 数量
                    ws.Cells(r, c).HorizontalAlignment = xlRight
                Case Else
                    ws.Cells(r, c).HorizontalAlignment = xlCenter
            End Select
        Next c
    Next r
End Sub

' S7 打印设置与页眉页脚
Public Sub ApplyPrintSetup(ByVal ws As Worksheet)
    Dim lastRow As Long: lastRow = Utils.LastUsedRow(ws)
    With ws.PageSetup
        .PrintArea = ws.Range(ws.Cells(1, 2), ws.Cells(lastRow, 15)).Address  ' B:O
        .PrintTitleRows = "$1:$1"
        .Orientation = xlLandscape
        .PaperSize = xlPaperA4
        .Zoom = 100
        .LeftHeader = Utils.GetLeafFolderName(Utils.WorkbookDir(ws.Parent))
        .CenterHeader = ws.Parent.Name
        .RightHeader = Format(FileDateTime(ws.Parent.FullName), "yyyy-mm-dd")
        .CenterFooter = "第 &P 页，共 &N 页"
        .LeftMargin = Application.CentimetersToPoints(1.5)
        .RightMargin = Application.CentimetersToPoints(1.5)
        .TopMargin = Application.CentimetersToPoints(1.5)
        .BottomMargin = Application.CentimetersToPoints(1.5)
        .HeaderMargin = Application.CentimetersToPoints(0.8)
        .FooterMargin = Application.CentimetersToPoints(0.8)
    End With
    ws.Cells.Borders.LineStyle = xlNone
End Sub

' ===== 新增：删除末尾无预览行 =====
Private Sub DeleteTrailingRowsWithoutPreview(ByVal ws As Worksheet)
    On Error GoTo FIN
    Dim headerRow As Long: headerRow = DetectHeaderRow(ws)
    If headerRow = 0 Then headerRow = 1

    Dim aPreview As Variant
    aPreview = Array("文档预览", "预览", "Preview", "Document Preview")
    Dim cPreview As Long: cPreview = FindHeaderColInRow(ws, headerRow, aPreview)
    If cPreview = 0 Then Exit Sub

    Dim lastRow As Long: lastRow = Utils.LastUsedRow(ws)
    If lastRow <= headerRow Then Exit Sub

    Dim removed As Long: removed = 0
    Dim r As Long
    For r = lastRow To headerRow + 1 Step -1
        If Not CellHasPreview(ws, r, cPreview) Then
            ws.Rows(r).Delete
            removed = removed + 1
        Else
            Exit For ' 仅删除末尾连续的无预览行
        End If
    Next r

    If removed > 0 Then
        Logger.LogInfo "Deleted trailing rows without preview on [" & ws.Name & "]: " & removed
    End If
FIN:
End Sub

Private Function CellHasPreview(ByVal ws As Worksheet, ByVal rowIndex As Long, ByVal colIndex As Long) As Boolean
    On Error GoTo SAFE
    Dim cell As Range: Set cell = ws.Cells(rowIndex, colIndex)
    Dim txt As String: txt = Trim$(CStr(cell.Value))
    If Len(txt) > 0 Then
        CellHasPreview = True
        Exit Function
    End If

    Dim left# , top#, right#, bottom#
    left = cell.Left: top = cell.Top: right = left + cell.Width: bottom = top + cell.Height

    Dim shp As Shape
    For Each shp In ws.Shapes
        If shp.Visible Then
            Dim cx As Double, cy As Double
            cx = shp.Left + shp.Width / 2
            cy = shp.Top + shp.Height / 2
            If cx > left And cx < right And cy > top And cy < bottom Then
                CellHasPreview = True
                Exit Function
            End If
        End If
    Next shp

    Dim ole As OLEObject
    For Each ole In ws.OLEObjects
        If ole.Visible Then
            Dim ocx As Double, ocy As Double
            ocx = ole.Left + ole.Width / 2
            ocy = ole.Top + ole.Height / 2
            If ocx > left And ocx < right And ocy > top And ocy < bottom Then
                CellHasPreview = True
                Exit Function
            End If
        End If
    Next ole

SAFE:
    ' 若未命中，视为无预览
    If Not CellHasPreview Then CellHasPreview = False
End Function