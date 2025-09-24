Option Explicit

' 构建“总 BOM 清单”：
' - 以当前目标工作簿（通常为“*_汇总.xls”）为驱动，定位“代号/总数量/分解链”列（同一表头行）
' - 扫描同目录所有不含“汇总”的 *.xls* 子装配/主装配清单文件
' - 在这些 BOM 表中按“零件号”精确匹配，复制描述性字段（首次出现优先）
' - 在输出表中：F=数量 使用汇总“总数量”；在 O=备注 右侧新增一列“计算说明”承载“分解链”
' - 对输出表调用 IconizeBooleanFlags/ApplyFontAndAlignment/ApplyPrintSetup 统一格式
Public Sub BuildTotalBOMFromSummary()
    On Error GoTo FAIL
    Dim wb As Workbook: Set wb = Utils.ResolveTargetWorkbook()
    If wb Is Nothing Then
        MsgBox "未找到可处理的目标工作簿（请打开 *_汇总.xls 文件）", vbExclamation
        Exit Sub
    End If

    Dim baseDir As String: baseDir = Utils.WorkbookDir(wb)
    Logger.LogInfo "T6: Summary workbook=" & wb.Name & ", Dir=" & baseDir

    Dim wsSum As Worksheet
    Dim cKey As Long, cQty As Long, cChain As Long, headerRow As Long
    Set wsSum = FindSummarySheetAndCols(wb, cKey, cQty, cChain, headerRow)
    If wsSum Is Nothing Then
        Logger.LogError "T6: 未在任何工作表中识别到 ‘代号/总数量/分解链’ 列"
        MsgBox "未在该工作簿中发现可用的 汇总 表（需要包含列：代号/总数量/分解链）", vbCritical
        Exit Sub
    End If
    Logger.LogInfo "T6: Summary sheet=" & wsSum.Name & _
                  ", headerRow=" & headerRow & ", keyCol=" & cKey & ", qtyCol=" & cQty & ", chainCol=" & cChain

    ' 预备输出工作表
    Dim wsOut As Worksheet
    Set wsOut = PrepareOutputSheet(wb)

    ' 构建同目录 BOM 工作簿列表（排除 汇总 与 映射/宏 工作簿）
    Dim bomBooks As Collection: Set bomBooks = New Collection
    Dim openedBooks As Collection: Set openedBooks = New Collection
    GatherBOMWorkbooks baseDir, wb, bomBooks, openedBooks
    Logger.LogInfo "T6: BOM workbook count=" & bomBooks.Count

    ' 目标字段别名集合（用于从来源行抽取）
    Dim a零件号, a序号, a代号, a名称, a数量, a材料, a处理, a渠道, aSUP, a型号
    Dim a组, a购, a加, a钣, a备注, a零件名称, a规格, a标准, a文档预览
    a零件号 = Array("零件号", "编码", "编号", "Part Number")
    a序号 = Array("序号", "项目号", "行号", "Index")
    a代号 = Array("代号", "编号", "Code")
    a名称 = Array("名称", "件名", "Name")
    a数量 = Array("数量", "数目", "Qty", "Quantity")
    a材料 = Array("材料", "材质", "Material")
    a处理 = Array("处理", "表面处理", "Finish", "Processing")
    a渠道 = Array("渠道", "供应商", "Supplier", "供方", "SUPPLIER")
    aSUP = Array("SUPPLIER", "Supplier", "渠道", "供方")
    a型号 = Array("型号", "Model", "Type")
    a组 = Array("组", "是否组装", "组装", "Assembly", "Is Assembly")
    a购 = Array("购", "是否外购", "外购", "Purchase", "Is Purchase")
    a加 = Array("加", "是否机加", "机加", "Machining", "Is Machining")
    a钣 = Array("钣", "是否钣金", "钣金", "Sheet Metal", "Is Sheet Metal")
    a备注 = Array("备注", "说明", "Remark", "备注信息")
    a零件名称 = Array("零件名称", "Part Name")
    a规格 = Array("规格", "Spec", "Specification")
    a标准 = Array("标准", "Standard")
    a文档预览 = Array("文档预览", "预览", "Preview")

    ' 写表头（含新增列“计算说明”放在 备注 右侧）
    WriteOutputHeader wsOut

    ' 开始逐行汇总：从实际表头行的下一行开始
    Dim rSum As Long, lastSumRow As Long
    lastSumRow = Utils.LastUsedRow(wsSum)
    Logger.LogInfo "T6: Summary lastSumRow=" & lastSumRow

    Dim outRow As Long: outRow = 2
    Dim key As String, qty As Variant, chain As String
    Dim oSrcWS As Worksheet
    Dim oSrcRow As Long
    Dim oPrevCol As Long

    For rSum = headerRow + 1 To lastSumRow
        On Error GoTo ROW_FAIL
        key = CStr(wsSum.Cells(rSum, cKey).Value)
        key = Trim$(key)
        If Len(key) = 0 Then GoTo CONTINUE_LOOP
        qty = wsSum.Cells(rSum, cQty).Value
        If Not IsNumeric(qty) Then qty = 0
        chain = CStr(wsSum.Cells(rSum, cChain).Value)

        Dim found As Boolean: found = False
        Dim srcValues(1 To 19) As Variant ' 对应最终 19 列（不含“计算说明”）的抽取值
        Dim wbi As Variant
        For Each wbi In bomBooks
            Dim wbBOM As Workbook: Set wbBOM = wbi
            If ExtractRowByKey(wbBOM, key, a零件号, _
                               a文档预览, a序号, a代号, a名称, a数量, a材料, a处理, a渠道, aSUP, a型号, _
                               a组, a购, a加, a钣, a备注, a零件名称, a规格, a标准, srcValues, oSrcWS, oSrcRow, oPrevCol) Then
                found = True
                Exit For
            End If
        Next wbi

        ' 输出一行
        ' 目标列顺序（19列）：
        ' A 零件号, B 文档预览, C 序号, D 代号, E 名称, F 数量, G 材料, H 处理, I 渠道, J 型号,
        ' K 组, L 购, M 加, N 钣, O 备注, [P 计算说明], Q 零件名称, R 规格, S 标准
        wsOut.Cells(outRow, 1).Value = key
        ' 默认先清空各列
        Dim c As Long
        For c = 2 To 19
            wsOut.Cells(outRow, c).ClearContents
        Next c

        If found Then
            ' 将抽取值填入（注意 srcValues 对应 A..S 但不含“计算说明”）
            ' A 已写，B..O、Q..S 从 srcValues 映射
            wsOut.Cells(outRow, 2).Value = srcValues(2)   ' 文档预览
            wsOut.Cells(outRow, 3).Value = srcValues(3)   ' 序号/项目号
            wsOut.Cells(outRow, 4).Value = srcValues(4)   ' 代号
            wsOut.Cells(outRow, 5).Value = srcValues(5)   ' 名称
            ' F 数量 使用汇总 qty 覆盖
            wsOut.Cells(outRow, 6).Value = CDbl(qty)
            wsOut.Cells(outRow, 7).Value = srcValues(7)   ' 材料
            wsOut.Cells(outRow, 8).Value = srcValues(8)   ' 处理
            wsOut.Cells(outRow, 9).Value = IIf(Len(CStr(srcValues(9))) > 0, srcValues(9), srcValues(10)) ' 渠道/SUPPLIER 兜底
            wsOut.Cells(outRow, 10).Value = srcValues(11) ' 型号
            wsOut.Cells(outRow, 11).Value = srcValues(12) ' 组
            wsOut.Cells(outRow, 12).Value = srcValues(13) ' 购
            wsOut.Cells(outRow, 13).Value = srcValues(14) ' 加
            wsOut.Cells(outRow, 14).Value = srcValues(15) ' 钣
            wsOut.Cells(outRow, 15).Value = srcValues(16) ' 备注
            ' P 计算说明
            wsOut.Cells(outRow, 16).Value = chain
            ' Q/R/S -> 零件名称/规格/标准
            wsOut.Cells(outRow, 17).Value = srcValues(17) ' 零件名称
            wsOut.Cells(outRow, 18).Value = srcValues(18) ' 规格
            wsOut.Cells(outRow, 19).Value = srcValues(19) ' 标准

            ' 若来源单元格内包含图片（形状），复制到汇总表 B 列对应单元格
            If Not oSrcWS Is Nothing And oPrevCol > 0 Then
                On Error Resume Next
                Dim copied As Long
                copied = CopyCellPictures(oSrcWS, oSrcRow, oPrevCol, wsOut, outRow, 2)
                Logger.LogInfo "T6: Copied preview pictures=" & copied & " for key='" & key & "'"
                If copied = 0 Then
                    Logger.LogWarn "T6: No preview picture found for key='" & key & "' (src='" & oSrcWS.Name & "', row=" & oSrcRow & ", col=" & oPrevCol & ")"
                End If
                On Error GoTo ROW_FAIL
            End If
        Else
            Logger.LogWarn "T6: BOM row not found for key='" & key & "'; outputting base info only"
            ' 未匹配：仍然输出基本信息（代号=key，数量=汇总），计算说明写入
            wsOut.Cells(outRow, 4).Value = key
            wsOut.Cells(outRow, 6).Value = CDbl(qty)
            wsOut.Cells(outRow, 16).Value = chain
        End If

        outRow = outRow + 1
CONTINUE_LOOP:
        On Error GoTo 0
    Next rSum

    ' 统一格式化与打印设置
    SingleSheetFormatter.IconizeBooleanFlags wsOut
    SingleSheetFormatter.ApplyFontAndAlignment wsOut
    SingleSheetFormatter.ApplyPrintSetup wsOut
    ' 隐藏“序号”列（C列）
    wsOut.Columns(3).Hidden = True

    Logger.LogInfo "T6: DONE. Output sheet='总 BOM 清单', rows=" & (outRow - 2)

CLEANUP:
    ' 关闭函数中打开的临时工作簿
    On Error Resume Next
    Dim wbTmp As Workbook
    For Each wbTmp In openedBooks
        wbTmp.Close SaveChanges:=False
    Next wbTmp
    On Error GoTo 0
    Exit Sub
ROW_FAIL:
    Logger.LogError "T6: row=" & rSum & " failed - " & Err.Description & " (key='" & key & "')"
    Err.Clear
    Resume CONTINUE_LOOP
FAIL:
    Logger.LogError "T6: failed - " & Err.Description
    Resume CLEANUP
End Sub

Private Function FindSummarySheetAndCols(ByVal wb As Workbook, ByRef cKey As Long, ByRef cQty As Long, ByRef cChain As Long, ByRef headerRow As Long) As Worksheet
    Dim ws As Worksheet
    Dim aKey, aQtyPrefer, aQtyFallback, aChain
    aKey = Array("代号", "零件号", "编码", "编号")
    aQtyPrefer = Array("总数量", "合计数量", "数量合计", "总数")
    aQtyFallback = Array("数量")
    aChain = Array("分解链", "计算说明", "展开链")

    Dim wantedKey As Object: Set wantedKey = CreateObject("Scripting.Dictionary")
    Dim wantedQtyPrefer As Object: Set wantedQtyPrefer = CreateObject("Scripting.Dictionary")
    Dim wantedQtyFallback As Object: Set wantedQtyFallback = CreateObject("Scripting.Dictionary")
    Dim wantedChain As Object: Set wantedChain = CreateObject("Scripting.Dictionary")
    Dim i As Long
    For i = LBound(aKey) To UBound(aKey): wantedKey(Utils.NormalizeName(CStr(aKey(i)))) = True: Next i
    For i = LBound(aQtyPrefer) To UBound(aQtyPrefer): wantedQtyPrefer(Utils.NormalizeName(CStr(aQtyPrefer(i)))) = True: Next i
    For i = LBound(aQtyFallback) To UBound(aQtyFallback): wantedQtyFallback(Utils.NormalizeName(CStr(aQtyFallback(i)))) = True: Next i
    For i = LBound(aChain) To UBound(aChain): wantedChain(Utils.NormalizeName(CStr(aChain(i)))) = True: Next i

    For Each ws In wb.Worksheets
        Dim maxScan As Long: maxScan = IIf(CFG_HEADER_SCAN_MAX_ROWS > 0, CFG_HEADER_SCAN_MAX_ROWS, 1)
        Dim row As Long
        For row = 1 To maxScan
            Dim lastCol As Long
            lastCol = ws.Cells(row, ws.Columns.Count).End(xlToLeft).Column
            If lastCol < 1 Then lastCol = 1
            Dim keyCol As Long: keyCol = 0
            Dim qtyPreferCol As Long: qtyPreferCol = 0
            Dim qtyFallbackCol As Long: qtyFallbackCol = 0
            Dim chainCol As Long: chainCol = 0
            For i = 1 To lastCol
                Dim h As String
                h = Utils.NormalizeName(CStr(ws.Cells(row, i).Value))
                If Len(h) > 0 Then
                    If keyCol = 0 And wantedKey.Exists(h) Then keyCol = i
                    If qtyPreferCol = 0 And wantedQtyPrefer.Exists(h) Then qtyPreferCol = i
                    If qtyFallbackCol = 0 And wantedQtyFallback.Exists(h) Then qtyFallbackCol = i
                    If chainCol = 0 And wantedChain.Exists(h) Then chainCol = i
                End If
            Next i
            Dim qtyCol As Long
            If qtyPreferCol > 0 Then qtyCol = qtyPreferCol Else qtyCol = qtyFallbackCol
            If keyCol > 0 And qtyCol > 0 And chainCol > 0 Then
                cKey = keyCol: cQty = qtyCol: cChain = chainCol: headerRow = row
                Set FindSummarySheetAndCols = ws
                Exit Function
            End If
        Next row
    Next ws
    Set FindSummarySheetAndCols = Nothing
End Function

Private Function PrepareOutputSheet(ByVal wb As Workbook) As Worksheet
    On Error Resume Next
    Application.DisplayAlerts = False
    Dim ws As Worksheet
    Set ws = Nothing
    Set ws = wb.Worksheets("总 BOM 清单")
    If Not ws Is Nothing Then ws.Delete
    Application.DisplayAlerts = True
    On Error GoTo 0

    Set PrepareOutputSheet = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
    PrepareOutputSheet.Name = "总 BOM 清单"
End Function

Private Sub WriteOutputHeader(ByVal ws As Worksheet)
    Dim headers
    headers = Array( _
        "零件号", "文档预览", "序号", "代号", "名称", "数量", _
        "材料", "处理", "渠道", "型号", _
        "组", "购", "加", "钣", _
        "备注", "计算说明", _
        "零件名称", "规格", "标准")
    Dim i As Long
    For i = LBound(headers) To UBound(headers)
        ws.Cells(1, i + 1).Value = headers(i)
    Next i
End Sub

' 在指定 BOM 工作簿中按“零件号”精确匹配 key，若命中，抽取一行的描述性字段
' 返回 True 表示找到并填充 srcValues(1..19) 其中：
'   1=零件号, 2=文档预览, 3=序号, 4=代号, 5=名称, 6=数量, 7=材料, 8=处理, 9=渠道, 10=SUPPLIER, 11=型号,
'   12=组,13=购,14=加,15=钣,16=备注,17=零件名称,18=规格,19=标准
Private Function ExtractRowByKey(ByVal wbBOM As Workbook, ByVal key As String, _
    ByVal a零件号 As Variant, ByVal a文档预览 As Variant, ByVal a序号 As Variant, ByVal a代号 As Variant, _
    ByVal a名称 As Variant, ByVal a数量 As Variant, ByVal a材料 As Variant, ByVal a处理 As Variant, _
    ByVal a渠道 As Variant, ByVal aSUP As Variant, ByVal a型号 As Variant, _
    ByVal a组 As Variant, ByVal a购 As Variant, ByVal a加 As Variant, ByVal a钣 As Variant, _
    ByVal a备注 As Variant, ByVal a零件名称 As Variant, ByVal a规格 As Variant, ByVal a标准 As Variant, _
    ByRef srcValues() As Variant, ByRef oSrcWS As Worksheet, ByRef oSrcRow As Long, ByRef oPrevCol As Long) As Boolean

    On Error GoTo FAIL
    Dim ws As Worksheet
    For Each ws In wbBOM.Worksheets
        If ws.Visible = xlSheetVisible Then
            Dim cPN As Long: cPN = Utils.GetColumnIndex(ws, a零件号)
            If cPN > 0 Then
                Dim rng As Range
                Set rng = ws.Columns(cPN).Find(What:=key, LookIn:=xlValues, LookAt:=xlWhole, _
                                               SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=False)
                If Not rng Is Nothing Then
                    ' 定位各列
                    Dim c文档预览 As Long: c文档预览 = Utils.GetColumnIndex(ws, a文档预览)
                    Dim c序号 As Long: c序号 = Utils.GetColumnIndex(ws, a序号)
                    Dim c代号 As Long: c代号 = Utils.GetColumnIndex(ws, a代号)
                    Dim c名称 As Long: c名称 = Utils.GetColumnIndex(ws, a名称)
                    Dim c数量 As Long: c数量 = Utils.GetColumnIndex(ws, a数量)
                    Dim c材料 As Long: c材料 = Utils.GetColumnIndex(ws, a材料)
                    Dim c处理 As Long: c处理 = Utils.GetColumnIndex(ws, a处理)
                    Dim c渠道 As Long: c渠道 = Utils.GetColumnIndex(ws, a渠道)
                    Dim cSUP As Long: cSUP = Utils.GetColumnIndex(ws, aSUP)
                    Dim c型号 As Long: c型号 = Utils.GetColumnIndex(ws, a型号)
                    Dim c组 As Long: c组 = Utils.GetColumnIndex(ws, a组)
                    Dim c购 As Long: c购 = Utils.GetColumnIndex(ws, a购)
                    Dim c加 As Long: c加 = Utils.GetColumnIndex(ws, a加)
                    Dim c钣 As Long: c钣 = Utils.GetColumnIndex(ws, a钣)
                    Dim c备注 As Long: c备注 = Utils.GetColumnIndex(ws, a备注)
                    Dim c零件名称 As Long: c零件名称 = Utils.GetColumnIndex(ws, a零件名称)
                    Dim c规格 As Long: c规格 = Utils.GetColumnIndex(ws, a规格)
                    Dim c标准 As Long: c标准 = Utils.GetColumnIndex(ws, a标准)

                    ' 提取
                    srcValues(1) = key
                    srcValues(2) = IIf(c文档预览 > 0, ws.Cells(rng.Row, c文档预览).Value, Empty)
                    srcValues(3) = IIf(c序号 > 0, ws.Cells(rng.Row, c序号).Value, Empty)
                    srcValues(4) = IIf(c代号 > 0, ws.Cells(rng.Row, c代号).Value, key)
                    srcValues(5) = IIf(c名称 > 0, ws.Cells(rng.Row, c名称).Value, Empty)
                    srcValues(6) = IIf(c数量 > 0, ws.Cells(rng.Row, c数量).Value, Empty)
                    srcValues(7) = IIf(c材料 > 0, ws.Cells(rng.Row, c材料).Value, Empty)
                    srcValues(8) = IIf(c处理 > 0, ws.Cells(rng.Row, c处理).Value, Empty)
                    srcValues(9) = IIf(c渠道 > 0, ws.Cells(rng.Row, c渠道).Value, Empty)
                    srcValues(10) = IIf(cSUP > 0, ws.Cells(rng.Row, cSUP).Value, Empty)
                    srcValues(11) = IIf(c型号 > 0, ws.Cells(rng.Row, c型号).Value, Empty)
                    srcValues(12) = IIf(c组 > 0, ws.Cells(rng.Row, c组).Value, Empty)
                    srcValues(13) = IIf(c购 > 0, ws.Cells(rng.Row, c购).Value, Empty)
                    srcValues(14) = IIf(c加 > 0, ws.Cells(rng.Row, c加).Value, Empty)
                    srcValues(15) = IIf(c钣 > 0, ws.Cells(rng.Row, c钣).Value, Empty)
                    srcValues(16) = IIf(c备注 > 0, ws.Cells(rng.Row, c备注).Value, Empty)
                    srcValues(17) = IIf(c零件名称 > 0, ws.Cells(rng.Row, c零件名称).Value, Empty)
                    srcValues(18) = IIf(c规格 > 0, ws.Cells(rng.Row, c规格).Value, Empty)
                    srcValues(19) = IIf(c标准 > 0, ws.Cells(rng.Row, c标准).Value, Empty)

                    Set oSrcWS = ws
                    oSrcRow = rng.Row
                    oPrevCol = c文档预览

                    ExtractRowByKey = True
                    Exit Function
                End If
            End If
        End If
    Next ws
    ExtractRowByKey = False
    Exit Function
FAIL:
    Logger.LogError "ExtractRowByKey failed on workbook '" & wbBOM.Name & "' (sheet='" & IIf(ws Is Nothing, "<unknown>", ws.Name) & "', key='" & key & "'): " & Err.Description
    ExtractRowByKey = False
End Function

Private Function CopyCellPictures(ByVal srcWS As Worksheet, ByVal srcRow As Long, ByVal srcCol As Long, _
    ByVal dstWS As Worksheet, ByVal dstRow As Long, ByVal dstCol As Long) As Long
    Dim cnt As Long: cnt = 0
    
    ' 确保显示绘图对象（避免以占位符或隐藏导致粘贴但不可见）
    Dim prevDisplay As XlDisplayDrawingObjects
    On Error Resume Next
    prevDisplay = Application.DisplayDrawingObjects
    Application.DisplayDrawingObjects = xlDisplayShapes
    On Error GoTo 0
    
    Dim cellRng As Range: Set cellRng = srcWS.Cells(srcRow, srcCol)
    Dim cellLeft As Double, cellTop As Double, cellRight As Double, cellBottom As Double
    cellLeft = cellRng.Left
    cellTop = cellRng.Top
    cellRight = cellLeft + cellRng.Width
    cellBottom = cellTop + cellRng.Height

    ' 先清理目标单元格中已有的预览形状，避免多次运行叠加/残留
    ClearCellShapes dstWS, dstRow, dstCol

    ' 1) 查找符合条件的最大形状（按面积），仅选一个避免大小不一的问题
    Dim bestShp As Shape, bestArea As Double
    Dim shp As Shape
    Dim ole As OLEObject
    bestArea = 0
    For Each shp In srcWS.Shapes
        Dim cx As Double, cy As Double
        cx = shp.Left + shp.Width / 2
        cy = shp.Top + shp.Height / 2
        If (cx > cellLeft And cx < cellRight And cy > cellTop And cy < cellBottom) Then
            Dim area As Double
            area = shp.Width * shp.Height
            If area > bestArea Then
                Set bestShp = shp
                bestArea = area
            End If
        End If
    Next shp
    
    ' 如果找到最佳形状，粘贴它
    If Not bestShp Is Nothing Then
        If PasteShapeToCell(bestShp, dstWS, dstRow, dstCol) Then cnt = cnt + 1
    End If

    ' 2) 如果 Shapes 中没找到，再从 OLEObjects 中找最大的一个
    If cnt = 0 Then
        Dim bestOle As OLEObject, bestOleArea As Double
        bestOleArea = 0
        ' 下面的循环变量 ole 需要在函数顶部声明
        For Each ole In srcWS.OLEObjects
            Dim ocx As Double, ocy As Double
            ocx = ole.Left + ole.Width / 2
            ocy = ole.Top + ole.Height / 2
            If (ocx > cellLeft And ocx < cellRight And ocy > cellTop And ocy < cellBottom) Then
                Dim oleArea As Double
                oleArea = ole.Width * ole.Height
                If oleArea > bestOleArea Then
                    Set bestOle = ole
                    bestOleArea = oleArea
                End If
            End If
        Next ole
        
        ' 如果找到最佳 OLE 对象，粘贴它
        If Not bestOle Is Nothing Then
            On Error Resume Next
            bestOle.Copy
            Dim pasted As ShapeRange
            Dim newOleShp As Shape
            Set pasted = dstWS.Shapes.Paste
            If Not pasted Is Nothing Then Set newOleShp = pasted(1)
            If newOleShp Is Nothing Then
                dstWS.Paste
                Set newOleShp = dstWS.Shapes(dstWS.Shapes.Count)
            End If
            On Error GoTo 0
            If Not newOleShp Is Nothing Then
                ResizeAndCenterShape newOleShp, dstWS, dstRow, dstCol
                cnt = cnt + 1
            End If
        End If
    End If

    ' 3) 若仍未复制到任何图片，回退：对该单元格作屏幕快照并粘贴为图片
    If cnt = 0 Then
        On Error Resume Next
        cellRng.CopyPicture Appearance:=xlScreen, Format:=xlPicture
        Dim snap As ShapeRange
        Dim newSnap As Shape
        Set snap = dstWS.Shapes.Paste
        If Not snap Is Nothing Then Set newSnap = snap(1)
        If newSnap Is Nothing Then
            dstWS.Paste
            Set newSnap = dstWS.Shapes(dstWS.Shapes.Count)
        End If
        On Error GoTo 0
        If Not newSnap Is Nothing Then
            ResizeAndCenterShape newSnap, dstWS, dstRow, dstCol
            cnt = cnt + 1
            Logger.LogInfo "T6: Fallback CopyPicture used for preview at row=" & dstRow
        Else
            Logger.LogWarn "T6: No picture copied for preview at row=" & dstRow & ", col=" & dstCol
        End If
    End If

    ' 恢复显示设置
    On Error Resume Next
    Application.DisplayDrawingObjects = prevDisplay
    On Error GoTo 0

    CopyCellPictures = cnt
End Function

Private Sub ResizeAndCenterShape(ByVal shp As Shape, ByVal dstWS As Worksheet, ByVal dstRow As Long, ByVal dstCol As Long)
    ' 若单元格过小，自动扩大列宽与行高以保证清晰可见
    Dim minWpt As Double: minWpt = 150   ' 最小列宽（点）
    Dim minHpt As Double: minHpt = 110   ' 最小行高（点）
    Dim curW As Double: curW = dstWS.Cells(dstRow, dstCol).Width
    Dim curH As Double: curH = dstWS.Cells(dstRow, dstCol).Height
    If curW < minWpt Then
        ' 将整列 B 的列宽扩大到约 26 字符宽（对应 ~200pt 左右，不同字体略异）
        If dstWS.Columns(dstCol).ColumnWidth < 26 Then dstWS.Columns(dstCol).ColumnWidth = 26
    End If
    If curH < minHpt Then
        If dstWS.Rows(dstRow).RowHeight < minHpt Then dstWS.Rows(dstRow).RowHeight = minHpt
    End If

    ' 取更新后的尺寸
    Dim cellW As Double, cellH As Double
    cellW = dstWS.Cells(dstRow, dstCol).Width
    cellH = dstWS.Cells(dstRow, dstCol).Height

    With shp
        On Error Resume Next
        .LockAspectRatio = msoTrue  ' 强制锁定纵横比
        .Placement = xlMoveAndSize
        .PrintObject = True
        .Visible = msoTrue
        On Error GoTo 0
        
        ' 计算等比缩放系数，严格按原图比例缩放（留4pt边距）
        Dim availW As Double, availH As Double
        availW = cellW - 4  ' 左右各留2pt
        availH = cellH - 4  ' 上下各留2pt
        
        Dim scaleW As Double, scaleH As Double, finalScale As Double
        scaleW = availW / .Width
        scaleH = availH / .Height
        finalScale = scaleW
        If scaleH < finalScale Then finalScale = scaleH  ' 取小者确保完全适配
        If finalScale <= 0 Then finalScale = 0.1  ' 防止异常
        
        ' 应用等比缩放
        .Width = .Width * finalScale
        .Height = .Height * finalScale
        
        ' 精确居中：计算单元格中心，减去图片中心偏移
        Dim cellCenterX As Double, cellCenterY As Double
        cellCenterX = dstWS.Cells(dstRow, dstCol).Left + cellW / 2
        cellCenterY = dstWS.Cells(dstRow, dstCol).Top + cellH / 2
        
        .Left = cellCenterX - .Width / 2
        .Top = cellCenterY - .Height / 2
        
        ' 置顶以避免被其他对象（如布尔图标）遮挡
        On Error Resume Next
        .ZOrder msoBringToFront
        On Error GoTo 0
    End With
End Sub

Private Function PasteShapeToCell(ByVal srcShp As Shape, ByVal dstWS As Worksheet, ByVal dstRow As Long, ByVal dstCol As Long) As Boolean
    Dim pasted As ShapeRange
    Dim newShp As Shape
    PasteShapeToCell = False
    On Error Resume Next
    ' 优先：图片类型用直接复制粘贴，保持分辨率（必要时 Excel 会以原位图/矢量复制）
    If srcShp.Type = msoPicture Or srcShp.Type = msoLinkedPicture Then
        srcShp.Copy
        Set pasted = dstWS.Shapes.Paste
        If Not pasted Is Nothing Then Set newShp = pasted(1)
    End If
    ' 回退：以打印质量位图粘贴
    If newShp Is Nothing Then
        srcShp.CopyPicture Appearance:=xlPrinter, Format:=xlBitmap
        Set pasted = dstWS.Shapes.Paste
        If Not pasted Is Nothing Then Set newShp = pasted(1)
        If newShp Is Nothing Then
            dstWS.Paste
            Set newShp = dstWS.Shapes(dstWS.Shapes.Count)
        End If
    End If
    Application.CutCopyMode = False
    On Error GoTo 0
    If Not newShp Is Nothing Then
        ResizeAndCenterShape newShp, dstWS, dstRow, dstCol
        PasteShapeToCell = True
    End If
End Function

Private Sub GatherBOMWorkbooks(ByVal baseDir As String, ByVal wbSummary As Workbook, _
    ByRef bomBooks As Collection, ByRef openedBooks As Collection)
    Dim f As String
    f = Dir(baseDir & "\*.xls*")
    Do While Len(f) > 0
        If InStr(1, f, "汇总", vbTextCompare) = 0 Then
            If StrComp(UCase$(f), UCase$(CFG_MAPPING_WORKBOOK_NAME), vbTextCompare) <> 0 Then
                ' 忽略当前汇总驱动工作簿自身以外的 BOM 文件
                Dim wb As Workbook
                Set wb = Utils.FindOpenWorkbookByName(f)
                If wb Is Nothing Then
                    On Error Resume Next
                    Set wb = Application.Workbooks.Open(FileName:=baseDir & "\" & f, ReadOnly:=True)
                    If Not wb Is Nothing Then openedBooks.Add wb
                    On Error GoTo 0
                End If
                If Not wb Is Nothing Then
                    bomBooks.Add wb
                End If
            End If
        End If
        f = Dir()
    Loop

    ' 也尝试把当前汇总文件对应的“同名主 BOM 文件”（去掉“_汇总”后）加入搜索优先队列前部
    Dim nameNoSum As String
    nameNoSum = Replace(wbSummary.Name, "_汇总", "")
    If StrComp(nameNoSum, wbSummary.Name, vbTextCompare) <> 0 Then
        Dim wbMain As Workbook: Set wbMain = Utils.FindOpenWorkbookByName(nameNoSum)
        If wbMain Is Nothing Then
            On Error Resume Next
            Set wbMain = Application.Workbooks.Open(FileName:=baseDir & "\" & nameNoSum, ReadOnly:=True)
            If Not wbMain Is Nothing Then openedBooks.Add wbMain
            On Error GoTo 0
        End If
        If Not wbMain Is Nothing Then
            ' 将其插入到集合前端（简单实现：先新建一个集合把它放第一位）
            Dim tmp As New Collection
            tmp.Add wbMain
            Dim it As Variant
            For Each it In bomBooks
                tmp.Add it
            Next it
            Set bomBooks = tmp
        End If
    End If
End Sub

Private Sub ClearCellShapes(ByVal ws As Worksheet, ByVal row As Long, ByVal col As Long)
    On Error Resume Next
    Dim L As Double, T As Double, R As Double, B As Double
    L = ws.Cells(row, col).Left
    T = ws.Cells(row, col).Top
    R = L + ws.Cells(row, col).Width
    B = T + ws.Cells(row, col).Height
    Dim s As Shape
    For Each s In ws.Shapes
        Dim cx As Double, cy As Double
        cx = s.Left + s.Width / 2
        cy = s.Top + s.Height / 2
        If (cx > L And cx < R And cy > T And cy < B) Then
            s.Delete
        End If
    Next s
    On Error GoTo 0
End Sub