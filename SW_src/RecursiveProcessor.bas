' RecursiveProcessor.bas - core recursion and export logic
Option Explicit

Public Sub ProcessDrawingRecursive(swApp As Object, drawingPath As String, depth As Integer, parentQty As Long, _
    ByRef visited As Object, ByRef summary As Object, ByVal topAsmName As String, ByVal parentChain As String, ByVal sessionOutDir As String)
    On Error GoTo EH
    If depth > 10 Then
        Logger_Warn "递归深度超限 (>10)：" & drawingPath
        Exit Sub
    End If
    
    If Not FileExists(drawingPath) Then
        Logger_Error "工程图文件不存在：" & drawingPath
        Exit Sub
    End If
    
    Dim key As String: key = LCase$(drawingPath)
    If visited.Exists(key) Then
        MsgBox "检测到循环引用：" & GetFileNameNoExt(drawingPath), vbExclamation
        Logger_Warn "循环引用：" & drawingPath
        Exit Sub
    End If
    visited.Add key, True
    
    Logger_Info String(depth * 2, " ") & "处理 [深度" & depth & "]：" & GetFileNameNoExt(drawingPath)
    
    Dim errs As Long, warns As Long
    Dim swModel As Object
    Set swModel = swApp.OpenDoc6(drawingPath, 3, 1, "", errs, warns) ' swDocDRAWING = 3
    If swModel Is Nothing Then
        Logger_Error "无法打开工程图：" & drawingPath & " (错误:" & errs & ", 警告:" & warns & ")"
        GoTo Clean
    End If
    
    If errs <> 0 Then
        Logger_Warn "打开文档时出现错误 (" & errs & ")：" & drawingPath
    End If
    If warns <> 0 Then
        Logger_Warn "打开文档时出现警告 (" & warns & ")：" & drawingPath
    End If
    
    Dim swDraw As Object: Set swDraw = swModel
    ' 获取第一张BOM表
    Dim bomAnn As Object ' BomTableAnnotation
    Set bomAnn = FindFirstBOM(swDraw)
    If bomAnn Is Nothing Then
        Logger_Warn "未找到BOM表：" & drawingPath
        GoTo CloseDoc
    End If
    
    ' 导出当前BOM为Excel(xls)并包含图片（统一输出到 sessionOutDir）
    Dim outXls As String
    outXls = sessionOutDir & "\" & GetFileNameNoExt(drawingPath) & ".xls"
    Dim ok As Boolean
    If FileExists(outXls) Then
        Logger_Info String(depth * 2, " ") & "检测到已有同名BOM，跳过导出：" & outXls
    Else
        On Error GoTo EH_Export
        ok = bomAnn.SaveAsExcel(outXls, True, True) ' 包含隐藏列与图片
        Logger_Info String(depth * 2, " ") & "导出BOM：" & outXls & " => " & ok
        On Error GoTo EH
    End If
    
    ' 遍历BOM行，识别"是否组装"列，递归
    ProcessBOMRows bomAnn, swApp, drawingPath, depth, parentQty, visited, summary, topAsmName, parentChain, sessionOutDir

CloseDoc:
    On Error Resume Next
    swApp.CloseDoc swModel.GetTitle
    On Error GoTo EH
Clean:
    On Error Resume Next
    visited.Remove key
    On Error GoTo EH
    Exit Sub
    
EH_Export:
    Logger_Error "BOM导出失败：" & drawingPath & " => " & Err.Description
    Resume Next
    
EH:
    Logger_Error "ProcessDrawingRecursive 出错：" & Err.Number & ": " & Err.Description & " (文件:" & drawingPath & ")"
    Resume Clean
End Sub

Public Sub ProcessBOMRows(bomAnn As Object, swApp As Object, drawingPath As String, depth As Integer, parentQty As Long, _
    ByRef visited As Object, ByRef summary As Object, ByVal topAsmName As String, ByVal parentChain As String, ByVal sessionOutDir As String)
    On Error GoTo EH
    Dim ta As Object: Set ta = bomAnn ' TableAnnotation
    Dim rows As Long: rows = ta.RowCount ' ...
    Dim cols As Long: cols = ta.ColumnCount ' ...
    
    Dim colItemNumber As Long: colItemNumber = FindColumnIndex(ta, Array("项目号", "ITEM NO", "Item", "Item Number", "序号", "项号"))
    Dim colQty As Long: colQty = DetectQuantityColumn(ta, colItemNumber)
    Dim colName As Long: colName = FindColumnIndex(ta, GetNameColumnNames())
    Dim colPartNumber As Long: colPartNumber = FindColumnIndex(ta, GetPartNumberColumnNames())
    Dim colAssemble As Long: colAssemble = FindColumnIndex(ta, GetAssemblyColumnNames())
    Dim colPreview As Long: colPreview = 0 ' 第一列通常为缩略图
    
    If colQty < 0 Then
        Logger_Warn "未定位到数量列，采用最佳猜测列索引"
        colQty = 3
    End If
    If colPartNumber < 0 Then
        Logger_Warn "未定位到PART NUMBER/代号列，默认尝试第C列(2)"
        colPartNumber = 2
    End If
    If colName < 0 Then
        Logger_Warn "未定位到名称列，默认尝试第F列(5)"
        colName = 5
    End If
    If colAssemble < 0 Then
        Logger_Warn "未定位到是否组装列，将按'是/Yes/Y/True/1'在最后一列尝试匹配"
        colAssemble = ta.ColumnCount - 1
    End If
    If colItemNumber < 0 Then
        Logger_Warn "未定位到项目号列，默认尝试第B列(1)"
        colItemNumber = 1
    End If
    
    Dim i As Long
    For i = 1 To rows - 1 ' 跳过标题行0
        Dim qty As Long: qty = Val(ta.Text(i, colQty))
        If parentQty > 0 Then qty = qty * parentQty
        Dim isAsm As Boolean: isAsm = IsAssembleCell(ta.Text(i, colAssemble))
        Dim partNo As String: partNo = Trim$(ta.Text(i, colPartNumber))
        Dim partName As String: partName = Trim$(ta.Text(i, colName))
        
        ' 获取项目号，如果项目号列为空则使用行索引
        Dim itemNumber As String
        itemNumber = Trim$(ta.Text(i, colItemNumber))
        If Len(itemNumber) = 0 Then
            itemNumber = CStr(i)
        End If
        
        Dim currentChain As String
        Dim asmCode As String
        asmCode = ExtractPartCode(GetFileNameNoExt(drawingPath))
        
        If Len(parentChain) = 0 Then
            currentChain = asmCode & "#" & itemNumber & ": " & qty
        Else
            currentChain = asmCode & "#" & itemNumber & ": " & Val(ta.Text(i, colQty)) & " x " & parentChain
        End If
        
        If isAsm Then
            ' 打开同名子装配体工程图递归
            Dim childDrw As String
            childDrw = GetFileFolder(drawingPath) & "\" & partNo & ".slddrw"
            If FileExists(childDrw) Then
                ProcessDrawingRecursive swApp, childDrw, depth + 1, qty, visited, summary, topAsmName, currentChain, sessionOutDir
            Else
                Logger_Warn "未找到子装配体工程图：" & childDrw
            End If
        Else
            ' 底层零件，加入汇总
            AddToSummary summary, partNo, partName, qty, currentChain
        End If
    Next
    Exit Sub
EH:
    Logger_Error "ProcessBOMRows 出错：" & Err.Number & ": " & Err.Description
End Sub

Private Function FindColumnIndex(ta As Object, names As Variant) As Long
    Dim c As Long
    For c = 0 To ta.ColumnCount - 1
        Dim title As String: title = UCase$(Trim$(ta.Text(0, c)))
        Dim i As Long
        For i = LBound(names) To UBound(names)
            If InStr(title, UCase$(names(i))) > 0 Then ' 使用包含匹配而非精确匹配
                FindColumnIndex = c
                Exit Function
            End If
        Next
    Next
    FindColumnIndex = -1
End Function

Private Function DetectQuantityColumn(ta As Object, ByVal colItemNumber As Long) As Long
    Dim idx As Long
    idx = FindColumnIndex(ta, GetQuantityColumnNames())
    If idx >= 0 Then
        DetectQuantityColumn = idx
        Exit Function
    End If

    Dim rows As Long: rows = ta.RowCount
    Dim limit As Long: limit = rows - 1
    If limit > 50 Then limit = 50

    Dim eIdx As Long: eIdx = 4 ' E列
    Dim dIdx As Long: dIdx = 3 ' D列
    Dim eScore As Double, dScore As Double
    eScore = ColumnNumericRatio(ta, eIdx, limit)
    dScore = ColumnNumericRatio(ta, dIdx, limit)

    If eIdx <> colItemNumber And eScore >= 0.6 Then
        DetectQuantityColumn = eIdx
        Exit Function
    End If
    If dIdx <> colItemNumber And dScore >= 0.6 Then
        DetectQuantityColumn = dIdx
        Exit Function
    End If

    Dim bestIdx As Long: bestIdx = -1
    Dim bestScore As Double: bestScore = 0
    Dim c As Long
    For c = 0 To ta.ColumnCount - 1
        If c <> colItemNumber Then
            Dim s As Double: s = ColumnNumericRatio(ta, c, limit)
            If s > bestScore Then
                bestScore = s
                bestIdx = c
            End If
        End If
    Next c
    If bestScore >= 0.6 Then
        DetectQuantityColumn = bestIdx
    Else
        DetectQuantityColumn = -1
    End If
End Function

Private Function ColumnNumericRatio(ta As Object, ByVal col As Long, ByVal limit As Long) As Double
    If col < 0 Or col > ta.ColumnCount - 1 Then
        ColumnNumericRatio = 0
        Exit Function
    End If
    Dim i As Long, cnt As Long, num As Long
    cnt = 0: num = 0
    For i = 1 To limit
        Dim t As String: t = Trim$(ta.Text(i, col))
        If Len(t) > 0 Then
            cnt = cnt + 1
            If IsNumeric(t) Then
                num = num + 1
            End If
        End If
    Next i
    If cnt = 0 Then
        ColumnNumericRatio = 0
    Else
        ColumnNumericRatio = num / cnt
    End If
End Function

Private Function IsAssembleCell(valText As String) As Boolean
    Dim t As String: t = UCase$(Trim$(valText))
    Dim names As Variant: names = GetAssemblyTrueValues()
    Dim i As Long
    For i = LBound(names) To UBound(names)
        If UCase$(CStr(names(i))) = t Then
            IsAssembleCell = True
            Exit Function
        End If
    Next i
    IsAssembleCell = False
End Function

Private Sub AddToSummary(ByRef summary As Object, ByVal partNo As String, ByVal partName As String, ByVal qty As Long, ByVal chain As String)
    Dim key As String: key = partNo
    Dim item As Object
    If summary.Exists(key) Then
        Set item = summary(key)
        Dim oldTotal As Long: oldTotal = CLng(item("TotalQty"))
        item("TotalQty") = oldTotal + qty
        ' 格式：A条线 + B条线 => 总数量
        item("Breakdown") = CStr(item("Breakdown")) & " + " & chain & " => " & (oldTotal + qty)
    Else
        Set item = CreateObject("Scripting.Dictionary")
        item.Add "PartNo", partNo
        item.Add "PartName", partName
        item.Add "TotalQty", qty
        ' 首次添加，格式：A条线 => 数量
        item.Add "Breakdown", chain & " => " & qty
        summary.Add key, item
    End If
End Sub

' 新增：导出前的子装配参与性确认（生成CSV并提示是否继续）
Public Function ConfirmSubAssemblyParticipation(swApp As Object, drawingPath As String) As Boolean
    On Error GoTo EH
    Dim items As Collection: Set items = New Collection
    Dim visited As Object: Set visited = CreateObject("Scripting.Dictionary")

    ' 扫描顶层工程图的子装配状态（递归深度与导出一致，但不执行导出）
    ScanParticipationRecursive swApp, drawingPath, 1, visited, items

    ' 统计
    Dim totalFlagYes As Long, readyCount As Long
    Dim cntNoDrw As Long, cntNoBom As Long, cntMissFlag As Long
    Dim it As Variant
    For Each it In items
        If CBool(it("IsAsmFlag")) Then
            totalFlagYes = totalFlagYes + 1
        End If
        Select Case CStr(it("Status"))
            Case "Included-Ready"
                readyCount = readyCount + 1
            Case "Skipped-NoDrawing"
                cntNoDrw = cntNoDrw + 1
            Case "Skipped-NoBOMTable"
                cntNoBom = cntNoBom + 1
            Case "Skipped-PropertyMissing"
                cntMissFlag = cntMissFlag + 1
        End Select
    Next

    Dim coverage As Double
    If totalFlagYes > 0 Then coverage = readyCount / totalFlagYes Else coverage = 1#

    ' 写出CSV检查表
    Dim outCsv As String
    outCsv = GetFileFolder(drawingPath) & "\" & GetFileNameNoExt(drawingPath) & "_参与性确认.csv"
    WriteParticipationCsv items, outCsv

    ' 汇总与交互
    Dim msg As String
    msg = "子装配参与性确认：" & vbCrLf & _
          "标注“是”的子装配： " & totalFlagYes & vbCrLf & _
          "可导出(有图+有BOM)： " & readyCount & vbCrLf & _
          "缺工程图： " & cntNoDrw & vbCrLf & _
          "无BOM表： " & cntNoBom & vbCrLf & _
          "疑似漏标(发现同名工程图)： " & cntMissFlag & vbCrLf & _
          "覆盖率： " & Format(coverage, "0.0%") & vbCrLf & vbCrLf & _
          "已生成检查表：" & outCsv & vbCrLf & _
          "是否继续执行导出？"

    Logger_Info "参与性确认输出：" & outCsv & _
                " | 标注是=" & totalFlagYes & ", 可导出=" & readyCount & _
                ", 无图=" & cntNoDrw & ", 无BOM=" & cntNoBom & ", 疑似漏标=" & cntMissFlag

    If (cntNoDrw + cntNoBom + cntMissFlag) > 0 And UCase$(CONFIRM_BLOCK_ON_SKIPPED) = "BLOCK" Then
        MsgBox "检测到阻断性问题（依据配置），流程中止。" & vbCrLf & vbCrLf & msg, vbExclamation
        ConfirmSubAssemblyParticipation = False
        Exit Function
    End If

    Dim ans As VbMsgBoxResult
    ans = MsgBox(msg, vbQuestion + vbYesNo, "子装配参与性确认")
    ConfirmSubAssemblyParticipation = (ans = vbYes)
    If ConfirmSubAssemblyParticipation Then
        ' 写入确认标记文件，后续可跳过重复检查
        WriteConfirmOk drawingPath, totalFlagYes, readyCount, cntNoDrw, cntNoBom, cntMissFlag
        Logger_Info "已写入确认标记：" & GetConfirmOkPath(drawingPath)
    End If
    Exit Function
EH:
    Logger_Error "参与性确认出错：" & Err.Number & ": " & Err.Description
    ConfirmSubAssemblyParticipation = True ' 容错：确认环节失败不阻断（可按需调整）
End Function

' 新增：确认标记文件路径
Private Function GetConfirmOkPath(drawingPath As String) As String
    GetConfirmOkPath = GetFileFolder(drawingPath) & "\" & GetFileNameNoExt(drawingPath) & "_参与性确认.ok"
End Function

' 新增：是否已存在确认标记
Public Function HasConfirmOk(drawingPath As String) As Boolean
    HasConfirmOk = FileExists(GetConfirmOkPath(drawingPath))
End Function

' 新增：写入确认标记文件
Private Sub WriteConfirmOk(drawingPath As String, totalFlagYes As Long, readyCount As Long, cntNoDrw As Long, cntNoBom As Long, cntMissFlag As Long)
    On Error GoTo EH
    Dim p As String: p = GetConfirmOkPath(drawingPath)
    Dim f As Integer: f = FreeFile
    Open p For Output As #f
    Print #f, "timestamp=" & Format$(Now, "yyyy-mm-dd hh:nn:ss")
    Print #f, "total_flag_yes=" & totalFlagYes
    Print #f, "ready_count=" & readyCount
    Print #f, "missing_drawing=" & cntNoDrw
    Print #f, "missing_bom=" & cntNoBom
    Print #f, "suspect_missing_flag=" & cntMissFlag
    Close #f
    Exit Sub
EH:
    On Error Resume Next
    Close #f
    Logger_Warn "确认标记写入失败：" & Err.Description
End Sub

' 新增：递归扫描参与性状态（不导出，只检查）
Private Sub ScanParticipationRecursive(swApp As Object, drawingPath As String, depth As Integer, _
    ByRef visited As Object, ByRef items As Collection)
    On Error GoTo EH
    If depth > 10 Then Exit Sub
    If Not FileExists(drawingPath) Then Exit Sub

    Dim key As String: key = "scan|" & LCase$(drawingPath)
    If visited.Exists(key) Then Exit Sub
    visited.Add key, True

    Dim errs As Long, warns As Long
    Dim swModel As Object
    Set swModel = swApp.OpenDoc6(drawingPath, 3, 1, "", errs, warns) ' swDocDRAWING=3
    If swModel Is Nothing Then GoTo Clean

    Dim swDraw As Object: Set swDraw = swModel
    Dim bomAnn As Object: Set bomAnn = FindFirstBOM(swDraw)

    If bomAnn Is Nothing Then
        ' 顶层无BOM：无从扫描子装配，记录为提示但不中止
        Logger_Warn "确认阶段：未找到BOM表（扫描受限）：" & drawingPath
        GoTo CloseDoc
    End If

    Dim ta As Object: Set ta = bomAnn ' TableAnnotation
    Dim rows As Long: rows = ta.RowCount
    Dim colQty As Long: colQty = FindColumnIndex(ta, Array("数量", "QTY", "Qty"))
    Dim colName As Long: colName = FindColumnIndex(ta, Array("名称", "PART NAME", "Name"))
    Dim colPartNumber As Long: colPartNumber = FindColumnIndex(ta, Array("零件号", "PART NUMBER", "Part Number", "PARTPATH", "零件路径"))
    Dim colAssemble As Long: colAssemble = FindColumnIndex(ta, Array("是否组装", "Is Assembly", "组装", "是否组件", "IS ASSEMBLY"))

    If colQty < 0 Then colQty = 3
    If colPartNumber < 0 Then colPartNumber = 2
    If colName < 0 Then colName = 5
    If colAssemble < 0 Then colAssemble = ta.ColumnCount - 1

    Dim i As Long
    For i = 1 To rows - 1
        Dim partNo As String: partNo = Trim$(ta.Text(i, colPartNumber))
        Dim partName As String: partName = Trim$(ta.Text(i, colName))
        Dim flagIsAsm As Boolean: flagIsAsm = IsAssembleCell(ta.Text(i, colAssemble))

        ' Heuristic：同目录下是否存在同名工程图
        Dim childDrw As String
        childDrw = GetFileFolder(drawingPath) & "\" & partNo & ".slddrw"
        Dim hasDrw As Boolean: hasDrw = FileExists(childDrw)
        Dim hasBom As Boolean: hasBom = False

        Dim status As String, reason As String
        If flagIsAsm Then
            If Not hasDrw Then
                status = "Skipped-NoDrawing": reason = "标注为子装配但未找到工程图"
            Else
                ' 检查子工程图是否有BOM
                Dim e1 As Long, w1 As Long, m As Object
                Set m = swApp.OpenDoc6(childDrw, 3, 1, "", e1, w1)
                If Not m Is Nothing Then
                    Dim b As Object: Set b = FindFirstBOM(m)
                    If Not b Is Nothing Then
                        hasBom = True
                        status = "Included-Ready": reason = "可导出（有图+有BOM）"
                        ' 递归扫描更深层级
                        ScanParticipationRecursive swApp, childDrw, depth + 1, visited, items
                    Else
                        status = "Skipped-NoBOMTable": reason = "工程图中未找到BOM表"
                    End If
                    On Error Resume Next
                    swApp.CloseDoc m.GetTitle
                    On Error GoTo 0
                Else
                    status = "Skipped-NoDrawing": reason = "工程图无法打开"
                End If
            End If
        Else
            If hasDrw Then
                status = "Skipped-PropertyMissing": reason = "疑似漏标（存在同名工程图）"
            Else
                ' 非子装配，且无同名工程图：对确认表噪音较大，可跳过不记录
                GoTo ContinueNext
            End If
        End If

        ' 记录条目
        Dim item As Object: Set item = CreateObject("Scripting.Dictionary")
        item.Add "PartNo", partNo
        item.Add "PartName", partName
        item.Add "IsAsmFlag", flagIsAsm
        item.Add "DrawingExists", hasDrw
        item.Add "BomExists", hasBom
        item.Add "Status", status
        item.Add "Reason", reason
        items.Add item

ContinueNext:
    Next

CloseDoc:
    On Error Resume Next
    swApp.CloseDoc swModel.GetTitle
    On Error GoTo 0
Clean:
    On Error Resume Next
    visited.Remove key
    On Error GoTo 0
    Exit Sub
EH:
    Logger_Error "ScanParticipationRecursive 出错：" & Err.Number & ": " & Err.Description & " (文件:" & drawingPath & ")"
    Resume Clean
End Sub

' 新增：输出CSV检查表
Private Sub WriteParticipationCsv(items As Collection, csvPath As String)
    On Error GoTo EH
    Dim f As Integer: f = FreeFile
    Open csvPath For Output As #f
    Print #f, "PartNo,PartName,IsAsmFlag,DrawingExists,BomExists,Status,Reason"
    Dim it As Variant
    For Each it In items
        Print #f, CsvEscape(it("PartNo")) & "," & _
                  CsvEscape(it("PartName")) & "," & _
                  IIf(CBool(it("IsAsmFlag")), "1", "0") & "," & _
                  IIf(CBool(it("DrawingExists")), "1", "0") & "," & _
                  IIf(CBool(it("BomExists")), "1", "0") & "," & _
                  CsvEscape(it("Status")) & "," & _
                  CsvEscape(it("Reason"))
    Next
    Close #f
    Exit Sub
EH:
    On Error Resume Next
    Close #f
    On Error GoTo 0
    Logger_Error "写入参与性CSV失败：" & Err.Number & ": " & Err.Description
End Sub

Private Function CsvEscape(ByVal s As String) As String
    If InStr(s, ",") > 0 Or InStr(s, """") > 0 Or InStr(s, vbCr) > 0 Or InStr(s, vbLf) > 0 Then
        s = Replace$(s, """", """""")
        CsvEscape = """" & s & """"
    Else
        CsvEscape = s
    End If
End Function

Private Function FindFirstBOM(swDraw As Object) As Object
    On Error GoTo EH
    ' 遍历Feature树找到第一个BOM
    Dim feat As Object: Set feat = swDraw.FirstFeature
    Do While Not feat Is Nothing
        Dim t As String: t = feat.GetTypeName
        If t = "BomFeat" Or t = "BomFeature" Then
            Dim bf As Object: Set bf = feat.GetSpecificFeature2
            If Not bf Is Nothing Then
                Dim tables As Variant: tables = bf.GetTableAnnotations
                If Not IsEmpty(tables) Then
                    Dim ta As Object: Set ta = tables(0) ' ITableAnnotation
                    Dim ba As Object: Set ba = ta       ' IBomTableAnnotation（晚绑定）
                    Set FindFirstBOM = ba
                    Exit Function
                End If
            End If
        End If
        Set feat = feat.GetNextFeature
    Loop
    Exit Function
EH:
    Logger_Warn "查找BOM表失败：" & Err.Number & ": " & Err.Description
    Set FindFirstBOM = Nothing
End Function