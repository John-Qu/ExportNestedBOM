' RecursiveProcessor.bas - core recursion and export logic
Option Explicit

Public Sub ProcessDrawingRecursive(swApp As Object, drawingPath As String, depth As Integer, parentQty As Long, _
    ByRef visited As Object, ByRef summary As Object, ByVal topAsmName As String, ByVal parentChain As String)
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
    
    ' 导出当前BOM为Excel(xls)并包含图片
    Dim outXls As String
    outXls = GetFileFolder(drawingPath) & "\" & GetFileNameNoExt(drawingPath) & ".xls"
    Dim ok As Boolean
    On Error GoTo EH_Export
    ok = bomAnn.SaveAsExcel(outXls, True, True) ' 包含隐藏列与图片
    Logger_Info String(depth * 2, " ") & "导出BOM：" & outXls & " => " & ok
    On Error GoTo EH
    
    ' 遍历BOM行，识别"是否组装"列，递归
    ProcessBOMRows bomAnn, swApp, drawingPath, depth, parentQty, visited, summary, topAsmName, parentChain

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

Private Function FindFirstBOM(swDraw As Object) As Object
    ' 遍历Feature树找到第一个BOM
    Dim feat As Object: Set feat = swDraw.FirstFeature
    Do While Not feat Is Nothing
        If feat.GetTypeName = "BomFeat" Then
            Dim bf As Object: Set bf = feat.GetSpecificFeature2
            Dim tables As Variant: tables = bf.GetTableAnnotations
            If Not IsEmpty(tables) Then
                Dim ta As Object: Set ta = tables(0)
                Dim ba As Object: Set ba = ta ' ITableAnnotation -> IBomTableAnnotation
                Set FindFirstBOM = ba
                Exit Function
            End If
        End If
        Set feat = feat.GetNextFeature
    Loop
End Function

Public Sub ProcessBOMRows(bomAnn As Object, swApp As Object, drawingPath As String, depth As Integer, parentQty As Long, _
    ByRef visited As Object, ByRef summary As Object, ByVal topAsmName As String, ByVal parentChain As String)
    On Error GoTo EH
    Dim ta As Object: Set ta = bomAnn ' TableAnnotation
    Dim rows As Long: rows = ta.RowCount ' <mcreference link="https://help.solidworks.com/2017/English/api/sldworksapi/SolidWorks.Interop.sldworks~SolidWorks.Interop.sldworks.ITableAnnotation~RowCount.html" index="2">2</mcreference>
    Dim cols As Long: cols = ta.ColumnCount ' <mcreference link="https://help.solidworks.com/2016/english/api/sldworksapi/SolidWorks.Interop.sldworks~SolidWorks.Interop.sldworks.ITableAnnotation~ColumnCount.html" index="3">3</mcreference>
    
    Dim colQty As Long: colQty = FindColumnIndex(ta, Array("数量", "QTY", "Qty"))
    Dim colName As Long: colName = FindColumnIndex(ta, Array("名称", "PART NAME", "Name"))
    Dim colPartNumber As Long: colPartNumber = FindColumnIndex(ta, Array("代号", "PART NUMBER", "Part Number", "PARTPATH", "零件路径"))
    Dim colAssemble As Long: colAssemble = FindColumnIndex(ta, Array("是否组装", "Is Assembly", "组装", "是否组件", "IS ASSEMBLY"))
    Dim colItemNumber As Long: colItemNumber = FindColumnIndex(ta, Array("项目号", "ITEM NO", "Item", "Item Number", "序号", "项号"))
    Dim colPreview As Long: colPreview = 0 ' 第一列通常为缩略图
    
    If colQty < 0 Then
        Logger_Warn "未定位到数量列，默认尝试第D列(3)"
        colQty = 3 ' 常见：A预览 B项目号 C PART NUMBER D数量
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
        colItemNumber = 1 ' 常见：A预览 B项目号 C PART NUMBER D数量
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
        ' 从工程图文件名中提取代号，以空格为分隔符取前半部分的字母和数字
        Dim asmCode As String
        asmCode = ExtractPartCode(GetFileNameNoExt(drawingPath))
        
        If Len(parentChain) = 0 Then
            ' 顶层装配体，格式：装配体代号#项目号: 数量
            currentChain = asmCode & "#" & itemNumber & ": " & qty
        Else
            ' 子装配体，格式：子装配体代号#项目号: 数量 x 上层装配体代号#项目号: 数量
            ' 使用当前装配体的代号（从工程图路径获取），而不是partNo
            currentChain = asmCode & "#" & itemNumber & ": " & Val(ta.Text(i, colQty)) & " x " & parentChain
        End If
        
        If isAsm Then
            ' 打开同名子装配体工程图递归
            Dim childDrw As String
            childDrw = GetFileFolder(drawingPath) & "\" & partNo & ".slddrw"
            If FileExists(childDrw) Then
                ProcessDrawingRecursive swApp, childDrw, depth + 1, qty, visited, summary, topAsmName, currentChain
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

Private Function IsAssembleCell(valText As String) As Boolean
    Dim t As String: t = UCase$(Trim$(valText))
    IsAssembleCell = (t = "是" Or t = "Y" Or t = "YES" Or t = "TRUE" Or t = "1")
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