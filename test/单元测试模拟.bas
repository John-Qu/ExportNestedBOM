' 单元测试模拟.bas - 用于验证核心功能的测试代码
Option Explicit

' 测试主入口
Public Sub RunAllTests()
    Debug.Print "=== 开始单元测试 ==="
    
    TestUtils
    TestColumnFinder  
    TestAssembleDetection
    TestSummaryLogic
    
    Debug.Print "=== 单元测试完成 ==="
End Sub

' 测试工具函数
Private Sub TestUtils()
    Debug.Print "--- 测试工具函数 ---"
    
    ' 测试文件名提取
    Dim result As String
    result = GetFileNameNoExt("C:\Test\TS180-01Z-01Z四柱平台.slddrw")
    Debug.Print "GetFileNameNoExt: " & result & " (预期: TS180-01Z-01Z四柱平台)"
    
    ' 测试文件夹提取
    result = GetFileFolder("C:\Test\TS180-01Z-01Z四柱平台.slddrw") 
    Debug.Print "GetFileFolder: " & result & " (预期: C:\Test)"
    
    ' 测试HTML编码
    result = HtmlEncode("<test>&""value""")
    Debug.Print "HtmlEncode: " & result & " (预期: &lt;test&gt;&amp;&quot;value&quot;)"
End Sub

' 测试列查找功能
Private Sub TestColumnFinder()
    Debug.Print "--- 测试列查找功能 ---"
    
    ' 模拟表格对象
    Set mockTable = CreateMockTable()
    
    ' 测试查找数量列
    Dim colIndex As Long
    colIndex = FindColumnIndex(mockTable, Array("数量", "QTY", "Qty"))
    Debug.Print "数量列索引: " & colIndex & " (预期: 2)"
    
    ' 测试查找代号列  
    colIndex = FindColumnIndex(mockTable, Array("代号", "PART NUMBER", "Part Number"))
    Debug.Print "代号列索引: " & colIndex & " (预期: 1)"
    
    ' 测试查找不存在的列
    colIndex = FindColumnIndex(mockTable, Array("不存在的列"))
    Debug.Print "不存在列索引: " & colIndex & " (预期: -1)"
End Sub

' 测试是否组装判断
Private Sub TestAssembleDetection() 
    Debug.Print "--- 测试是否组装判断 ---"
    
    Debug.Print "是否组装('是'): " & IsAssembleCell("是") & " (预期: True)"
    Debug.Print "是否组装('YES'): " & IsAssembleCell("YES") & " (预期: True)"
    Debug.Print "是否组装('y'): " & IsAssembleCell("y") & " (预期: True)"  
    Debug.Print "是否组装('True'): " & IsAssembleCell("True") & " (预期: True)"
    Debug.Print "是否组装('1'): " & IsAssembleCell("1") & " (预期: True)"
    Debug.Print "是否组装('否'): " & IsAssembleCell("否") & " (预期: False)"
    Debug.Print "是否组装(''): " & IsAssembleCell("") & " (预期: False)"
End Sub

' 测试汇总逻辑
Private Sub TestSummaryLogic()
    Debug.Print "--- 测试汇总逻辑 ---"
    
    Dim summary As Object
    Set summary = CreateObject("Scripting.Dictionary")
    
    ' 添加第一个零件
    AddToSummary summary, "TS180-01Z-003", "四孔平板", 1, "四柱平台#10: 1"
    
    ' 添加同一零件的另一个实例  
    AddToSummary summary, "TS180-01Z-003", "四孔平板", 3, "立柱平台#20: 3"
    
    ' 检查汇总结果
    If summary.Exists("TS180-01Z-003") Then
        Dim item As Object: Set item = summary("TS180-01Z-003")
        Debug.Print "零件总数: " & item("TotalQty") & " (预期: 4)"
        Debug.Print "计算过程: " & item("Breakdown")
    End If
End Sub

' 创建模拟表格对象（用于测试）
Private Function CreateMockTable() As Object
    Set CreateMockTable = New MockTableClass
End Function

' 模拟表格类（简化版本，仅用于测试）
Private Class MockTableClass
    Public Function Text(row As Long, col As Long) As String
        ' 模拟表头行  
        If row = 0 Then
            Select Case col
                Case 0: Text = "DOCUMENT PREVIEW"
                Case 1: Text = "PART NUMBER" 
                Case 2: Text = "数量"
                Case 3: Text = "代号"
                Case 4: Text = "名称"
                Case 5: Text = "PART NAME"
                Case 6: Text = "是否组装"
                Case Else: Text = "其他列"
            End Select
        Else
            ' 模拟数据行
            Text = "测试数据" & row & "-" & col
        End If
    End Function
    
    Public Property Get ColumnCount() As Long
        ColumnCount = 7
    End Property
    
    Public Property Get RowCount() As Long  
        RowCount = 5
    End Property
End Class

' 模拟的AddToSummary函数（复制自Utils.bas以便测试）
Private Sub AddToSummary(ByRef summary As Object, ByVal partNo As String, ByVal partName As String, ByVal qty As Long, ByVal chain As String)
    Dim key As String: key = partNo
    Dim item As Object
    If summary.Exists(key) Then
        Set item = summary(key)
        item("TotalQty") = CLng(item("TotalQty")) + qty
        item("Breakdown") = CStr(item("Breakdown")) & " + " & chain & " => " & item("TotalQty")
    Else
        Set item = CreateObject("Scripting.Dictionary")
        item.Add "PartNo", partNo
        item.Add "PartName", partName  
        item.Add "TotalQty", qty
        item.Add "Breakdown", chain & " => " & qty
        summary.Add key, item
    End If
End Sub

' 模拟的FindColumnIndex函数（复制自RecursiveProcessor.bas以便测试）
Private Function FindColumnIndex(ta As Object, names As Variant) As Long
    Dim c As Long
    For c = 0 To ta.ColumnCount - 1
        Dim title As String: title = UCase$(Trim$(ta.Text(0, c)))
        Dim i As Long
        For i = LBound(names) To UBound(names)
            If InStr(title, UCase$(names(i))) > 0 Then
                FindColumnIndex = c
                Exit Function
            End If
        Next
    Next
    FindColumnIndex = -1
End Function

' 模拟的IsAssembleCell函数（复制自RecursiveProcessor.bas以便测试）
Private Function IsAssembleCell(valText As String) As Boolean
    Dim t As String: t = UCase$(Trim$(valText))
    IsAssembleCell = (t = "是" Or t = "Y" Or t = "YES" Or t = "TRUE" Or t = "1")
End Function

' 从Utils.bas复制的函数用于测试
Private Function GetFileNameNoExt(path As String) As String
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

Private Function GetFileFolder(path As String) As String
    Dim i As Long
    For i = Len(path) To 1 Step -1
        If Mid$(path, i, 1) = "\" Or Mid$(path, i, 1) = "/" Then
            GetFileFolder = Left$(path, i - 1)
            Exit Function
        End If
    Next
    GetFileFolder = CurDir$()
End Function

Private Function HtmlEncode(ByVal s As String) As String
    s = Replace(s, "&", "&amp;")
    s = Replace(s, "<", "&lt;")
    s = Replace(s, ">", "&gt;")
    s = Replace(s, "\"", "&quot;")
    HtmlEncode = s
End Function