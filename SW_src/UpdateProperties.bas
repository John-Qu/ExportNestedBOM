'========================================================
' SolidWorks 2019 批量添加/更新自定义属性（带可编辑确认界面）
' 新增属性：设计（OS用户名）、定型日期（文件创建日期）、
'          型号、SUPPLIER（若无则空行由你填写）
' 界面：每项可编辑 + 勾选导入，默认全选
'========================================================
Option Explicit

' ExportNestedBOM - SolidWorks 侧属性批量工具
' 入口：Run_AddCustomProps（窗体驱动），支持处理当前打开文档或选择文件夹批量。
' 行为概述：
' - 读取已有自定义属性（型号、SUPPLIER/渠道、处理、是否组装/外购/机加/钣金等），不覆盖非空值。
' - 缺省回退：若“是否钣金”未设置，按几何判断 IsSheetMetal() 回填“是/否”。
' - 写入策略：优先文档级（Custom），可选配置级（当前配置），Delete+Add2 文本类型。
' - 记录：通过 Logger 输出关键步骤与统计。

' SolidWorks doc type constants
Private Const swDocPART As Long = 1
Private Const swDocASSEMBLY As Long = 2

Private Const swOpenDocOptions_Silent As Long = 1
' 属性类型：文本
Private Const swCustomInfoText As Long = 30

Dim swApp As Object ' SldWorks.SldWorks

Public Sub Run_AddCustomProps()
    On Error Resume Next
    Set swApp = Application.SldWorks
    On Error GoTo 0
    If swApp Is Nothing Then
        MsgBox "无法获取 SolidWorks 应用对象。请从 SolidWorks 内部运行此宏。", vbCritical
        Exit Sub
    End If
    
    Dim ans As VbMsgBoxResult
    ans = MsgBox("是否处理当前打开的文件？" & vbCrLf & _
                 "是 = 处理当前文件" & vbCrLf & _
                 "否 = 选择文件夹批量处理", vbQuestion + vbYesNoCancel, "批量添加/更新自定义属性")
    If ans = vbCancel Then Exit Sub
    
    If ans = vbYes Then
        ProcessActiveDoc
    Else
        ProcessFolderBatch
    End If
End Sub

Private Sub ProcessActiveDoc()
    Dim swModel As Object
    Set swModel = swApp.ActiveDoc
    If swModel Is Nothing Then
        MsgBox "没有打开的模型文件。", vbExclamation
        Exit Sub
    End If

    ' 新增：仅按名称/内容匹配删除相关方程式 + 删除指定属性
    PreClean_RemoveEquationsAndProps swModel

    Dim names As Variant, values As Variant, fileInfo As String
    BuildPropArrays swModel, names, values, fileInfo
    
    Dim dlg As frmProps
    Set dlg = VBA.UserForms.Add("frmProps")
    dlg.InitWithData fileInfo, names, values
    dlg.Show vbModal
    
    If dlg.UserAction = 1 Then ' OK
        WriteSelectedProps swModel, dlg.SelectedNames, dlg.SelectedValues, dlg.SelectedChecks
        MsgBox "已写入当前文件的自定义属性。", vbInformation
    ElseIf dlg.UserAction = 3 Then
        ' Cancel - do nothing
    End If
    Unload dlg
End Sub

Private Sub ProcessFolderBatch()
    Dim folder As String
    folder = BrowseForFolder("选择包含模型文件的文件夹（仅处理顶层 *.sldprt; *.sldasm）")
    If Len(folder) = 0 Then Exit Sub
    
    Dim exts As Variant: exts = Array("*.sldprt", "*.sldasm")
    Dim processed As Long, updated As Long, skipped As Long
    Dim cancelAll As Boolean: cancelAll = False
    
    Dim i As Long, f As String
    For i = LBound(exts) To UBound(exts)
        f = Dir(AddTrailingSlash(folder) & exts(i))
        Do While Len(f) > 0
            If cancelAll Then Exit Do
            
            Dim fullPath As String
            fullPath = AddTrailingSlash(folder) & f
            
            Dim docType As Long
            docType = IIf(LCase$(Right$(f, 7)) = "sldasm", swDocASSEMBLY, swDocPART)
            
            Dim swModel As Object
            Set swModel = swApp.GetOpenDocumentByName(fullPath)
            
            Dim newlyOpened As Boolean: newlyOpened = False
            If swModel Is Nothing Then
                On Error Resume Next
                Set swModel = swApp.OpenDoc6(fullPath, docType, swOpenDocOptions_Silent, "", 0, 0)
                On Error GoTo 0
                newlyOpened = Not swModel Is Nothing
            End If
            
            If Not swModel Is Nothing Then
                processed = processed + 1
                
                ' 新增：预清理（仅删除命中关键词的方程式 + 删除指定属性）
                PreClean_RemoveEquationsAndProps swModel

                Dim names As Variant, values As Variant, fileInfo As String
                BuildPropArrays swModel, names, values, fileInfo
                
                Dim dlg As frmProps
                Set dlg = VBA.UserForms.Add("frmProps")
                dlg.InitWithData fileInfo, names, values
                dlg.Caption = "属性确认 - " & f
                dlg.Show vbModal
                
                If dlg.UserAction = 1 Then
                    WriteSelectedProps swModel, dlg.SelectedNames, dlg.SelectedValues, dlg.SelectedChecks
                    ' 保存
                    On Error Resume Next
                    Dim errNum As Long, warnNum As Long
                    swModel.Save3 1, errNum, warnNum
                    On Error GoTo 0
                    updated = updated + 1
                ElseIf dlg.UserAction = 2 Then
                    skipped = skipped + 1
                ElseIf dlg.UserAction = 3 Then
                    cancelAll = True
                End If
                
                Unload dlg
                
                If newlyOpened Then
                    On Error Resume Next
                    swApp.CloseDoc fullPath
                    On Error GoTo 0
                End If
            End If
            
            f = Dir
        Loop
        If cancelAll Then Exit For
    Next i
    
    Dim summary As String
    summary = "处理完成：" & vbCrLf & _
              "总计文件: " & processed & vbCrLf & _
              "已更新: " & updated & vbCrLf & _
              "已跳过: " & skipped & IIf(cancelAll, vbCrLf & "用户中止。", "")
    MsgBox summary, vbInformation
End Sub

' 构建属性名/值数组，以及文件信息字符串
Private Sub BuildPropArrays(ByVal swModel As Object, ByRef names As Variant, ByRef values As Variant, ByRef fileInfo As String)
    Dim title As String: title = SafeStr(swModel.GetTitle)
    Dim path As String: path = SafeStr(swModel.GetPathName)

    Dim fileBase As String, codeStr As String, nameStr As String
    fileBase = BaseNameNoExt(IIf(Len(path) > 0, path, title))
    ParseCodeAndName fileBase, codeStr, nameStr

    Dim folderName As String, projCode As String, projName As String
    folderName = ParentFolderName(path)
    ParseProjectFromFolder folderName, projCode, projName

    Dim massKg As String: massKg = GetMassKgString(swModel)

    Dim createDate As String: createDate = GetInternalCreationDate(swModel)

    ' 读取现有自定义属性（若无则留空）
    Dim modelNo As String: modelNo = GetCustomPropValue(swModel, "型号")
    Dim supplier As String: supplier = GetCustomPropValue(swModel, "SUPPLIER")
    Dim procStr As String: procStr = GetCustomPropValue(swModel, "处理")
    Dim isBuy As String: isBuy = GetCustomPropValue(swModel, "是否采购")
    Dim isAssm As String: isAssm = GetCustomPropValue(swModel, "是否装配")
    Dim isMach As String: isMach = GetCustomPropValue(swModel, "是否机加")
    Dim existingIsSheet As String: existingIsSheet = GetCustomPropValue(swModel, "是否钣金")
    Dim memo As String: memo = GetCustomPropValue(swModel, "备注")
    Dim designer As String: designer = GetCustomPropValue(swModel, "设计")
    If Len(designer) = 0 Then designer = Environ$("USERNAME")

    ' “是否钣金”属性不存在时，回退用几何判断
    Dim isSheetFlag As String
    If Len(existingIsSheet) > 0 Then
        isSheetFlag = existingIsSheet
    Else
        isSheetFlag = IIf(IsSheetMetal(swModel), "是", "否")
    End If
    
    ' 汇总为数组（顺序即界面行顺序）
    Dim n() As String, v() As String
    Dim idx As Long: idx = -1
    
    AddNV n, v, idx, "项目代号", projCode
    AddNV n, v, idx, "项目名称", projName
    AddNV n, v, idx, "代号", codeStr
    AddNV n, v, idx, "名称", nameStr
    AddNV n, v, idx, "SUPPLIER", supplier
    AddNV n, v, idx, "型号", modelNo
    AddNV n, v, idx, "质量", massKg
    AddNV n, v, idx, "处理", procStr
    AddNV n, v, idx, "是否装配", isAssm
    AddNV n, v, idx, "是否采购", isBuy
    AddNV n, v, idx, "是否钣金", isSheetFlag
    AddNV n, v, idx, "是否机加", isMach
    AddNV n, v, idx, "设计", designer
    AddNV n, v, idx, "定型日期", createDate
    AddNV n, v, idx, "备注", memo

    names = n
    values = v
    
    fileInfo = "文件: " & title & vbCrLf & _
               "路径: " & path & vbCrLf & _
               "提示：在下方直接编辑值，并勾选需要导入的属性。"
End Sub

' 同时写入：文档级(自定义) 和 当前配置(配置特定)，采用“先删后增”的强力模式
Private Sub WriteSelectedProps(ByVal swModel As Object, ByVal names As Variant, ByVal values As Variant, ByVal checks As Variant)
    On Error Resume Next
    Dim cfgName As String
    cfgName = swModel.ConfigurationManager.ActiveConfiguration.name
    
    Dim cpmDoc As Object, cpmCfg As Object
    Set cpmDoc = swModel.Extension.CustomPropertyManager("")       ' 自定义（文档级）
    Set cpmCfg = swModel.Extension.CustomPropertyManager(cfgName)  ' 配置特定（当前配置）
    
    Dim i As Long
    For i = LBound(names) To UBound(names)
        If checks(i) Then
            ' 对“自定义”和“配置特定”都执行强力写入
            WriteProp_DeleteAndAdd cpmDoc, CStr(names(i)), CStr(values(i))
            WriteProp_DeleteAndAdd cpmCfg, CStr(names(i)), CStr(values(i))
        End If
    Next i
    On Error GoTo 0
End Sub

' 【核心修正】强力写入：先删除，再添加。这能解决新增属性失败的问题。
Private Sub WriteProp_DeleteAndAdd(ByVal cpm As Object, ByVal propName As String, ByVal propValue As String)
    If cpm Is Nothing Then Exit Sub
    On Error Resume Next
    ' 1. 先尝试删除，清除任何可能存在的残留状态
    cpm.Delete propName
    ' 2. 再用 Add2 新建为“文本”类型。这是最可靠的新增方法。
    '    swCustomInfoText 就是我们之前定义的常量 30
    cpm.Add2 propName, swCustomInfoText, propValue
    On Error GoTo 0
End Sub

' ===== 工具函数 =====

Private Sub AddNV(ByRef n() As String, ByRef v() As String, ByRef idx As Long, ByVal name As String, ByVal value As String)
    idx = idx + 1
    ReDim Preserve n(0 To idx)
    ReDim Preserve v(0 To idx)
    n(idx) = name
    v(idx) = value
End Sub

Private Sub ParseCodeAndName(ByVal baseName As String, ByRef codeStr As String, ByRef nameStr As String)
    Dim s As String: s = Replace(baseName, "　", " ")
    Dim p As Long: p = InStr(1, s, " ")
    If p > 0 Then
        codeStr = Trim$(Left$(s, p - 1))
        nameStr = Trim$(Mid$(s, p + 1))
    Else
        codeStr = s
        nameStr = ""
    End If
End Sub

Private Sub ParseProjectFromFolder(ByVal folderName As String, ByRef projCode As String, ByRef projName As String)
    projCode = "": projName = ""
    If Len(folderName) = 0 Then Exit Sub
    Dim p As Long: p = InStr(1, folderName, "_")
    If p > 0 Then
        projCode = Trim$(Left$(folderName, p - 1))
        projName = Trim$(Mid$(folderName, p + 1))
    Else
        projCode = folderName
        projName = ""
    End If
End Sub

Private Function GetMassKgString(ByVal swModel As Object) As String
    On Error Resume Next
    swModel.ForceRebuild3 False
    Dim mp As Object
    Set mp = swModel.Extension.CreateMassProperty
    Dim m As Double: m = 0#
    If Not mp Is Nothing Then m = mp.Mass
    If m < 0.0000001 Then
        GetMassKgString = ""
    Else
        GetMassKgString = FormatNumber(m, 3, vbTrue, vbFalse, vbFalse)
    End If
    On Error GoTo 0
End Function

Private Function IsSheetMetal(ByVal swModel As Object) As Boolean
    On Error Resume Next
    If swModel.GetType <> swDocPART Then
        IsSheetMetal = False
        Exit Function
    End If
    Dim part As Object
    Set part = swModel
    If Not part Is Nothing Then
        IsSheetMetal = part.IsSheetMetal
    Else
        IsSheetMetal = False
    End If
    On Error GoTo 0
End Function
' 【改进】读取自定义属性：优先从“配置特定”读取，若无则回退到“自定义”
Private Function GetCustomPropValue(ByVal swModel As Object, ByVal propName As String) As String
    On Error Resume Next
    Dim valOut As String: valOut = ""
    Dim rawVal As String, resVal As String
    Dim wasResolved As Boolean, linkToText As Boolean
    Dim ret As Long
    
    ' 1. 优先尝试从“配置特定”属性中获取
    Dim cfgName As String: cfgName = swModel.ConfigurationManager.ActiveConfiguration.name
    Dim cpmCfg As Object: Set cpmCfg = swModel.Extension.CustomPropertyManager(cfgName)
    If Not cpmCfg Is Nothing Then
        ret = cpmCfg.Get6(propName, False, rawVal, resVal, wasResolved, linkToText)
        If ret = 2 Then valOut = IIf(Len(resVal) > 0, resVal, rawVal)
    End If
    
    ' 2. 如果在“配置特定”中没找到，则回退到“自定义”（文档级）属性
    If Len(valOut) = 0 Then
        Dim cpmDoc As Object: Set cpmDoc = swModel.Extension.CustomPropertyManager("")
        If Not cpmDoc Is Nothing Then
            ret = cpmDoc.Get6(propName, False, rawVal, resVal, wasResolved, linkToText)
            If ret = 2 Then valOut = IIf(Len(resVal) > 0, resVal, rawVal)
        End If
    End If
    
    GetCustomPropValue = SafeStr(valOut)
    On Error GoTo 0
End Function

' 读取 SolidWorks 文件内部摘要信息中的创建日期
Private Function GetInternalCreationDate(ByVal swModel As Object) As String
    On Error Resume Next
    Dim createDateStr As String
    createDateStr = swModel.SummaryInfo(swSumInfoCreateDate)
    If Len(createDateStr) > 0 Then
        ' 对API返回的日期字符串进行解析和重新格式化，确保格式统一
        GetInternalCreationDate = Format$(CDate(createDateStr), "yyyy-mm-dd")
    Else
        GetInternalCreationDate = ""
    End If
    On Error GoTo 0
End Function

Private Function SafeStr(ByVal v As Variant) As String
    If IsNull(v) Or IsEmpty(v) Then
        SafeStr = ""
    Else
        SafeStr = CStr(v)
    End If
End Function

Private Function BaseNameNoExt(ByVal fullPathOrTitle As String) As String
    Dim s As String: s = fullPathOrTitle
    Dim p As Long: p = InStrRev(s, "\")
    If p > 0 Then s = Mid$(s, p + 1)
    p = InStrRev(s, ".")
    If p > 0 Then
        BaseNameNoExt = Left$(s, p - 1)
    Else
        BaseNameNoExt = s
    End If
End Function

Private Function ParentFolderName(ByVal fullPath As String) As String
    On Error Resume Next
    If Len(fullPath) = 0 Then Exit Function
    Dim parentPath As String
    Dim p As Long
    p = InStrRev(fullPath, "\")
    If p = 0 Then Exit Function
    parentPath = Left$(fullPath, p - 1)
    p = InStrRev(parentPath, "\")
    If p = 0 Then
        ParentFolderName = parentPath
    Else
        ParentFolderName = Mid$(parentPath, p + 1)
    End If
    On Error GoTo 0
End Function

Private Function BrowseForFolder(ByVal prompt As String) As String
    On Error Resume Next
    Dim sh As Object, fol As Object
    Set sh = CreateObject("Shell.Application")
    Set fol = sh.BrowseForFolder(0, prompt, 0)
    If Not fol Is Nothing Then
        Dim p As String
        p = fol.Self.path
        If Len(p) > 0 Then BrowseForFolder = p
    End If
    On Error GoTo 0
End Function

Private Function AddTrailingSlash(ByVal p As String) As String
    If Len(p) = 0 Then
        AddTrailingSlash = ""
    ElseIf Right$(p, 1) = "\" Then
        AddTrailingSlash = p
    Else
        AddTrailingSlash = p & "\"
    End If
End Function


' 仅删除名称/内容包含目标关键词的方程式；并删除对应属性（文档级 + 所有配置）
Private Sub PreClean_RemoveEquationsAndProps(ByVal swModel As Object)
    On Error Resume Next

    Dim targets As Variant
    targets = Array("项目代号代码", "名称代码", "代号代码")

    ' 1) 过滤删除方程式（EquationMgr）
    Dim eqMgr As Object
    Set eqMgr = swModel.GetEquationMgr
    If Not eqMgr Is Nothing Then
        Dim eqCount As Long
        eqCount = eqMgr.GetCount
        If eqCount > 0 Then
            Dim i As Long
            For i = eqCount - 1 To 0 Step -1
                Dim eqText As String
                eqText = CStr(eqMgr.Equation(i)) ' 形如："\"D1@Sketch1\" = 5mm" 或含表达式文本
                If ContainsAny(eqText, targets) Then
                    eqMgr.Delete i
                End If
            Next i
        End If
    End If

    ' 2) 删除指定的自定义属性（文档级）
    Dim cpmDoc As Object
    Set cpmDoc = swModel.Extension.CustomPropertyManager("")
    If Not cpmDoc Is Nothing Then
        Dim k As Long
        For k = LBound(targets) To UBound(targets)
            cpmDoc.Delete CStr(targets(k))
        Next k
    End If

    ' 3) 删除指定的自定义属性（所有配置）
    Dim cfgMgr As Object
    Set cfgMgr = swModel.ConfigurationManager
    If Not cfgMgr Is Nothing Then
        Dim vCfgNames As Variant
        vCfgNames = cfgMgr.GetConfigurationNames
        If Not IsEmpty(vCfgNames) Then
            Dim idx As Long, cfgName As String
            For idx = LBound(vCfgNames) To UBound(vCfgNames)
                cfgName = CStr(vCfgNames(idx))
                Dim cpmCfg As Object
                Set cpmCfg = swModel.Extension.CustomPropertyManager(cfgName)
                If Not cpmCfg Is Nothing Then
                    For k = LBound(targets) To UBound(targets)
                        cpmCfg.Delete CStr(targets(k))
                    Next k
                End If
            Next idx
        End If
    End If

    On Error GoTo 0
End Sub

' 文本包含任一目标关键字则返回 True（不区分大小写）
Private Function ContainsAny(ByVal text As String, ByVal targets As Variant) As Boolean
    Dim t As Variant
    For Each t In targets
        If Len(text) > 0 And InStr(1, text, CStr(t), vbTextCompare) > 0 Then
            ContainsAny = True
            Exit Function
        End If
    Next t
    ContainsAny = False
End Function