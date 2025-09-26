Option Explicit

Public Sub T1_Run_ToolboxReplace_ActiveSheet()
    Dim wb As Workbook: Set wb = Utils.ResolveTargetWorkbook()
    If wb Is Nothing Then
        MsgBox "未找到可处理的目标工作簿（请打开要处理的 .xls 文件）", vbExclamation
        Exit Sub
    End If

    Dim dirPath As String: dirPath = Utils.WorkbookDir(wb)
    Logger.LogInit dirPath, "T1"
    Logger.LogInfo "Target Workbook=" & wb.Name & ", Dir=" & dirPath

    Dim mapping As Object
    Set mapping = Utils.LoadToolboxMapping(dirPath)
    Logger.LogInfo "Mapping entries=" & mapping.Count

    Dim replaced As Long, unmatched As Long
    On Error GoTo FIN
    ' 激活目标工作簿，以便获取其活动工作表
    wb.Activate
    Dim ws As Worksheet: Set ws = ActiveSheet
    Logger.LogInfo "Processing ActiveSheet=" & ws.Name
    ' 先执行列标题重命名与列顺序调整（T2）
    SingleSheetFormatter.RenameHeadersAndReorder ws
    ' 再执行用例T1的替换
    SingleSheetFormatter.ApplyToolboxNameReplacement ws, mapping, replaced, unmatched
    ' 执行用例T3：布尔图标化
    SingleSheetFormatter.IconizeBooleanFlags ws
    Logger.LogInfo "DONE: replaced=" & replaced & ", unmatched=" & unmatched & "; Log=" & Logger.LogPath
FIN:
    Logger.LogClose
End Sub

Public Sub T1_Run_ToolboxReplace_AllVisibleSheets()
    Dim wb As Workbook: Set wb = Utils.ResolveTargetWorkbook()
    If wb Is Nothing Then
        MsgBox "未找到可处理的目标工作簿（请打开要处理的 .xls 文件）", vbExclamation
        Exit Sub
    End If

    Dim dirPath As String: dirPath = Utils.WorkbookDir(wb)
    Logger.LogInit dirPath, "T1"
    Logger.LogInfo "Target Workbook=" & wb.Name & ", Dir=" & dirPath

    Dim ws As Worksheet
    For Each ws In wb.Worksheets
        If ws.Visible = xlSheetVisible Then
            On Error Resume Next
            SingleSheetFormatter.RenameHeadersAndReorder ws
            Dim mapping As Object: Set mapping = Utils.LoadToolboxMapping(dirPath)
            Dim replaced As Long, unmatched As Long
            SingleSheetFormatter.ApplyToolboxNameReplacement ws, mapping, replaced, unmatched
            SingleSheetFormatter.IconizeBooleanFlags ws
            On Error GoTo 0
        End If
    Next ws
    Logger.LogInfo "DONE: all visible sheets processed; Log=" & Logger.LogPath
    Logger.LogClose
End Sub

' 新增：运行格式与打印（S2/S3/S6/S7）
Public Sub Run_Format_CurrentSheet()
    Dim wb As Workbook: Set wb = Utils.ResolveTargetWorkbook()
    If wb Is Nothing Then
        MsgBox "未找到可处理的目标工作簿（请打开要处理的 .xls 文件）", vbExclamation
        Exit Sub
    End If

    Dim dirPath As String: dirPath = Utils.WorkbookDir(wb)
    Logger.LogInit dirPath, "T4"
    On Error GoTo FIN
    wb.Activate
    Dim ws As Worksheet: Set ws = ActiveSheet
    Logger.LogInfo "Formatting ActiveSheet=" & ws.Name
    SingleSheetFormatter.FormatSingleBOMSheet ws
    Logger.LogInfo "DONE: formatted sheet=" & ws.Name
FIN:
    Logger.LogClose
End Sub

Public Sub Run_Format_AllVisibleSheets()
    Dim wb As Workbook: Set wb = Utils.ResolveTargetWorkbook()
    If wb Is Nothing Then
        MsgBox "未找到可处理的目标工作簿（请打开要处理的 .xls 文件）", vbExclamation
        Exit Sub
    End If

    Dim dirPath As String: dirPath = Utils.WorkbookDir(wb)
    Logger.LogInit dirPath, "T4"
    On Error GoTo FIN
    Dim ws As Worksheet
    For Each ws In wb.Worksheets
        If ws.Visible = xlSheetVisible Then
            Logger.LogInfo "Formatting sheet=" & ws.Name
            SingleSheetFormatter.FormatSingleBOMSheet ws
        End If
    Next ws
    Logger.LogInfo "DONE: formatted all visible sheets"
FIN:
    Logger.LogClose
End Sub

' 新增：格式化并导出当前工作簿所有可见工作表为 PDF（T5/E4）
Public Sub Run_FormatAndExport_CurrentWorkbook()
    Dim wb As Workbook: Set wb = Utils.ResolveTargetWorkbook()
    If wb Is Nothing Then
        MsgBox "未找到可处理的目标工作簿（请打开要处理的 .xls 文件）", vbExclamation
        Exit Sub
    End If

    Dim dirPath As String: dirPath = Utils.WorkbookDir(wb)
    Logger.LogInit dirPath, "T5"
    On Error GoTo FIN
    Dim ws As Worksheet
    Dim pdfPath As String
    For Each ws In wb.Worksheets
        If ws.Visible = xlSheetVisible Then
            Logger.LogInfo "Formatting+Exporting sheet=" & ws.Name
            SingleSheetFormatter.FormatSingleBOMSheet ws
            pdfPath = PdfExport.ExportWorksheetToPdf(ws)
            If Len(pdfPath) > 0 Then
                Logger.LogInfo "PDF generated: " & pdfPath
            End If
        End If
    Next ws
    Logger.LogInfo "DONE: All visible sheets formatted and exported to PDF/"
FIN:
    Logger.LogClose
End Sub

Public Sub Run_Generate_TotalBOM_FromSummary()
    Dim wb As Workbook: Set wb = Utils.ResolveTargetWorkbook()
    If wb Is Nothing Then
        MsgBox "未找到目标工作簿（请激活 *_汇总.xls 工作簿）", vbExclamation
        Exit Sub
    End If

    Dim baseDir As String: baseDir = Utils.WorkbookDir(wb)
    Logger.LogInit baseDir, "T6"

    On Error GoTo FAIL

    SummaryProcessor.BuildTotalBOMFromSummary
    ' 生成后，统一对总表图片执行 50% 缩放（解锁纵横比）
    ' SummaryProcessor.ScaleAllPicturesInTotalBOMTo50
    ' 继续按 T7：生成分类表并导出 PDF
    ' SummaryProcessor.BuildCategorySheetsFromTotalBOM
    
    Logger.LogClose
    Exit Sub
FAIL:
    Logger.LogError "Run_Generate_TotalBOM_FromSummary failed: " & Err.Description
    Logger.LogClose
End Sub

Public Sub Run_Merge_SubBOMs_Into_CurrentWorkbook()
    On Error GoTo EH
    Dim topWb As Workbook: Set topWb = Utils.ResolveTargetWorkbook()
    If topWb Is Nothing Then
        MsgBox "未找到目标工作簿（请打开顶层装配的 .xls 工作簿）", vbExclamation
        Exit Sub
    End If

    Dim baseDir As String: baseDir = Utils.WorkbookDir(topWb)
    Logger.LogInit baseDir, "T8"
    Logger.LogInfo "Merge: top workbook=" & topWb.Name & ", baseDir=" & baseDir

    Dim merged As Long: merged = MergeSubBOMsIntoWorkbook(baseDir, topWb)
    Logger.LogInfo "Merge: DONE. merged sheets=" & merged & "; Log=" & Logger.LogPath
    Logger.LogClose
    Exit Sub
EH:
    Logger.LogError "Run_Merge_SubBOMs_Into_CurrentWorkbook failed: " & Err.Description
    Logger.LogClose
End Sub

Private Function MergeSubBOMsIntoWorkbook(ByVal baseDir As String, ByVal topWb As Workbook) As Long
    On Error GoTo EH
    Dim f As String
    Dim cnt As Long: cnt = 0
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim wasAlreadyOpen As Boolean
    Dim topWasShared As Boolean

    If Len(baseDir) = 0 Then
        Logger.LogWarn "Merge: current workbook has no Path; skip"
        MergeSubBOMsIntoWorkbook = 0
        Exit Function
    End If

    ' 目标工作簿解除共享并保存（如需）
    topWasShared = topWb.MultiUserEditing
    EnsureWorkbookExclusive topWb
    If topWasShared And Not topWb.MultiUserEditing Then
        On Error Resume Next
        topWb.Save
        On Error GoTo EH
        Logger.LogInfo "Merge: target workbook unshared and saved: " & topWb.Name
    End If

    ' 将顶层工作簿中的 "Sheet1" 重命名为 顶层装配体清单的文件名（去扩展名）
    Dim desired As String, targetSheetName As String
    desired = topWb.Name
    If InStrRev(desired, ".") > 0 Then desired = Left$(desired, InStrRev(desired, ".") - 1)
    targetSheetName = MakeUniqueSheetName(topWb, SafeSheetName(desired))
    Dim wsTop As Worksheet, foundSheet As Worksheet
    Set foundSheet = Nothing
    For Each wsTop In topWb.Worksheets
        If LCase$(Trim$(wsTop.Name)) = "sheet1" Then
            Set foundSheet = wsTop
            Exit For
        End If
    Next
    If Not foundSheet Is Nothing Then
        If StrComp(foundSheet.Name, targetSheetName, vbTextCompare) <> 0 Then
            On Error Resume Next
            foundSheet.Name = targetSheetName
            If Err.Number = 0 Then
                Logger.LogInfo "Rename: 'Sheet1' -> '" & targetSheetName & "'"
            Else
                Logger.LogWarn "Rename: failed to rename 'Sheet1' to '" & targetSheetName & "': " & Err.Description
                Err.Clear
            End If
            On Error GoTo EH
        End If
    End If
    Dim scanPattern As String
        scanPattern = baseDir & Application.PathSeparator & "*.xls*"
        Logger.LogInfo "Merge: scanning pattern=" & scanPattern
        f = Dir(scanPattern)
     Do While f <> ""
         ' 排除当前工作簿、汇总文件、映射宏工作簿与 Excel 临时锁文件
         If StrComp(f, topWb.Name, vbTextCompare) <> 0 _
            And InStr(1, f, "_汇总", vbTextCompare) = 0 _
            And StrComp(f, CFG_MAPPING_WORKBOOK_NAME, vbTextCompare) <> 0 _
            And Left$(f, 2) <> "~$" Then
 
             ' 打开或复用已打开的工作簿
             Set wb = Utils.FindOpenWorkbookByName(f)
             wasAlreadyOpen = Not (wb Is Nothing)
             If Not wasAlreadyOpen Then
                    Set wb = Application.Workbooks.Open(FileName:=baseDir & Application.PathSeparator & f, ReadOnly:=False)
             End If

            ' 源工作簿解除共享（如需）
            Dim srcWasShared As Boolean: srcWasShared = wb.MultiUserEditing
            EnsureWorkbookExclusive wb
            If srcWasShared And Not wb.MultiUserEditing Then
                Logger.LogInfo "Merge: source workbook unshared: " & wb.Name
            End If

            ' 选择第一个可见工作表作为拷贝源
            Set ws = Nothing
            Dim tmpWS As Worksheet
            For Each tmpWS In wb.Worksheets
                If tmpWS.Visible = xlSheetVisible Then
                    Set ws = tmpWS
                    Exit For
                End If
            Next

            If Not ws Is Nothing Then
                Dim nameNoExt As String: nameNoExt = f
                If InStrRev(nameNoExt, ".") > 0 Then nameNoExt = Left$(nameNoExt, InStrRev(nameNoExt, ".") - 1)

                Dim sheetName As String
                sheetName = MakeUniqueSheetName(topWb, SafeSheetName(nameNoExt))

                ' 拷贝工作表到当前工作簿尾部，并重命名
                ws.Copy After:=topWb.Worksheets(topWb.Worksheets.Count)
                topWb.Worksheets(topWb.Worksheets.Count).Name = sheetName
                Logger.LogInfo "Merge: copied sheet from '" & wb.Name & "' as '" & sheetName & "'"
                cnt = cnt + 1
            Else
                Logger.LogWarn "Merge: no visible sheet in workbook='" & wb.Name & "'"
            End If

            ' 关闭我们临时打开的工作簿（如解除共享则保存以持久化）
            If Not wasAlreadyOpen Then
                Application.DisplayAlerts = False
                wb.Close SaveChanges:=(srcWasShared And Not wb.MultiUserEditing)
                Application.DisplayAlerts = True
            End If
        End If
        f = Dir()
    Loop

    ' 应用页眉页脚设定到所有工作表
    'ApplyHeaderFooterForAllSheets topWb, baseDir

    MergeSubBOMsIntoWorkbook = cnt
    Exit Function
EH:
    Logger.LogError "MergeSubBOMsIntoWorkbook failed: " & Err.Description
    MergeSubBOMsIntoWorkbook = cnt
End Function

Private Function SafeSheetName(ByVal nm As String) As String
    Dim s As String: s = Trim$(nm)
    ' 替换非法字符 ： / \ ? * [ ]
    s = Replace$(s, ":", "_")
    s = Replace$(s, "/", "_")
    s = Replace$(s, "\", "_")
    s = Replace$(s, "?", "_")
    s = Replace$(s, "*", "_")
    s = Replace$(s, "[", "_")
    s = Replace$(s, "]", "_")
    If Len(s) = 0 Then s = "Sheet"
    If Len(s) > 31 Then s = Left$(s, 31)
    SafeSheetName = s
End Function

Private Function MakeUniqueSheetName(ByVal wb As Workbook, ByVal baseName As String) As String
    Dim candidate As String: candidate = baseName
    Dim suffix As Long: suffix = 2
    Do While SheetExists(wb, candidate)
        Dim sfx As String: sfx = " (" & suffix & ")"
        Dim maxBase As Long: maxBase = 31 - Len(sfx)
        If maxBase < 1 Then maxBase = 1
        candidate = Left$(baseName, maxBase) & sfx
        suffix = suffix + 1
    Loop
    MakeUniqueSheetName = candidate
End Function

Private Function SheetExists(ByVal wb As Workbook, ByVal name As String) As Boolean
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = wb.Worksheets(name)
    SheetExists = Not (ws Is Nothing)
    On Error GoTo 0
End Function

Private Sub EnsureWorkbookExclusive(ByVal wb As Workbook)
    On Error Resume Next
    If wb.MultiUserEditing Then
        Application.DisplayAlerts = False
        wb.ExclusiveAccess
        Application.DisplayAlerts = True
        If Not wb.MultiUserEditing Then
            Logger.LogInfo "ExclusiveAccess: unshared '" & wb.Name & "'"
        Else
            Logger.LogWarn "ExclusiveAccess: still shared '" & wb.Name & "'"
        End If
    End If
    On Error GoTo 0
End Sub

' 为所有工作表设置指定的页眉页脚
Private Sub ApplyHeaderFooterForAllSheets(ByVal wb As Workbook, ByVal baseDir As String)
    On Error Resume Next
    Dim parentName As String: parentName = GetParentFolderName(baseDir)
    Dim bookNameNoExt As String: bookNameNoExt = wb.Name
    If InStrRev(bookNameNoExt, ".") > 0 Then bookNameNoExt = Left$(bookNameNoExt, InStrRev(bookNameNoExt, ".") - 1)
    Dim osUser As String: osUser = GetOSUserName()

    Dim ws As Worksheet
    For Each ws In wb.Worksheets
        Dim totalPages As Long: totalPages = GetTotalPagesForSheetExact(ws)
        With ws.PageSetup
            .LeftHeader = parentName
            .CenterHeader = bookNameNoExt
            .RightHeader = "&A"
            .LeftFooter = "&D &T"
            .CenterFooter = "第 &P 页，共 " & CStr(totalPages) & " 页"
            .RightFooter = osUser
        End With
        Logger.LogInfo "HeaderFooter: applied to sheet='" & ws.Name & "', totalPages=" & totalPages
    Next ws
    On Error GoTo 0
End Sub

Private Function GetTotalPagesForSheetExact(ByVal ws As Worksheet) As Long
    On Error Resume Next
    Dim prevSheet As Worksheet
    Set prevSheet = ws.Parent.ActiveSheet

    Dim prevEvents As Boolean: prevEvents = Application.EnableEvents
    Dim prevUpdating As Boolean: prevUpdating = Application.ScreenUpdating
    Application.EnableEvents = False
    Application.ScreenUpdating = False

    ws.Activate
    Dim n As Variant
    n = ExecuteExcel4Macro("GET.DOCUMENT(50)")

    Dim pages As Long
    If IsError(n) Then
        pages = GetTotalPagesForSheet(ws)
    Else
        pages = CLng(n)
        If pages < 1 Then pages = GetTotalPagesForSheet(ws)
    End If

    If Not prevSheet Is Nothing Then prevSheet.Activate
    Application.EnableEvents = prevEvents
    Application.ScreenUpdating = prevUpdating

    GetTotalPagesForSheetExact = pages
    On Error GoTo 0
End Function

Private Function GetParentFolderName(ByVal dirPath As String) As String
    Dim p As String: p = Trim$(dirPath)
    If Len(p) = 0 Then GetParentFolderName = "": Exit Function
    ' 去除末尾分隔符（兼容 : / \）
    Do While Right$(p, 1) = Application.PathSeparator Or Right$(p, 1) = "/" Or Right$(p, 1) = "\"
        p = Left$(p, Len(p) - 1)
    Loop
    ' 先用 Application.PathSeparator 解析
    Dim sep As String: sep = Application.PathSeparator
    Dim pos As Long: pos = InStrRev(p, sep)
    If pos > 1 Then
        Dim parentPath As String: parentPath = Left$(p, pos - 1)
        Dim pos2 As Long: pos2 = InStrRev(parentPath, sep)
        If pos2 > 0 Then
            GetParentFolderName = Mid$(parentPath, pos2 + 1)
            Exit Function
        Else
            GetParentFolderName = parentPath
            Exit Function
        End If
    End If
    ' 回退：尝试 / 与 \
    Dim lastSepSlash As Long: lastSepSlash = InStrRev(p, "/")
    Dim lastSepBack As Long: lastSepBack = InStrRev(p, "\")
    Dim lastSep As Long: lastSep = IIf(lastSepSlash > lastSepBack, lastSepSlash, lastSepBack)
    If lastSep > 1 Then
        Dim parentPath2 As String: parentPath2 = Left$(p, lastSep - 1)
        Dim last2Slash As Long: last2Slash = InStrRev(parentPath2, "/")
        Dim last2Back As Long: last2Back = InStrRev(parentPath2, "\")
        Dim last2 As Long: last2 = IIf(last2Slash > last2Back, last2Slash, last2Back)
        If last2 > 0 Then
            GetParentFolderName = Mid$(parentPath2, last2 + 1)
        Else
            GetParentFolderName = parentPath2
        End If
    Else
        GetParentFolderName = ""
    End If
End Function

Private Function GetTotalPagesForSheet(ByVal ws As Worksheet) As Long
    On Error Resume Next
    ' 计算方法：总页数 ≈ (水平分页数+1) * (垂直分页数+1)
    ' 该方法在常规自动分页场景下可靠；若有手动分页/缩放变化，会随 PageSetup 自动更新。
    Dim h As Long: h = ws.HPageBreaks.Count
    Dim v As Long: v = ws.VPageBreaks.Count
    Dim pages As Long: pages = (h + 1) * (v + 1)
    If pages < 1 Then pages = 1
    GetTotalPagesForSheet = pages
    On Error GoTo 0
End Function

' 获取操作系统用户名（优先 USERNAME，其次 USER；最后回退 Excel 用户名）
Private Function GetOSUserName() As String
    Dim u As String
    u = Environ$("USERNAME")
    If Len(u) = 0 Then u = Environ$("USER")
    If Len(u) = 0 Then u = Application.UserName
    GetOSUserName = u
End Function

Public Sub Run_Print_AllSheets_Separately()
    On Error Resume Next
    Dim ws As Worksheet
    For Each ws In ActiveWorkbook.Worksheets
        If ws.Visible = xlSheetVisible Then
            ws.PrintOut
        End If
    Next ws
    On Error GoTo 0
End Sub

Public Sub Run_Export_AllSheets_ToPDF()
    On Error Resume Next
    Dim wb As Workbook: Set wb = ActiveWorkbook
    Dim baseDir As String: baseDir = Utils.WorkbookDir(wb)
    Dim sep As String: sep = Application.PathSeparator
    Dim ws As Worksheet
    For Each ws In wb.Worksheets
        If ws.Visible = xlSheetVisible Then
            ws.ExportAsFixedFormat Type:=xlTypePDF, _
                Filename:=baseDir & sep & ws.Name & ".pdf", _
                Quality:=xlQualityStandard, IncludeDocProperties:=True, _
                IgnorePrintAreas:=False, OpenAfterPublish:=False
        End If
    Next ws
    On Error GoTo 0
End Sub