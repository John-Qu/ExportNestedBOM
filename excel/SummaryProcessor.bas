' Attribute VB_Name = "SummaryProcessor"
Option Explicit

' 从“顶层装配体\汇总”驱动，生成“总 BOM 清单”和分类表，并保持与单表一致的格式规范
Public Sub BuildTotalSummaryFromTopSummary()
    On Error GoTo EH
    LogInit ActiveWbDir()

    Dim topWb As Workbook: Set topWb = Application.ActiveWorkbook
    Dim topWs As Worksheet
    Set topWs = Nothing

    Dim ws As Worksheet
    For Each ws In topWb.Worksheets
        If StrComp(ws.Name, "汇总", vbTextCompare) = 0 Then
            Set topWs = ws: Exit For
        End If
    Next ws

    If topWs Is Nothing Then
        LogError "未找到“汇总”工作表"
        GoTo DONE
    End If

    Dim folder As String: folder = ActiveWbDir()
    Dim bomFiles As Collection: Set bomFiles = ListBOMFiles(folder)
    If bomFiles Is Nothing Or bomFiles.Count = 0 Then
        LogWarn "未找到任何子装配 BOM 文件（排除包含“汇总”的文件）"
    End If

    ' 准备“总 BOM 清单”工作表
    Dim totalWs As Worksheet
    On Error Resume Next
    Set totalWs = topWb.Worksheets("总 BOM 清单")
    On Error GoTo 0
    If totalWs Is Nothing Then
        Set totalWs = topWb.Worksheets.Add(After:=topWb.Worksheets(topWb.Worksheets.Count))
        totalWs.Name = "总 BOM 清单"
    Else
        totalWs.Cells.Clear
    End If

    ' 写标题行（按规范最终列序）
    Dim order As Variant: order = CFG_FinalColumnOrder()
    Dim i As Long
    For i = LBound(order) To UBound(order)
        totalWs.Cells(1, i + 1).Value = CStr(order(i))
    Next i
    ' 备注右侧新增“计算说明”
    totalWs.Cells(1, GetColumnIndex(totalWs, "备注") + 1).Value = "备注" ' 占位以确保索引
    Dim calcCol As Long
    calcCol = GetColumnIndex(totalWs, "备注") + 1
    totalWs.Cells(1, calcCol).Value = "计算说明"

    Dim lastTopRow As Long: lastTopRow = LastUsedRow(topWs, 1)
    Dim keyCol As Long
    keyCol = GetColumnIndex(topWs, "代号")
    If keyCol = 0 Then
        LogError "“汇总”表未找到“代号”列"
        GoTo DONE
    End If

    Dim outRow As Long: outRow = 2
    Dim dictDone As Object: Set dictDone = CreateObject("Scripting.Dictionary")
    dictDone.CompareMode = 1

    Dim r As Long
    For r = 2 To lastTopRow
        Dim key As String: key = Trim$(CStr(topWs.Cells(r, keyCol).Value))
        If Len(key) = 0 Then GoTo NEXTROW
        If dictDone.Exists(key) Then GoTo NEXTROW

        Dim sumQty As Double: sumQty = 0
        Dim descRowCopied As Boolean: descRowCopied = False
        Dim calcExplain As String: calcExplain = ""

        Dim filePath As Variant
        For Each filePath In bomFiles
            Dim wb As Workbook
            On Error Resume Next
            Set wb = Application.Workbooks.Open(CStr(filePath), ReadOnly:=True)
            On Error GoTo 0
            If Not wb Is Nothing Then
                Dim wsCandidate As Worksheet
                For Each wsCandidate In wb.Worksheets
                    If wsCandidate.Visible = xlSheetVisible Then
                        Dim partNoCol As Long: partNoCol = GetColumnIndex(wsCandidate, "零件号")
                        Dim qtyCol As Long: qtyCol = GetColumnIndex(wsCandidate, "数量")
                        If partNoCol > 0 And qtyCol > 0 Then
                            Dim lastRow As Long: lastRow = LastUsedRow(wsCandidate, partNoCol)
                            Dim rr As Long
                            For rr = 2 To lastRow
                                If Trim$(CStr(wsCandidate.Cells(rr, partNoCol).Value)) = key Then
                                    Dim q As Double: q = Val(CStr(wsCandidate.Cells(rr, qtyCol).Value))
                                    sumQty = sumQty + q
                                    If Len(calcExplain) = 0 Then
                                        calcExplain = "=" & CStr(q)
                                    Else
                                        calcExplain = calcExplain & "+" & CStr(q)
                                    End If
                                    If Not descRowCopied Then
                                        ' 复制描述性字段到输出表
                                        Call CopyRowByHeaders(wsCandidate, rr, totalWs, outRow)
                                        descRowCopied = True
                                    End If
                                End If
                            Next rr
                        End If
                    End If
                Next wsCandidate
                wb.Close SaveChanges:=False
                Set wb = Nothing
            Else
                LogWarn "无法打开子装配文件: " & CStr(filePath)
            End If
        Next filePath

        If Not descRowCopied Then
            ' 如果所有子装配都未匹配到，跳过但记录
            LogWarn "未在子装配中找到 key=" & key & " 的条目"
            GoTo NEXTROW
        End If

        ' 填充总数量与计算说明
        Dim qtyOutCol As Long: qtyOutCol = GetColumnIndex(totalWs, "数量")
        If qtyOutCol > 0 Then totalWs.Cells(outRow, qtyOutCol).Value = sumQty
        If Len(calcExplain) > 0 Then
            totalWs.Cells(outRow, calcCol).Value = calcExplain & "=" & CStr(sumQty)
        End If

        dictDone.Add key, True
        outRow = outRow + 1

NEXTROW:
    Next r

    ' 对“总 BOM 清单”做格式化（S3-S7）
    ApplyHeaderRenameRules totalWs
    ReorderColumnsToSpec totalWs
    ApplyBooleanIconization totalWs
    ApplyFontAndAlignment totalWs
    ApplyPrintSetup totalWs

    ' G6 分类输出工作表
    BuildCategorySheetsFromTotal totalWs

    LogInfo "BuildTotalSummaryFromTopSummary 完成"
    GoTo DONE
EH:
    LogError "BuildTotalSummaryFromTopSummary 失败: " & Err.Description
DONE:
    LogClose
End Sub

Private Sub BuildCategorySheetsFromTotal(ByVal totalWs As Worksheet)
    On Error GoTo EH
    Dim wb As Workbook: Set wb = totalWs.Parent

    ' 预定义分类
    Dim categories As Object: Set categories = CreateObject("Scripting.Dictionary")
    categories.CompareMode = 1
    categories("外购件") = Array("购", CFG_Icon_True)
    categories("钣金件") = Array("钣", CFG_Icon_True)
    categories("机箱模型") = Array("名称", "KW_ENCLOSURE") ' 特殊：基于关键词匹配

    Dim order As Variant: order = CFG_FinalColumnOrder()

    Dim key As Variant
    For Each key In categories.Keys
        Dim wsCat As Worksheet
        On Error Resume Next
        Set wsCat = wb.Worksheets(CStr(key))
        On Error GoTo 0
        If wsCat Is Nothing Then
            Set wsCat = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
            wsCat.Name = CStr(key)
        Else
            wsCat.Cells.Clear
        End If
        ' 标题
        Dim i As Long
        For i = LBound(order) To UBound(order)
            wsCat.Cells(1, i + 1).Value = CStr(order(i))
        Next i

        ' 复制符合条件的行
        Dim totalLast As Long: totalLast = LastUsedRow(totalWs, 1)
        Dim outRow As Long: outRow = 2
        Dim cond As Variant: cond = categories(key)

        Dim r As Long
        For r = 2 To totalLast
            If CStr(key) = "机箱模型" Then
                Dim colName As Long: colName = GetColumnIndex(totalWs, "名称")
                If colName > 0 Then
                    Dim nameVal As String: nameVal = CStr(totalWs.Cells(r, colName).Value)
                    If IsNameMatchesKeywords(nameVal, CFG_Keywords_Enclosure()) Then
                        CopyRowByHeaders totalWs, r, wsCat, outRow
                        outRow = outRow + 1
                    End If
                End If
            Else
                Dim targetHeader As String: targetHeader = CStr(cond(0))
                Dim expected As String: expected = CStr(cond(1))
                Dim col As Long: col = GetColumnIndex(totalWs, targetHeader)
                If col > 0 Then
                    If Trim$(CStr(totalWs.Cells(r, col).Value)) = expected Then
                        CopyRowByHeaders totalWs, r, wsCat, outRow
                        outRow = outRow + 1
                    End If
                End If
            End If
        Next r

        ' 分类表同样格式化与打印设置
        ApplyFontAndAlignment wsCat
        ApplyPrintSetup wsCat

        Set wsCat = Nothing
    Next key

    Exit Sub
EH:
    LogError "BuildCategorySheetsFromTotal 分类构建失败: " & Err.Description
End Sub

Private Function IsNameMatchesKeywords(ByVal nameVal As String, ByVal kws As Variant) As Boolean
    Dim i As Long
    For i = LBound(kws) To UBound(kws)
        If InStr(1, nameVal, CStr(kws(i)), vbTextCompare) > 0 Then
            IsNameMatchesKeywords = True
            Exit Function
        End If
    Next i
    IsNameMatchesKeywords = False
End Function