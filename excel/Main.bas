' Attribute VB_Name = "Main"
Option Explicit

' 入口一：仅格式化当前工作簿的 BOM 数据表，并导出各工作表 PDF（不做总汇总与分类）
Public Sub Run_FormatAndExport_CurrentWorkbook()
    On Error GoTo EH
    LogInit ActiveWbDir()

    ' 格式化
    FormatActiveWorkbookBOMSheets

    ' 导出 PDF
    ExportWorkbookSheetsToPdf Application.ActiveWorkbook, ActiveWbDir() & Application.PathSeparator & CFG_PDF_OutputDir

    LogInfo "Run_FormatAndExport_CurrentWorkbook 完成"
    GoTo DONE
EH:
    LogError "Run_FormatAndExport_CurrentWorkbook 失败: " & Err.Description
DONE:
    LogClose
End Sub

' 入口二：以当前工作簿的“汇总”驱动，生成“总 BOM 清单”与分类表，统一格式化并导出 PDF
Public Sub Run_BuildSummary_And_Export()
    On Error GoTo EH
    LogInit ActiveWbDir()

    BuildTotalSummaryFromTopSummary

    ' 导出 PDF（包含总表与分类表）
    ExportWorkbookSheetsToPdf Application.ActiveWorkbook, ActiveWbDir() & Application.PathSeparator & CFG_PDF_OutputDir

    ' 可选合并：主模型与子装配 PDF 合并（需要你构建 pdfList 并调用 TryMergePdfs）
    ' 示例（按需替换为真实列表）:
    ' Dim pdfs As New Collection
    ' pdfs.Add ActiveWbDir() & Application.PathSeparator & CFG_PDF_OutputDir & Application.PathSeparator & "主模型.pdf"
    ' pdfs.Add ActiveWbDir() & Application.PathSeparator & CFG_PDF_OutputDir & Application.PathSeparator & "子装配A.pdf"
    ' TryMergePdfs ActiveWbDir() & Application.PathSeparator & CFG_PDF_OutputDir & Application.PathSeparator & "合并输出.pdf", pdfs

    LogInfo "Run_BuildSummary_And_Export 完成"
    GoTo DONE
EH:
    LogError "Run_BuildSummary_And_Export 失败: " & Err.Description
DONE:
    LogClose
End Sub

' 入口三：完整流水线（格式化所有数据表 -> 生成总汇总与分类 -> 导出 PDF）
Public Sub Run_FullPipeline()
    On Error GoTo EH
    LogInit ActiveWbDir()

    FormatActiveWorkbookBOMSheets
    BuildTotalSummaryFromTopSummary
    ExportWorkbookSheetsToPdf Application.ActiveWorkbook, ActiveWbDir() & Application.PathSeparator & CFG_PDF_OutputDir

    LogInfo "Run_FullPipeline 完成"
    GoTo DONE
EH:
    LogError "Run_FullPipeline 失败: " & Err.Description
DONE:
    LogClose
End Sub

' 获取目标工作簿（非宏文件），优先选择第一个非xlsm的已打开工作簿
Public Function GetTargetWorkbook() As Workbook
    Dim wb As Workbook
    For Each wb In Application.Workbooks
        If LCase$(Right$(wb.Name, 5)) <> ".xlsm" Then
            Set GetTargetWorkbook = wb
            Exit Function
        End If
    Next wb
    Set GetTargetWorkbook = Nothing
End Function
Public Sub Run_ToolboxReplace_StepByStep()
    On Error GoTo EH
    Dim wbTarget As Workbook: Set wbTarget = GetTargetWorkbook()
    If wbTarget Is Nothing Then
        LogInit ActiveWbDir()
        LogError "未找到目标工作簿，请先打开要处理的BOM文件再运行。"
        GoTo DONE
    End If
    wbTarget.Activate

    LogInit ActiveWbDir()
    SingleSheetFormatter.ApplyToolboxReplacement_StepByStep

    LogInfo "Run_ToolboxReplace_StepByStep 完成"
    GoTo DONE
EH:
    LogError "Run_ToolboxReplace_StepByStep 失败: " & Err.Description
DONE:
    LogClose
End Sub

Public Sub Run_ToolboxReplace_StepByStep_WPS()
    Dim wb As Workbook
    Dim toolboxMap As Object
    Set wb = GetTargetWorkbook()
    Set toolboxMap = Utils.LoadToolboxMapping()
    Call ApplyToolboxReplacement_StepByStep_WPS(wb, toolboxMap)
    MsgBox "全部副本已生成，请逐一检查。", vbInformation
End Sub