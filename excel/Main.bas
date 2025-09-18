Attribute VB_Name = "Main"
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