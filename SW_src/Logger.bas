' Logger.bas - simple logging
Option Explicit

Private gLogFile As String

Public Sub Logger_Init(path As String)
    If Len(Trim$(path)) = 0 Then
        ' 兜底：若未提供日志路径，写到临时目录
        gLogFile = Environ$("TEMP") & "\ExportNestedBOM_" & Format(Now, "YYYYMMDD_hhnnss") & ".log"
    Else
        gLogFile = path
    End If
    On Error Resume Next
    Dim f As Integer: f = FreeFile
    Open gLogFile For Output As #f
    Print #f, Now & " [INIT] 日志启动"
    Close #f
    On Error GoTo 0
End Sub

Public Sub Logger_Info(msg As String)
    WriteLog "INFO", msg
End Sub

Public Sub Logger_Warn(msg As String)
    WriteLog "WARN", msg
End Sub

Public Sub Logger_Error(msg As String)
    WriteLog "ERROR", msg
End Sub

Private Sub WriteLog(level As String, msg As String)
    On Error Resume Next
    ' 保护：若日志文件过大（>5MB），重新开始
    Dim needRotate As Boolean: needRotate = False
    Dim fsize As Long
    fsize = FileLen(gLogFile)
    If fsize > 5 * 1024 * 1024 Then needRotate = True
    
    Dim f As Integer: f = FreeFile
    If needRotate Then
        Open gLogFile For Output As #f
        Print #f, Now & " [ROTATE] 由于文件过大重置日志"
    Else
        Open gLogFile For Append As #f
    End If
    Print #f, Now & " [" & level & "] " & msg
    Close #f
    On Error GoTo 0
End Sub