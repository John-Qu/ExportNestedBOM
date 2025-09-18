Attribute VB_Name = "Logger"
Option Explicit

Private g_logFilePath As String
Private g_logInited As Boolean

Public Sub LogInit(Optional ByVal baseDir As String = "")
    On Error Resume Next
    Dim dirPath As String
    If Len(baseDir) = 0 Then
        If Not Application Is Nothing And Not Application.ActiveWorkbook Is Nothing Then
            dirPath = Application.ActiveWorkbook.Path
        Else
            dirPath = CurDir
        End If
    Else
        dirPath = baseDir
    End If

    Dim logsDir As String
    logsDir = dirPath & Application.PathSeparator & "logs"
    Call EnsureDirExists(logsDir)

    Dim stamp As String
    stamp = Format(Now, "yyyymmdd-hhnnss")
    g_logFilePath = logsDir & Application.PathSeparator & "run-" & stamp & ".txt"

    Dim ff As Integer: ff = FreeFile
    Open g_logFilePath For Append As #ff
    Print #ff, "===== Log Start " & Now & " ====="
    Close #ff
    g_logInited = True
    On Error GoTo 0
End Sub

Public Sub LogInfo(ByVal msg As String)
    LogWrite "INFO", msg
End Sub

Public Sub LogWarn(ByVal msg As String)
    LogWrite "WARN", msg
End Sub

Public Sub LogError(ByVal msg As String)
    LogWrite "ERROR", msg
End Sub

Private Sub LogWrite(ByVal level As String, ByVal msg As String)
    On Error Resume Next
    Dim line As String
    line = Format(Now, "yyyy-mm-dd hh:nn:ss") & " [" & level & "] " & msg
    Debug.Print line
    If Not g_logInited Then Call LogInit
    Dim ff As Integer: ff = FreeFile
    Open g_logFilePath For Append As #ff
    Print #ff, line
    Close #ff
    On Error GoTo 0
End Sub

Public Sub LogClose()
    If g_logInited Then
        LogInfo "===== Log End ====="
        g_logInited = False
    End If
End Sub