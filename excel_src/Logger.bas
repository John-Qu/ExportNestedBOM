Option Explicit

Private g_inited As Boolean
Private g_logFile As Integer
Private g_logPath As String

Public Sub LogInit(ByVal baseDir As String, Optional ByVal tag As String = "run")
    On Error Resume Next
    If Right$(baseDir, 1) = "\" Or Right$(baseDir, 1) = "/" Then
        baseDir = Left$(baseDir, Len(baseDir) - 1)
    End If
    Dim logDir As String
    logDir = baseDir & "\" & CFG_LOG_FOLDER
    If Dir(logDir, vbDirectory) = "" Then MkDir logDir
    Dim ts As String
    ts = Format(Now, "yyyymmdd-hhnnss")
    g_logPath = logDir & "\" & tag & "-" & ts & ".txt"
    g_logFile = FreeFile
    Open g_logPath For Append As #g_logFile
    g_inited = True
    Print #g_logFile, "=== START " & Now & " ==="
    On Error GoTo 0
End Sub

Public Sub LogInfo(ByVal msg As String)
    If Not g_inited Then Exit Sub
    Print #g_logFile, Format(Now, "yyyy-mm-dd hh:nn:ss") & " [INFO] " & msg
End Sub

Public Sub LogWarn(ByVal msg As String)
    If Not g_inited Then Exit Sub
    Print #g_logFile, Format(Now, "yyyy-mm-dd hh:nn:ss") & " [WARN] " & msg
End Sub

Public Sub LogError(ByVal msg As String)
    If Not g_inited Then Exit Sub
    Print #g_logFile, Format(Now, "yyyy-mm-dd hh:nn:ss") & " [ERROR] " & msg
End Sub

Public Function LogPath() As String
    LogPath = g_logPath
End Function

Public Sub LogClose()
    If g_inited Then
        Print #g_logFile, "=== END " & Now & " ==="
        Close #g_logFile
        g_inited = False
    End If
End Sub