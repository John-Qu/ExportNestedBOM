' Configuration.bas - Runtime configuration and validation
Option Explicit

' 配置结构
Public Type SystemConfig
    MaxRecursionDepth As Integer
    EnableCircularCheck As Boolean
    AutoCreateBackup As Boolean
    LogLevel As String ' INFO, WARN, ERROR
    OutputFormat As String ' XLS, XLSX
    IncludeImages As Boolean
    OverwriteExisting As Boolean
End Type

Private gConfig As SystemConfig

' 初始化默认配置
Public Sub InitializeConfig()
    With gConfig
        .MaxRecursionDepth = 10
        .EnableCircularCheck = True
        .AutoCreateBackup = False
        .LogLevel = "INFO"
        .OutputFormat = "XLS"
        .IncludeImages = True
        .OverwriteExisting = True
    End With
End Sub

' 获取当前配置
Public Function GetConfig() As SystemConfig
    GetConfig = gConfig
End Function

' 设置配置项
Public Sub SetMaxRecursionDepth(depth As Integer)
    If depth >= 1 And depth <= 20 Then
        gConfig.MaxRecursionDepth = depth
    Else
        Err.Raise 5, "SetMaxRecursionDepth", "递归深度必须在1-20之间"
    End If
End Sub

Public Sub SetLogLevel(level As String)
    Select Case UCase$(level)
        Case "INFO", "WARN", "ERROR"
            gConfig.LogLevel = UCase$(level)
        Case Else
            Err.Raise 5, "SetLogLevel", "日志级别必须是 INFO, WARN, ERROR 之一"
    End Select
End Sub

' 验证系统环境
Public Function ValidateEnvironment() As String
    Dim issues As String
    
    ' 检查SolidWorks版本
    On Error Resume Next
    Dim swApp As Object
    Set swApp = Application.SldWorks
    If swApp Is Nothing Then
        issues = issues & "- 无法连接到SolidWorks应用程序" & vbCrLf
    Else
        Dim version As String
        version = swApp.RevisionNumber
        If Len(version) = 0 Then
            issues = issues & "- 无法获取SolidWorks版本信息" & vbCrLf
        End If
    End If
    On Error GoTo 0
    
    ' 检查临时目录写权限
    Dim tempDir As String
    tempDir = Environ$("TEMP")
    If Not CanWriteToDirectory(tempDir) Then
        issues = issues & "- 临时目录无写权限：" & tempDir & vbCrLf
    End If
    
    ValidateEnvironment = issues
End Function

' 显示配置信息
Public Sub ShowConfiguration()
    Dim msg As String
    msg = "当前配置：" & vbCrLf & _
          "最大递归深度：" & gConfig.MaxRecursionDepth & vbCrLf & _
          "循环引用检查：" & IIf(gConfig.EnableCircularCheck, "启用", "禁用") & vbCrLf & _
          "自动备份：" & IIf(gConfig.AutoCreateBackup, "启用", "禁用") & vbCrLf & _
          "日志级别：" & gConfig.LogLevel & vbCrLf & _
          "输出格式：" & gConfig.OutputFormat & vbCrLf & _
          "包含图片：" & IIf(gConfig.IncludeImages, "是", "否") & vbCrLf & _
          "覆盖文件：" & IIf(gConfig.OverwriteExisting, "是", "否")
    
    MsgBox msg, vbInformation, "系统配置"
End Sub