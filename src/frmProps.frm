VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmProps 
   Caption         =   "UserForm1"
   ClientHeight    =   11445
   ClientLeft      =   105
   ClientTop       =   450
   ClientWidth     =   10980
   OleObjectBlob   =   "frmProps.frx":0000
   StartUpPosition =   1  '����������
End
Attribute VB_Name = "frmProps"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'========================================================
' frmProps - 属性确认对话框（动态行： 名称 | 值(可编辑) | 勾选导入）
' 控件需要：lblInfo, fraList, cmdOK, cmdSkip, cmdCancel, （可选）chkAll
'========================================================
Option Explicit

Private mFileInfo As String
Private mNames As Variant
Private mValues As Variant

' 动态控件数组
Private lbls() As MSForms.Label
Private txts() As MSForms.TextBox
Private chks() As MSForms.CheckBox

' 输出
Public UserAction As Integer ' 1=OK, 2=Skip, 3=Cancel
Public SelectedNames As Variant
Public SelectedValues As Variant
Public SelectedChecks As Variant

Public Sub InitWithData(ByVal fileInfo As String, ByVal names As Variant, ByVal values As Variant)
    mFileInfo = fileInfo
    mNames = names
    mValues = values
    
    Me.Caption = "属性确认"
    lblInfo.Caption = mFileInfo
    
    BuildRows
End Sub

Private Sub BuildRows()
    Dim i As Long, topY As Single
    Dim rowH As Single: rowH = 24
    Dim leftName As Single: leftName = 10
    Dim leftVal As Single: leftVal = 200
    Dim leftChk As Single: leftChk = 470
    Dim widthName As Single: widthName = 180
    Dim widthVal As Single: widthVal = 260
    
    ' 清空原有动态控件
    Dim ctl As MSForms.Control
    For Each ctl In fraList.Controls
        If TypeName(ctl) = "Label" Or TypeName(ctl) = "TextBox" Or TypeName(ctl) = "CheckBox" Then
            ' 延迟删除避免遍历冲突
        End If
    Next ctl
    Do While fraList.Controls.Count > 0
        fraList.Controls.Remove fraList.Controls(0).name
        If fraList.Controls.Count = 0 Then Exit Do
    Loop
    
    ReDim lbls(LBound(mNames) To UBound(mNames))
    ReDim txts(LBound(mNames) To UBound(mNames))
    ReDim chks(LBound(mNames) To UBound(mNames))
    
    topY = 6
    
    ' 表头
    Dim hdr1 As MSForms.Label, hdr2 As MSForms.Label, hdr3 As MSForms.Label
    Set hdr1 = fraList.Controls.Add("Forms.Label.1", "hdrName")
    hdr1.Caption = "属性名"
    hdr1.Left = leftName: hdr1.Top = topY: hdr1.Width = widthName: hdr1.Font.Bold = True
    
    Set hdr2 = fraList.Controls.Add("Forms.Label.1", "hdrVal")
    hdr2.Caption = "值（可编辑）"
    hdr2.Left = leftVal: hdr2.Top = topY: hdr2.Width = widthVal: hdr2.Font.Bold = True
    
    Set hdr3 = fraList.Controls.Add("Forms.Label.1", "hdrChk")
    hdr3.Caption = "导入"
    hdr3.Left = leftChk: hdr3.Top = topY: hdr3.Width = 40: hdr3.Font.Bold = True
    
    topY = topY + rowH
    
    ' 行
    For i = LBound(mNames) To UBound(mNames)
        Dim lb As MSForms.Label, tb As MSForms.TextBox, ck As MSForms.CheckBox
        
        Set lb = fraList.Controls.Add("Forms.Label.1", "lbl" & CStr(i))
        lb.Caption = CStr(mNames(i))
        lb.Left = leftName
        lb.Top = topY + 3
        lb.Width = widthName
        
        Set tb = fraList.Controls.Add("Forms.TextBox.1", "txt" & CStr(i))
        tb.Text = CStr(mValues(i))
        tb.Left = leftVal
        tb.Top = topY
        tb.Width = widthVal
        tb.Tag = CStr(mNames(i)) ' 保存属性名
        tb.IntegralHeight = False
        
        Set ck = fraList.Controls.Add("Forms.CheckBox.1", "chk" & CStr(i))
        ck.value = True ' 默认勾选导入
        ck.Left = leftChk
        ck.Top = topY + 2
        ck.Width = 40
        
        Set lbls(i) = lb
        Set txts(i) = tb
        Set chks(i) = ck
        
        topY = topY + rowH
    Next i
    
    ' 设置滚动高度
    fraList.ScrollTop = 0
    fraList.ScrollHeight = topY + 6
End Sub

Private Sub cmdOK_Click()
    CollectAndClose 1
End Sub

Private Sub cmdSkip_Click()
    CollectAndClose 2
End Sub

Private Sub cmdCancel_Click()
    UserAction = 3
    Me.Hide
End Sub

Private Sub CollectAndClose(ByVal actionCode As Integer)
    Dim i As Long
    Dim n() As String, v() As String, c() As Boolean
    ReDim n(LBound(mNames) To UBound(mNames))
    ReDim v(LBound(mValues) To UBound(mValues))
    ReDim c(LBound(mNames) To UBound(mNames))
    
    For i = LBound(mNames) To UBound(mNames)
        n(i) = lbls(i).Caption
        v(i) = txts(i).Text
        c(i) = chks(i).value
    Next i
    
    SelectedNames = n
    SelectedValues = v
    SelectedChecks = c
    UserAction = actionCode
    Me.Hide
End Sub

Private Sub chkAll_Click()
    On Error Resume Next
    Dim i As Long
    For i = LBound(chks) To UBound(chks)
        chks(i).value = chkAll.value
    Next i
    On Error GoTo 0
End Sub

Private Sub UserForm_Initialize()
    ' 可在此设定一些默认尺寸或字体
End Sub
