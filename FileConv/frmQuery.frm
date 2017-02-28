VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmQuery 
   Caption         =   "Error message content"
   ClientHeight    =   2600
   ClientLeft      =   90
   ClientTop       =   410
   ClientWidth     =   7060
   OleObjectBlob   =   "frmQuery.frx":0000
   StartUpPosition =   1  '所有者中心
End
Attribute VB_Name = "frmQuery"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim queryTitle As String
Dim queryMessage As String
Dim queryRecommendStrings As String
Dim button1_Message As String
Dim button2_Message As String
Dim button3_Message As String
Dim SameOperationMessage As String

Dim Same As Boolean
Dim result As Long
Dim mfilter As Long
Dim btnCount As Integer
Dim showCheck As Boolean

Public Property Let RepeatOptionVisible(value As Boolean)
    If (showCheck <> value) Then
        Me.Height = Me.Height + IIf(value, 1, -1) * cbxSame.Height
    End If
    showCheck = value
    cbxSame.Visible = showCheck
End Property

Public Property Get RepeatOptionVisible() As Boolean
    RepeatOptionVisible = showCheck
End Property

Public Property Let ButtonCount(value As Integer)
    If value < 1 Or value > 3 Then
        Error 9
    Else
        btn1.Visible = True
        btn2.Visible = (value >= 2)
        btn3.Visible = (value >= 3)
        Select Case value
        Case 1
            btn1.Left = (Me.InsideWidth - btn1.Width) / 2
        Case 2
            btn1.Left = (Me.InsideWidth - btn1.Width * 2.5) / 2
            btn2.Left = btn1.Left + btn1.Width * 1.5
        Case 3
            btn1.Left = (Me.InsideWidth - btn1.Width * 4) / 2
            btn2.Left = btn1.Left + btn1.Width * 1.5
            btn3.Left = btn2.Left + btn2.Width * 1.5
        End Select
    End If
End Property

Public Property Get value() As Long
    value = result
End Property

Public Property Get IsSameClicked() As Boolean
    IsSameClicked = Same
End Property

Private Sub LoadStrings()
    queryTitle = "基本错误提示窗口"
    queryMessage = "<错误提示信息>"
    queryRecommendStrings = "你想要准备怎么操作？"
    SameOperationMessage = "对于以后的文件进行相同的操作"
    button1_Message = "缺省按钮1"
    button2_Message = "缺省按钮2"
    button3_Message = "缺省按钮3"
End Sub

Private Sub btn1_Click()
    result = 1
    Me.Hide
End Sub

Private Sub btn2_Click()
    result = 2
    Me.Hide
End Sub

Private Sub btn3_Click()
    result = 3
    Me.Hide
End Sub

Private Sub cbxSame_Click()
    Same = cbxSame.value
    btn1.Enabled = (Not Same) Or ((mfilter And 1) <> 0)
    btn2.Enabled = (Not Same) Or ((mfilter And 2) <> 0)
    btn3.Enabled = (Not Same) Or ((mfilter And 4) <> 0)
End Sub

Private Sub UserForm_Initialize()
    LoadStrings
    mfilter = -1
    Me.Caption = queryTitle
    msg1.Caption = queryMessage
    msg2.Caption = queryRecommendStrings
    btn1.Caption = button1_Message
    btn2.Caption = button2_Message
    btn3.Caption = button3_Message
    cbxSame.Caption = SameOperationMessage
    showCheck = True
    Same = False
    ButtonCount = 3
    
End Sub

Public Property Get OptionFilter() As Long
    OptionFilter = mfilter
End Property

Public Property Let OptionFilter(ByVal value As Long)
    mfilter = value
End Property

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    Cancel = 1
End Sub
