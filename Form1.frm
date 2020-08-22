VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "随机数游戏"
   ClientHeight    =   4590
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6495
   LinkTopic       =   "Form1"
   ScaleHeight     =   4590
   ScaleWidth      =   6495
   StartUpPosition =   1  '所有者中心
   Begin VB.CommandButton Command2 
      Caption         =   "生成数字"
      Height          =   615
      Left            =   600
      TabIndex        =   2
      Top             =   3120
      Width           =   2055
   End
   Begin VB.CommandButton Command1 
      Caption         =   "开始游戏"
      Height          =   615
      Left            =   3480
      TabIndex        =   0
      Top             =   3120
      Width           =   2055
   End
   Begin VB.Image Image1 
      Height          =   1035
      Index           =   2
      Left            =   1920
      Picture         =   "Form1.frx":0000
      Stretch         =   -1  'True
      Top             =   1800
      Width           =   2160
   End
   Begin VB.Image Image1 
      Height          =   1035
      Index           =   1
      Left            =   1800
      Picture         =   "Form1.frx":081E
      Stretch         =   -1  'True
      Top             =   1920
      Width           =   2160
   End
   Begin VB.Image Image1 
      Height          =   1035
      Index           =   0
      Left            =   2040
      Picture         =   "Form1.frx":120A
      Stretch         =   -1  'True
      Top             =   1800
      Width           =   2160
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   1215
      Left            =   480
      TabIndex        =   1
      Top             =   240
      Width           =   5535
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim a As Integer

Private Sub Command1_Click()
Dim b As String
begin:
b = InputBox("请输入一个在[100,200]之间的整数", "猜测")
If (IsNumeric(b) = False) Then
res = MsgBox("输入有误", 0 + 48, "提示")
GoTo begin
End If
If (b = "") Then
res = MsgBox("输入数字为空", 0 + 48, "提示")
Image1(0).Visible = False
Image1(1).Visible = True
Image1(2).Visible = False
GoTo begin
End If
If (b > a) Then
res = MsgBox("输入的数字大了", 0 + 48, "重新输入嗷")
Image1(0).Visible = False
Image1(1).Visible = True
Image1(2).Visible = False
GoTo begin
End If
If (b < a) Then
res = MsgBox("输入的数字小了", 0 + 48, "重新输入嗷")
Image1(0).Visible = False
Image1(1).Visible = True
Image1(2).Visible = False
GoTo begin
End If
If (b = a) Then
res1 = MsgBox("恭喜你猜对了，是否继续游戏？", 1 + 64, "congratulations！")
Image1(0).Visible = False
Image1(1).Visible = False
Image1(2).Visible = True
End If
If (res1 = 1) Then
Command2_Click
GoTo begin
End If
'res是输入错误msgbox的返回值
'res1是输入正确msgbox的返回值
End Sub

Private Sub Command2_Click()
Image1(0).Visible = True
Image1(1).Visible = False
Image1(2).Visible = False
Command1.Enabled = False
Randomize
a = Int((100 - 200 + 1) * Rnd + 200)
'Command2.Caption = a 取消注释可以显示产生的随机数
Command1.Enabled = True
End Sub

Private Sub Form_Load()
Label1.Caption = "本程序随机生成一个[100，200]之间的整数，请玩家猜测该数字。"
Label1.FontSize = 15
Command1.Enabled = False
Image1(0).Visible = False
Image1(1).Visible = False
Image1(2).Visible = False
End Sub
