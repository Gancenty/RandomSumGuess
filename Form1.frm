VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "�������Ϸ"
   ClientHeight    =   4590
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6495
   LinkTopic       =   "Form1"
   ScaleHeight     =   4590
   ScaleWidth      =   6495
   StartUpPosition =   1  '����������
   Begin VB.CommandButton Command2 
      Caption         =   "��������"
      Height          =   615
      Left            =   600
      TabIndex        =   2
      Top             =   3120
      Width           =   2055
   End
   Begin VB.CommandButton Command1 
      Caption         =   "��ʼ��Ϸ"
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
b = InputBox("������һ����[100,200]֮�������", "�²�")
If (IsNumeric(b) = False) Then
res = MsgBox("��������", 0 + 48, "��ʾ")
GoTo begin
End If
If (b = "") Then
res = MsgBox("��������Ϊ��", 0 + 48, "��ʾ")
Image1(0).Visible = False
Image1(1).Visible = True
Image1(2).Visible = False
GoTo begin
End If
If (b > a) Then
res = MsgBox("��������ִ���", 0 + 48, "���������")
Image1(0).Visible = False
Image1(1).Visible = True
Image1(2).Visible = False
GoTo begin
End If
If (b < a) Then
res = MsgBox("���������С��", 0 + 48, "���������")
Image1(0).Visible = False
Image1(1).Visible = True
Image1(2).Visible = False
GoTo begin
End If
If (b = a) Then
res1 = MsgBox("��ϲ��¶��ˣ��Ƿ������Ϸ��", 1 + 64, "congratulations��")
Image1(0).Visible = False
Image1(1).Visible = False
Image1(2).Visible = True
End If
If (res1 = 1) Then
Command2_Click
GoTo begin
End If
'res���������msgbox�ķ���ֵ
'res1��������ȷmsgbox�ķ���ֵ
End Sub

Private Sub Command2_Click()
Image1(0).Visible = True
Image1(1).Visible = False
Image1(2).Visible = False
Command1.Enabled = False
Randomize
a = Int((100 - 200 + 1) * Rnd + 200)
'Command2.Caption = a ȡ��ע�Ϳ�����ʾ�����������
Command1.Enabled = True
End Sub

Private Sub Form_Load()
Label1.Caption = "�������������һ��[100��200]֮�������������Ҳ²�����֡�"
Label1.FontSize = 15
Command1.Enabled = False
Image1(0).Visible = False
Image1(1).Visible = False
Image1(2).Visible = False
End Sub
