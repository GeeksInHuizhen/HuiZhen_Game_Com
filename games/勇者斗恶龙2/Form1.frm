VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   7485
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   9750
   LinkTopic       =   "Form1"
   ScaleHeight     =   7485
   ScaleWidth      =   9750
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command4 
      Caption         =   "对话"
      Height          =   615
      Left            =   5880
      TabIndex        =   8
      Top             =   6480
      Width           =   1815
   End
   Begin VB.CommandButton Command3 
      Caption         =   "重生"
      Height          =   495
      Left            =   3480
      TabIndex        =   7
      Top             =   6720
      Width           =   1815
   End
   Begin VB.CommandButton Command2 
      Caption         =   "防御"
      Height          =   615
      Left            =   6240
      TabIndex        =   6
      Top             =   5400
      Width           =   1095
   End
   Begin VB.ListBox List1 
      Height          =   4200
      Left            =   3600
      TabIndex        =   5
      Top             =   960
      Width           =   3975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "攻击"
      Height          =   975
      Left            =   3600
      TabIndex        =   0
      Top             =   5400
      Width           =   1215
   End
   Begin VB.Label Label4 
      Caption         =   "勇士血量"
      Height          =   615
      Left            =   7920
      TabIndex        =   4
      Top             =   2880
      Width           =   1935
   End
   Begin VB.Label Label3 
      Caption         =   "巨龙血量"
      Height          =   615
      Left            =   1440
      TabIndex        =   3
      Top             =   3480
      Width           =   1815
   End
   Begin VB.Label Label2 
      Height          =   1215
      Left            =   7920
      TabIndex        =   2
      Top             =   3960
      Width           =   1695
   End
   Begin VB.Label Label1 
      Height          =   975
      Left            =   1560
      TabIndex        =   1
      Top             =   4200
      Width           =   1575
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim yongshixuexian As Integer
Dim yongshixueliang As Integer
Dim julongxueliang As Integer
Dim i As Boolean
Dim julongxuexian As Integer
Private Sub Command1_Click()
If i = False Then
If yongshixueliang > 0 And julongxueliang > 0 Then
julongxueliang = julongxueliang - 125
List1.AddItem "勇士发动了攻击"
If (65 < Int(Rnd * 100) + 1) Then
yongshixueliang = yongshixueliang - 25
List1.AddItem "巨龙向勇士吐出了烈焰"
End If
If yongshixueliang = 0 Then
List1.AddItem "可悲的勇士死了，但恶魔会放过他吗？"
End If
If julongxueliang = 0 Then
List1.AddItem "伟大的勇士击杀了巨龙"
End If
Else
MsgBox "恶魔强迫你重生"
End If
If yongshixueliang = 0 And julongxueliang = 0 Then
List1.AddItem "勇士和巨龙同归于尽了，恶魔失去了心爱的玩物"
i = True
End If
Label1.Caption = (Str(julongxueliang) + "/" + Str(julongxuexian))
Label2.Caption = (Str(yongshixueliang) + "/" + Str(yongshixuexian))
Else
List1.Clear
List1.AddItem "恶魔出现了"
List1.AddItem "恶魔出现了"
List1.AddItem "恶魔出现了"
List1.AddItem "恶魔出现了"
List1.AddItem "恶魔出现了"
List1.AddItem "恶魔出现了"
List1.AddItem "恶魔出现了"
End If
End Sub
Private Sub Command2_Click()
List1.AddItem "勇士可笑地选择了防御，然而巨龙并没有攻击"
End Sub
Private Sub Command3_Click()
yongshixuexian = 100
yongshixueliang = 100
julongxueliang = 1000
julongxuexian = 1000
List1.Clear
Label1.Caption = (Str(julongxueliang) + "/" + Str(julongxuexian))
Label2.Caption = (Str(yongshixueliang) + "/" + Str(yongshixuexian))
List1.AddItem "恶魔的游戏又开始了"
End Sub

Private Sub Form_Load()
yongshixuexian = 100
yongshixueliang = 100
julongxueliang = 1000
julongxuexian = 1000
List1.Clear
Label1.Caption = (Str(julongxueliang) + "/" + Str(julongxuexian))
Label2.Caption = (Str(yongshixueliang) + "/" + Str(yongshixuexian))
List1.AddItem "恶魔驱使着勇士向巨龙走去"
End Sub

