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
   StartUpPosition =   3  '����ȱʡ
   Begin VB.CommandButton Command4 
      Caption         =   "�Ի�"
      Height          =   615
      Left            =   5880
      TabIndex        =   8
      Top             =   6480
      Width           =   1815
   End
   Begin VB.CommandButton Command3 
      Caption         =   "����"
      Height          =   495
      Left            =   3480
      TabIndex        =   7
      Top             =   6720
      Width           =   1815
   End
   Begin VB.CommandButton Command2 
      Caption         =   "����"
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
      Caption         =   "����"
      Height          =   975
      Left            =   3600
      TabIndex        =   0
      Top             =   5400
      Width           =   1215
   End
   Begin VB.Label Label4 
      Caption         =   "��ʿѪ��"
      Height          =   615
      Left            =   7920
      TabIndex        =   4
      Top             =   2880
      Width           =   1935
   End
   Begin VB.Label Label3 
      Caption         =   "����Ѫ��"
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
List1.AddItem "��ʿ�����˹���"
If (65 < Int(Rnd * 100) + 1) Then
yongshixueliang = yongshixueliang - 25
List1.AddItem "��������ʿ�³�������"
End If
If yongshixueliang = 0 Then
List1.AddItem "�ɱ�����ʿ���ˣ�����ħ��Ź�����"
End If
If julongxueliang = 0 Then
List1.AddItem "ΰ�����ʿ��ɱ�˾���"
End If
Else
MsgBox "��ħǿ��������"
End If
If yongshixueliang = 0 And julongxueliang = 0 Then
List1.AddItem "��ʿ�;���ͬ���ھ��ˣ���ħʧȥ���İ�������"
i = True
End If
Label1.Caption = (Str(julongxueliang) + "/" + Str(julongxuexian))
Label2.Caption = (Str(yongshixueliang) + "/" + Str(yongshixuexian))
Else
List1.Clear
List1.AddItem "��ħ������"
List1.AddItem "��ħ������"
List1.AddItem "��ħ������"
List1.AddItem "��ħ������"
List1.AddItem "��ħ������"
List1.AddItem "��ħ������"
List1.AddItem "��ħ������"
End If
End Sub
Private Sub Command2_Click()
List1.AddItem "��ʿ��Ц��ѡ���˷�����Ȼ��������û�й���"
End Sub
Private Sub Command3_Click()
yongshixuexian = 100
yongshixueliang = 100
julongxueliang = 1000
julongxuexian = 1000
List1.Clear
Label1.Caption = (Str(julongxueliang) + "/" + Str(julongxuexian))
Label2.Caption = (Str(yongshixueliang) + "/" + Str(yongshixuexian))
List1.AddItem "��ħ����Ϸ�ֿ�ʼ��"
End Sub

Private Sub Form_Load()
yongshixuexian = 100
yongshixueliang = 100
julongxueliang = 1000
julongxuexian = 1000
List1.Clear
Label1.Caption = (Str(julongxueliang) + "/" + Str(julongxuexian))
Label2.Caption = (Str(yongshixueliang) + "/" + Str(yongshixuexian))
List1.AddItem "��ħ��ʹ����ʿ�������ȥ"
End Sub

