VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "打地鼠（窗口）"
   ClientHeight    =   3030
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   3030
   ScaleWidth      =   4560
   Begin VB.CommandButton Command2 
      Caption         =   "暂停"
      Height          =   495
      Left            =   2640
      TabIndex        =   3
      Top             =   1200
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   360
      Top             =   1800
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   700
      Left            =   720
      Top             =   2160
   End
   Begin VB.CommandButton Command1 
      Caption         =   "开始"
      Height          =   495
      Left            =   2640
      TabIndex        =   0
      Top             =   1920
      Width           =   975
   End
   Begin VB.Label Label6 
      Height          =   375
      Left            =   720
      TabIndex        =   7
      Top             =   1200
      Width           =   1215
   End
   Begin VB.Label Label5 
      Caption         =   "最高分"
      Height          =   375
      Left            =   0
      TabIndex        =   6
      Top             =   1200
      Width           =   615
   End
   Begin VB.Label Label4 
      Caption         =   "时间"
      Height          =   495
      Left            =   0
      TabIndex        =   5
      Top             =   600
      Width           =   615
   End
   Begin VB.Label Label3 
      Caption         =   "得分"
      Height          =   255
      Left            =   0
      TabIndex        =   4
      Top             =   120
      Width           =   495
   End
   Begin VB.Label Label2 
      Height          =   375
      Left            =   720
      TabIndex        =   2
      Top             =   600
      Width           =   1095
   End
   Begin VB.Label Label1 
      Height          =   255
      Left            =   720
      TabIndex        =   1
      Top             =   120
      Width           =   1215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim conn As New ADODB.Connection        '声明ADO的connection对象实例
Dim rs As New ADODB.Recordset           '声明ADO的recordset对象实例
Dim a As Integer
Dim b As Integer
Dim c As Integer '得分
Dim z As Boolean '记录此时窗体应具有的状态
Dim t As Integer

Private Sub Command1_Click()
Form2.Visible = False
Timer1.Enabled = True '窗体弹出的时间间隔
Command1.Visible = False
Timer2.Enabled = True '一段时间后结束
Command2.Visible = True
c = 0
z = False
t = 0
rs.Open "SELECT*FROM info", conn, 1, 3
rs.MoveFirst
Do While rs.EOF = False
If rs.Fields("passname") = 远哥的单机登录窗口.dlyhm Then
Label6.Caption = rs.Fields("max2")
rs.Close
Exit Do
End If
rs.MoveNext
Loop
End Sub



Private Sub Command2_Click()
If Timer1.Enabled = True Then
Timer1.Enabled = False
Timer2.Enabled = False
Command2.Caption = "继续"
Else
Timer1.Enabled = True
Timer2.Enabled = True
Command2.Caption = "暂停"
End If
End Sub

Private Sub Form_Load()
conn.ConnectionString = "provider=Microsoft.ACE.OLEDB.12.0;DATA Source=" + App.Path + "\sjk.accdb"
conn.Open
Set rs.ActiveConnection = conn          '设置ADO记录集recordset对象连接属性
End Sub

Private Sub Form_Unload(Cancel As Integer)
Unload Form2
End Sub

Private Sub Timer1_Timer()

If z = False Then '随机位置弹出窗口
a = Int(Rnd() * 10001)
b = Int(Rnd() * 10001)
Form2.Left = a
Form2.Top = b
Form2.Visible = True
z = True
ElseIf z = True And Form2.Visible = False Then
c = c + 1
Label1.Caption = Str(c)
z = False
If Rnd() < 0.3 Then t = t - 1

Else: z = False
Form2.Visible = False
End If
Timer1.Interval = 2000 / (((c \ 10 + 1) / (c \ 10 + 2)) * 6)

End Sub

Private Sub Timer2_Timer() '结束计分
t = t + 1
If t >= 31 Then
Timer2.Enabled = False
Timer1.Enabled = False
Command2.Visible = False
Label1.Caption = "得分" + Str(c)
MsgBox "时间到"
rs.Open "SELECT*FROM info", conn, 1, 3
If rs.Fields("max2") <= c Then
rs.Close
rs.Open "SELECT*FROM info", conn, 1, 3
Do While rs.EOF = False
If rs.Fields("passname") = 远哥的单机登录窗口.dlyhm Then
Dim q1 As String
Dim w1 As String
Dim e1 As Integer
q1 = rs.Fields("passname")
w1 = rs.Fields("password")
e1 = rs.Fields("max")
rs.Delete adAffectCurrent
rs.AddNew
rs.Fields("passname").Value = q1
rs.Fields("password").Value = w1
rs.Fields("max").Value = e1
rs.Fields("max2").Value = Str(c)
rs.Update
Exit Do
End If
rs.MoveNext
Loop
End If
rs.Close
Command1.Visible = True
rs.Open "SELECT*FROM info", conn, 1, 3
rs.MoveFirst
Do While rs.EOF = False
If rs.Fields("passname") = 远哥的单机登录窗口.dlyhm Then
Label6.Caption = rs.Fields("max2")
rs.Close
Exit Do
End If
rs.MoveNext
Loop
Else: Label2.Caption = Str(30 - t)
End If

End Sub
