VERSION 5.00
Begin VB.Form ���� 
   Caption         =   "Form1"
   ClientHeight    =   7215
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   9900
   LinkTopic       =   "Form1"
   ScaleHeight     =   7215
   ScaleWidth      =   9900
   StartUpPosition =   2  '��Ļ����
   WindowState     =   2  'Maximized
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   1320
      Top             =   4320
   End
   Begin VB.PictureBox Picture1 
      Height          =   2895
      Left            =   5040
      Picture         =   "Form1.frx":0000
      ScaleHeight     =   2835
      ScaleWidth      =   3915
      TabIndex        =   3
      Top             =   1320
      Width           =   3975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "��ʼ"
      Height          =   975
      Left            =   2520
      TabIndex        =   2
      Top             =   1320
      Width           =   1575
   End
   Begin VB.Label Label4 
      Caption         =   "������߷�"
      Height          =   375
      Left            =   0
      TabIndex        =   5
      Top             =   1200
      Width           =   1455
   End
   Begin VB.Label Label3 
      Height          =   615
      Left            =   0
      TabIndex        =   4
      Top             =   1680
      Width           =   1455
   End
   Begin VB.Label Label2 
      Height          =   495
      Left            =   0
      TabIndex        =   1
      Top             =   600
      Width           =   1335
   End
   Begin VB.Label Label1 
      Caption         =   "�÷�"
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   1335
   End
End
Attribute VB_Name = "����"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim conn As New ADODB.Connection        '����ADO��connection����ʵ��
Dim rs As New ADODB.Recordset           '����ADO��recordset����ʵ��
Dim a As Integer
Dim b As Integer
Dim c As Integer
Dim t As Integer
Dim xl As Integer
Dim d As Boolean
Dim m As Boolean
Dim z As Boolean

Private Sub Command1_Click()
Picture1.Visible = False
Timer1.Enabled = True
Command1.Visible = False
Picture1.Visible = True
t = 30
Label2.Caption = Str(t)
c = 0
Label1.Caption = "�÷�" + Str(c)
xl = 3
Label2.Caption = Str(xl)
d = False
m = True
rs.Open "SELECT*FROM info", conn, 1, 3
rs.MoveFirst
Do While rs.EOF = False
If rs.Fields("passname") = Զ��ĵ�����¼����.dlyhm Then
Label3.Caption = rs.Fields("max")
rs.Close
Exit Do
End If
rs.MoveNext
Loop
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If m And KeyCode = 27 Then
Timer1.Enabled = False
m = False
End If
If (m = False) And (KeyCode = 8) Then
Timer1.Enabled = True
m = True
End If
End Sub

Private Sub Form_Load()
conn.ConnectionString = "provider=Microsoft.ACE.OLEDB.12.0;DATA Source=" + App.Path + "\sjk.accdb"
conn.Open
Set rs.ActiveConnection = conn          '����ADO��¼��recordset������������
End Sub


'Private Sub Picture1_Click()
'c = c + 1
'Label1.Caption = "�÷�" + Str(c)
'Picture1.Visible = False
'd = True
'If Rnd() * 100 < c Then
'X = X + 1
'Label2.Caption = Str(X)
'End If
'End Sub

Private Sub Picture1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
c = c + 1
Label1.Caption = "�÷�" + Str(c)
Picture1.Visible = False
d = True
If Rnd() * 100 < c / 4 Then
xl = xl + 1
Label2.Caption = Str(xl)
End If
End Sub

Private Sub Timer1_Timer()
If Picture1.Visible = True Then
Picture1.Visible = False
If Not d Then
xl = xl - 1
Label2.Caption = Str(xl)
End If
End If
If z = False Then
a = Int(Rnd() * 14000)
b = Int(Rnd() * 10000)
Picture1.Left = a
Picture1.Top = b
Picture1.Visible = True
z = True
Else: z = False
End If
d = False
If xl = 0 Then
Timer1.Enabled = False
Picture1.Visible = False
Label1.Caption = "�÷�" + Str(c)
Command1.Visible = True
rs.Open "SELECT*FROM info", conn, 1, 3
If rs.Fields("max") <= c Then
rs.Close
rs.Open "SELECT*FROM info", conn, 1, 3
Do While rs.EOF = False
If rs.Fields("passname") = Զ��ĵ�����¼����.dlyhm Then
Dim q1 As String
Dim w1 As String
Dim e1 As Integer
q1 = rs.Fields("passname")
w1 = rs.Fields("password")
e1 = rs.Fields("max2")
rs.Delete adAffectCurrent
rs.AddNew
rs.Fields("passname").Value = q1
rs.Fields("password").Value = w1
rs.Fields("max2").Value = e1
rs.Fields("max").Value = Str(c)
rs.Update
Exit Do
End If
rs.MoveNext
Loop
End If
rs.Close
End If
Timer1.Interval = 2000 / (((c \ 5 + 1) / (c \ 5 + 2)) * 4.5)
End Sub

'Private Sub Timer3_Timer()
't = t - 1
'Label2.Caption = Str(t)
'If t = 0 Then
'Timer1.Enabled = False
'Picture1.Visible = False
'Label1.Caption = "�÷�" + Str(c)
'Command1.Visible = True
'Timer3.Enabled = False
'End If
'End Sub

