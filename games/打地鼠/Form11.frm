VERSION 5.00
Begin VB.Form Զ��ĵ�����¼���� 
   Caption         =   "Զ��ĵ�����¼����"
   ClientHeight    =   7695
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   7515
   FillColor       =   &H00FFFFFF&
   ForeColor       =   &H8000000E&
   LinkTopic       =   "Form1"
   ScaleHeight     =   7695
   ScaleWidth      =   7515
   StartUpPosition =   3  '����ȱʡ
   Begin VB.CommandButton Command7 
      Caption         =   "���ޱ���"
      Height          =   615
      Left            =   4560
      TabIndex        =   19
      Top             =   1320
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.CommandButton Command6 
      Caption         =   "ʱ�ս���"
      Height          =   615
      Left            =   4560
      TabIndex        =   18
      Top             =   360
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.CommandButton Command5 
      Caption         =   "ɾ��"
      Height          =   495
      Left            =   3240
      TabIndex        =   17
      Top             =   5400
      Width           =   1095
   End
   Begin VB.CommandButton Command4 
      Caption         =   "����"
      Height          =   495
      Left            =   3240
      TabIndex        =   16
      Top             =   4560
      Width           =   1095
   End
   Begin VB.CommandButton Command3 
      Caption         =   "�˳�"
      Height          =   375
      Left            =   4920
      TabIndex        =   15
      Top             =   3960
      Width           =   975
   End
   Begin VB.TextBox Text4 
      Height          =   495
      Left            =   3240
      TabIndex        =   14
      Top             =   3840
      Width           =   1215
   End
   Begin VB.ListBox List1 
      Height          =   2760
      Left            =   840
      TabIndex        =   13
      Top             =   3840
      Width           =   1815
   End
   Begin VB.TextBox Text3 
      Height          =   375
      Left            =   480
      TabIndex        =   6
      Text            =   "<������֤��>"
      Top             =   2640
      Width           =   1455
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Check1"
      Height          =   255
      Left            =   600
      TabIndex        =   5
      Top             =   2280
      Width           =   255
   End
   Begin VB.CommandButton Command2 
      Caption         =   "ע��"
      Height          =   375
      Left            =   2160
      TabIndex        =   3
      Top             =   1800
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "��¼"
      Height          =   375
      Left            =   2160
      TabIndex        =   2
      Top             =   1200
      Width           =   975
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   480
      TabIndex        =   1
      Text            =   "<��������>"
      Top             =   1800
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   480
      TabIndex        =   0
      Text            =   "<�����û���>"
      Top             =   1200
      Width           =   1455
   End
   Begin VB.Label Label8 
      Caption         =   "ע����Ϸ��Ϊ��������"
      Height          =   375
      Left            =   5400
      TabIndex        =   20
      Top             =   7440
      Width           =   2055
   End
   Begin VB.Label Label7 
      BackColor       =   &H8000000E&
      Caption         =   "�����壿��һ��"
      BeginProperty Font 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   255
      Left            =   3600
      TabIndex        =   12
      Top             =   2640
      Width           =   1335
   End
   Begin VB.Label Label6 
      Caption         =   "v 2.0.0"
      Height          =   255
      Left            =   120
      TabIndex        =   11
      Top             =   7320
      Width           =   1095
   End
   Begin VB.Label Label5 
      BeginProperty Font 
         Name            =   "Brush Script Std"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF00&
      Height          =   375
      Left            =   3240
      TabIndex        =   10
      Top             =   2640
      Width           =   255
   End
   Begin VB.Label Label4 
      BeginProperty Font 
         Name            =   "���Ĳ���"
         Size            =   7.5
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2880
      TabIndex        =   9
      Top             =   2640
      Width           =   255
   End
   Begin VB.Label Label3 
      BeginProperty Font 
         Name            =   "��������"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF00FF&
      Height          =   375
      Left            =   2520
      TabIndex        =   8
      Top             =   2640
      Width           =   255
   End
   Begin VB.Label Label2 
      BeginProperty Font 
         Name            =   "OCR A Std"
         Size            =   7.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   375
      Left            =   2160
      TabIndex        =   7
      Top             =   2640
      Width           =   255
   End
   Begin VB.Label Label1 
      Caption         =   "��ʾ����"
      Height          =   255
      Left            =   840
      TabIndex        =   4
      Top             =   2280
      Width           =   855
   End
End
Attribute VB_Name = "Զ��ĵ�����¼����"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' ������ע��
'
'              .======.
'              | INRI |
'              |      |
'              |      |
'     .========'      '========.
'     |   _      xxxx      _   |
'     |  /_;-.__ / _\  _.-;_\  |
'     |     `-._`'`_/'`.-'     |
'     '========.`\   /`========'
'              | |  / |
'              |/-.(  |
'              |\_._\ |
'              | \ \`;|
'              |  > |/|
'              | / // |
'              | |//  |
'              | \(\  |
'              |  ``  |
'              |      |
'              |      |
'              |      |
'              |      |
'  \\    _  _\\| \//  |//_   _ \// _
' ^ `^`^ ^`` `^ ^` ``^^`  `^^` `^ `^
' Ү�ձ��� ����bug
'


Dim conn As New ADODB.Connection        '����ADO��connection����ʵ��
Dim rs As New ADODB.Recordset           '����ADO��recordset����ʵ��
Dim xsmm As Boolean                     '����Ƿ���ʾ����
Dim a1 As Integer                       '��֤��1234
Dim b1 As Integer
Dim c1 As Integer
Public dlyhm As String
Dim d1 As Integer
Private Sub Check1_Click()
If xsmm Then                            '��֤�Ƿ�������ʾ
  If Text2.Text = "<��������>" Then     '��֤Text2_Click�Ƿ�ı�������ʾ
    Text2.Text = ""
  End If
  Text2.PasswordChar = "*"
  xsmm = False
Else
  Text2.PasswordChar = ""
  xsmm = True
End If
End Sub
Private Sub Command1_Click()
Dim d As Boolean                        '��֤��֤����ȷ
Dim c As Boolean                        '��֤�û����Ƿ�ע��
d = False
c = True
rs.Open "SELECT*FROM info", conn, 1, 3
rs.MoveFirst
If (" " + Text3.Text) = Str(a1 * 1000 + b1 * 100 + c1 * 10 + d1) Then
  d = True
Else
  MsgBox "������֤�����"
  c = False
End If
If d Then
  Do While rs.EOF = False
    If Not Text1.Text = "" Then
      If Text1.Text = rs.Fields("passname") Then    '�û�����ȷ...
        If Text2.Text = rs.Fields("password") Then
          MsgBox "��¼�ɹ�"
          dlyhm = Text1.Text
          c = False
          Command6.Visible = True
          Command7.Visible = True
          If Text1.Text = "����Ա" Then
            List1.Clear
            List1.Visible = True
            Command3.Visible = True
            Command4.Visible = True
            Command5.Visible = True
            Text4.Visible = True
            rs.MoveFirst
            Do While rs.EOF = False
              List1.AddItem rs.Fields("passname")
              rs.MoveNext
            Loop
          End If
        Else
          MsgBox "�����������"
          c = False
        End If
        Exit Do
      End If
    End If
    rs.MoveNext
  Loop
End If
If c Then MsgBox "�û����������"
rs.Close
Randomize                               '������֤��
Text3.Text = ""
a1 = Int(Rnd() * 9 + 1)
b1 = Int(Rnd() * 9 + 1)
c1 = Int(Rnd() * 9 + 1)
d1 = Int(Rnd() * 9 + 1)
Label2.Caption = Str(a1)
Label3.Caption = Str(b1)
Label4.Caption = Str(c1)
Label5.Caption = Str(d1)
End Sub
Private Sub Command2_Click()
Dim X As Boolean                        '��֤�û����Ƿ��Ѵ���
X = True
rs.Open "SELECT*FROM info", conn, 1, 3
rs.MoveFirst
Do While rs.EOF = False
  If Not Text1.Text = "" Then
    If Text1.Text = rs.Fields("passname") Then    '�û����Ѵ���...
      MsgBox "�û����Ѵ���"
      X = False
    End If
  End If
  rs.MoveNext
Loop
If X Then
  If (Not Text1.Text = "") And (Not Text2.Text = "") And (Not Text1.Text = "<�����û���>") And (Not Text2.Text = "<��������>") Then
    rs.AddNew
    rs.Fields("passname").Value = Text1.Text
    rs.Fields("password").Value = Text2.Text
    rs.Fields("max").Value = "0"
    rs.Fields("max2").Value = "0"
    rs.Update
    MsgBox "ע��ɹ�"
    Text1.Text = ""
    Text2.Text = ""
  Else
    MsgBox "�������û���������"
  End If
End If
rs.Close
End Sub
Private Sub Command3_Click()
List1.Clear
List1.Visible = False
Command3.Visible = False
Command4.Visible = False
Command5.Visible = False
Text4.Visible = False
End Sub
Private Sub Command4_Click()
Dim X As Boolean
X = False
rs.Open "SELECT*FROM info", conn, 1, 3
rs.MoveFirst
Do While rs.EOF = False
  If Not Text4.Text = "" Then
    If Text4.Text = rs.Fields("passname") Then    '�û�������...
      X = True
    End If
  End If
  rs.MoveNext
Loop
If X Then
  MsgBox "���û�����"
Else
  MsgBox "���û�������"
End If
rs.Close
End Sub
Private Sub Command5_Click()
If Str(2333 Mod 6 + 6666) = " " + InputBox("���������", "���������") Then
  rs.Open "SELECT*FROM info", conn, 1, 3
  rs.MoveFirst
  Do While rs.EOF = False
    If Not Text4.Text = "" Then
      If Text4.Text = rs.Fields("passname") Then    '�û�������...
        rs.Delete adAffectCurrent
        MsgBox "��ɾ��"
      End If
    End If
    rs.MoveNext
  Loop
  List1.Clear
  rs.MoveFirst
  Do While rs.EOF = False
    List1.AddItem rs.Fields("passname")
    rs.MoveNext
  Loop
  rs.Close
Else
  MsgBox "�������"
End If
End Sub

Private Sub Command6_Click()
Form1.Visible = True
End Sub

Private Sub Command7_Click()
����.Visible = True
End Sub

Private Sub Form_Load()
conn.ConnectionString = "provider=Microsoft.ACE.OLEDB.12.0;DATA Source=" + App.Path + "\sjk.accdb"
conn.Open
Set rs.ActiveConnection = conn          '����ADO��¼��recordset������������
xsmm = (Text2.PasswordChar = "*")       '��������
List1.Clear
List1.Visible = False
Command3.Visible = False
Command4.Visible = False
Command5.Visible = False
Text4.Visible = False
End Sub
Private Sub Label7_Click()              '������֤��
Randomize
Text3.Text = ""
a1 = Int(Rnd() * 9 + 1)
b1 = Int(Rnd() * 9 + 1)
c1 = Int(Rnd() * 9 + 1)
d1 = Int(Rnd() * 9 + 1)
Label2.Caption = Str(a1)
Label3.Caption = Str(b1)
Label4.Caption = Str(c1)
Label5.Caption = Str(d1)
End Sub
Private Sub Text1_Click()               '����ʱ���
If Text1.Text = "<�����û���>" Then
  Text1.Text = ""
End If
End Sub
Private Sub Text2_Click()               '����ʱ���
If Text2.Text = "<��������>" Then
  Text2.Text = ""
  If Not (xsmm) Then                    '�Ƿ�ѡ����ʾ���롱
    Text2.PasswordChar = "*"
  End If
End If
End Sub
Private Sub Text3_Click()               '����ʱ���
If Text3.Text = "<������֤��>" Then
  Randomize                             '������֤��
  Text3.Text = ""
  a1 = Int(Rnd() * 9 + 1)
  b1 = Int(Rnd() * 9 + 1)
  c1 = Int(Rnd() * 9 + 1)
  d1 = Int(Rnd() * 9 + 1)
  Label2.Caption = Str(a1)
  Label3.Caption = Str(b1)
  Label4.Caption = Str(c1)
  Label5.Caption = Str(d1)
End If
End Sub

'
'              .======.
'              | INRI |
'              |      |
'              |      |
'     .========'      '========.
'     |   _      xxxx      _   |
'     |  /_;-.__ / _\  _.-;_\  |
'     |     `-._`'`_/'`.-'     |
'     '========.`\   /`========'
'              | |  / |
'              |/-.(  |
'              |\_._\ |
'              | \ \`;|
'              |  > |/|
'              | / // |
'              | |//  |
'              | \(\  |
'              |  ``  |
'              |      |
'              |      |
'              |      |
'              |      |
'  \\    _  _\\| \//  |//_   _ \// _
' ^ `^`^ ^`` `^ ^` ``^^`  `^^` `^ `^
' Ү�ձ��� ����bug
