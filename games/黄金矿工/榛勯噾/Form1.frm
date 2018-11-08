VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Form1"
   ClientHeight    =   10740
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   14145
   LinkTopic       =   "Form1"
   ScaleHeight     =   10740
   ScaleWidth      =   14145
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command2 
      Caption         =   "商店"
      Height          =   735
      Left            =   12360
      TabIndex        =   2
      Top             =   240
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "开始"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15.75
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2055
      Left            =   5040
      TabIndex        =   1
      Top             =   6000
      Width           =   3615
   End
   Begin VB.Timer Timer4 
      Enabled         =   0   'False
      Interval        =   40000
      Left            =   2880
      Top             =   9840
   End
   Begin VB.Timer Timer3 
      Left            =   2280
      Top             =   9840
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   1560
      Top             =   9840
   End
   Begin VB.Timer Timer1 
      Interval        =   10
      Left            =   960
      Top             =   9840
   End
   Begin VB.Image Image2 
      Height          =   840
      Index           =   2
      Left            =   2040
      Picture         =   "Form1.frx":0000
      Top             =   1440
      Width           =   1170
   End
   Begin VB.Image Image2 
      Height          =   840
      Index           =   1
      Left            =   3720
      Picture         =   "Form1.frx":3775
      Top             =   3480
      Width           =   1170
   End
   Begin VB.Image Image2 
      Height          =   840
      Index           =   0
      Left            =   8400
      Picture         =   "Form1.frx":6EEA
      Stretch         =   -1  'True
      Top             =   3000
      Width           =   1170
   End
   Begin VB.Image Image1 
      Height          =   450
      Index           =   4
      Left            =   5520
      Picture         =   "Form1.frx":A65F
      Stretch         =   -1  'True
      Top             =   3240
      Width           =   465
   End
   Begin VB.Image Image1 
      Height          =   450
      Index           =   3
      Left            =   4560
      Picture         =   "Form1.frx":D916
      Stretch         =   -1  'True
      Top             =   4320
      Width           =   465
   End
   Begin VB.Image Image1 
      Height          =   450
      Index           =   2
      Left            =   8040
      Picture         =   "Form1.frx":10BCD
      Stretch         =   -1  'True
      Top             =   4440
      Width           =   465
   End
   Begin VB.Image Image1 
      Height          =   450
      Index           =   1
      Left            =   3840
      Picture         =   "Form1.frx":13E84
      Stretch         =   -1  'True
      Top             =   2640
      Width           =   465
   End
   Begin VB.Image Image1 
      Height          =   450
      Index           =   0
      Left            =   8160
      Picture         =   "Form1.frx":1713B
      Stretch         =   -1  'True
      Top             =   2400
      Width           =   465
   End
   Begin VB.Label Label1 
      Caption         =   "0"
      Height          =   495
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   615
   End
   Begin VB.Shape Shape2 
      Height          =   255
      Left            =   6600
      Shape           =   3  'Circle
      Top             =   720
      Width           =   255
   End
   Begin VB.Line Line1 
      X1              =   6720
      X2              =   6720
      Y1              =   360
      Y2              =   840
   End
   Begin VB.Shape Shape1 
      Height          =   495
      Left            =   6480
      Top             =   120
      Width           =   495
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public stq As Integer
Public zy As Integer
Public sd As Integer
Public money As Integer
Dim fs As Integer
Dim stone As Integer
Dim st As Boolean
Dim jingkuai(10) As Boolean
Dim nd As Integer
Dim js1 As Integer
Dim fx1 As Boolean
Dim fx As Boolean
Dim js As Integer
Dim ll As Integer

Private Sub Command1_Click()
Command2.Visible = False
Label1.Caption = "0"
fs = 0
st = False
js1 = 0
js = 0
ll = 480
fx = True
fx1 = True
nd = 270
Timer4.Enabled = True
Command1.Visible = False
For I = 0 To 10
 jingkuai(I) = True
Next I
Randomize
Image1(0).Left = Int(Rnd() * 8000 + 2000)
Image1(0).Top = Int(Rnd() * 4000 + 3000)
Image1(0).Width = Int(Rnd() * 300 + 175)
Image1(0).Height = Image1(0).Width
For m = 1 To 4
 Image1(m).Left = Int(Rnd() * 8000 + 2000)
 Image1(m).Top = Int(Rnd() * 4000 + 3000)
 Image1(m).Width = Int(Rnd() * 200 + 175)
 Image1(m).Height = Image1(m).Width
Next m
For n = 0 To 2 Step 1
 Image2(n).Left = Int(Rnd() * 8000 + 2000)
 Image2(n).Top = Int(Rnd() * 4000 + 3000)
Next n
End Sub

Private Sub Command2_Click()
Form2.Visible = True
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 8 And Not st Then
 Timer1.Enabled = False
 Timer2.Enabled = True
 st = False
End If
If KeyAscii = 8 And st And zy > 0 Then
 Image2(stone).Left = 100000
 Image2(stone).Top = 100000
 zy = zy - 1
 st = False
 js1 = js1 \ 10
End If
End Sub

Private Sub Form_Load()
stq = 1
zy = 0
sd = 10
st = False
js1 = 0
js = 0
ll = 480
fx = True
fx1 = True
nd = 270
For I = 0 To 10
 jingkuai(I) = True
Next I
Randomize
Image1(0).Left = Int(Rnd() * 8000 + 2000)
Image1(0).Top = Int(Rnd() * 4000 + 3000)
Image1(0).Width = Int(Rnd() * 300 + 175)
Image1(0).Height = Image1(0).Width
For m = 1 To 4
 Image1(m).Left = Int(Rnd() * 8000 + 2000)
 Image1(m).Top = Int(Rnd() * 4000 + 3000)
 Image1(m).Width = Int(Rnd() * 200 + 175)
 Image1(m).Height = Image1(m).Width
Next m
For n = 0 To 2 Step 1
 Image2(n).Left = Int(Rnd() * 8000 + 2000)
 Image2(n).Top = Int(Rnd() * 4000 + 3000)
Next n
End Sub

Private Sub Timer1_Timer()
If fx = True Then
 Line1.X2 = ll * Cos(js / 100 * 3.14) + Line1.X1
 Line1.Y2 = ll * Sin(js / 100 * 3.14) + Line1.Y1
 Shape2.Left = Line1.X2 - 255 / 2
 Shape2.Top = Line1.Y2 - 255 / 2
 js = js + 1
 If js = 100 Then
  fx = False
 End If
End If
If fx = False Then
 Line1.X2 = ll * Cos(js / 100 * 3.14) + Line1.X1
 Line1.Y2 = ll * Sin(js / 100 * 3.14) + Line1.Y1
 Shape2.Left = Line1.X2 - 255 / 2
 Shape2.Top = Line1.Y2 - 255 / 2
 js = js - 1
 If js = 0 Then
  fx = True
 End If
End If
End Sub

Private Sub Timer2_Timer()
If js1 = 0 And fx1 = False Then
 Timer2.Enabled = False
 Timer1.Enabled = True
 fx1 = True
End If
If fx1 = True Then
 Line1.X2 = 10 * sd * Cos(js / 100 * 3.14) + Line1.X2
 Line1.Y2 = 10 * sd * Sin(js / 100 * 3.14) + Line1.Y2
 Shape2.Left = Line1.X2 - 255 / 2
 Shape2.Top = Line1.Y2 - 255 / 2
 js1 = js1 + 1
 For I = 0 To 4
  If Line1.Y2 >= (Image1(I).Top - nd + 375 / 2) And Line1.Y2 <= (Image1(I).Top + nd + 375 / 2) And Line1.X2 >= (Image1(I).Left - nd + 375 / 2) And Line1.X2 <= (Image1(I).Left + nd + 375 / 2) Then
   Image1(I).Top = 100000
   Image1(I).Left = 100000
   fs = fs + 100 * (I + 1)
   money = money + 150
   Label1.Caption = CStr(Val(Label1.Caption) + 100 * (I + 1))
   If Val(Label1.Caption) = 1575 Then
    MsgBox "下一关"
    Randomize
    Image1(0).Left = Int(Rnd() * 8000 + 2000)
    Image1(0).Top = Int(Rnd() * 4000 + 3000)
    Image1(0).Width = Int(Rnd() * 300 + 175)
    Image1(0).Height = Image1(0).Width
    For m = 1 To 4
     Image1(m).Left = Int(Rnd() * 8000 + 2000)
     Image1(m).Top = Int(Rnd() * 4000 + 3000)
     Image1(m).Width = Int(Rnd() * 200 + 175)
     Image1(m).Height = Image1(m).Width
    Next m
    For n = 0 To 2 Step 1
     Image2(n).Left = Int(Rnd() * 8000 + 2000)
     Image2(n).Top = Int(Rnd() * 4000 + 3000)
    Next n
    Label1.Caption = "0"
   End If
   fx1 = False
  End If
 Next I
 For I = 0 To 2
  If Line1.Y2 >= (Image2(I).Top - nd + 840 / 2) And Line1.Y2 <= (Image2(I).Top + nd + 840 / 2) And Line1.X2 >= (Image2(I).Left - nd + 1170 / 2) And Line1.X2 <= (Image2(I).Left + nd + 1170 / 2) Then
   fx1 = False
   st = True
   stone = I
   js1 = js1 * 10
  End If
 Next I
 If js1 = Int(80 / (sd / 10)) Then
  fx1 = False
 End If
End If
If fx1 = False And (Not st) Then
 Line1.X2 = -sd * 10 * Cos(js / 100 * 3.14) + Line1.X2
 Line1.Y2 = -sd * 10 * Sin(js / 100 * 3.14) + Line1.Y2
 Shape2.Left = Line1.X2 - 255 / 2
 Shape2.Top = Line1.Y2 - 255 / 2
 js1 = js1 - 1
End If
If fx1 = False And st Then
 Line1.X2 = -sd * Cos(js / 100 * 3.14) + Line1.X2
 Line1.Y2 = -sd * Sin(js / 100 * 3.14) + Line1.Y2
 Shape2.Left = Line1.X2 - 255 / 2
 Shape2.Top = Line1.Y2 - 255 / 2
 js1 = js1 - 1
 If js1 = 0 Then
  Image2(stone).Left = 100000
  Image2(stone).Top = 100000
  money = money + (10 * stq)
  fs = fs + (25 * stq)
  Label1.Caption = CStr(Val(Label1.Caption) + (25 * stq))
  If Val(Label1.Caption) = 1575 Then
   MsgBox "下一关"
   Randomize
   Image1(0).Left = Int(Rnd() * 8000 + 2000)
   Image1(0).Top = Int(Rnd() * 4000 + 3000)
   Image1(0).Width = Int(Rnd() * 300 + 175)
   Image1(0).Height = Image1(0).Width
   For m = 1 To 4
    Image1(m).Left = Int(Rnd() * 8000 + 2000)
    Image1(m).Top = Int(Rnd() * 4000 + 3000)
    Image1(m).Width = Int(Rnd() * 200 + 175)
    Image1(m).Height = Image1(m).Width
   Next m
   For n = 0 To 2 Step 1
    Image2(n).Left = Int(Rnd() * 8000 + 2000)
    Image2(n).Top = Int(Rnd() * 4000 + 3000)
   Next n
  Label1.Caption = "0"
  End If
  st = False
 End If
 Image2(stone).Left = -sd * Cos(js / 100 * 3.14) + Image2(stone).Left
 Image2(stone).Top = -sd * Sin(js / 100 * 3.14) + Image2(stone).Top
End If
End Sub

Private Sub Timer4_Timer()
If fs >= 1200 Then
 MsgBox "下一关"
 st = False
 Randomize
 Image1(0).Left = Int(Rnd() * 8000 + 2000)
 Image1(0).Top = Int(Rnd() * 4000 + 3000)
 Image1(0).Width = Int(Rnd() * 300 + 175)
 Image1(0).Height = Image1(0).Width
 For m = 1 To 4
  Image1(m).Left = Int(Rnd() * 8000 + 2000)
  Image1(m).Top = Int(Rnd() * 4000 + 3000)
  Image1(m).Width = Int(Rnd() * 200 + 175)
  Image1(m).Height = Image1(m).Width
 Next m
 For n = 0 To 2 Step 1
  Image2(n).Left = Int(Rnd() * 8000 + 2000)
  Image2(n).Top = Int(Rnd() * 4000 + 3000)
 Next n
 Label1.Caption = "0"
 Timer4.Interval = Timer4.Interval - 1000
 fs = 0
 Command1.Visible = True
 Command2.Visible = True
 Timer4.Enabled = False
Else
 MsgBox "时间到"
 Command1.Visible = True
 Timer4.Enabled = False
End If
End Sub
