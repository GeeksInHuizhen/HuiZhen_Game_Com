VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Alpha CAN'T Go"
   ClientHeight    =   6210
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   8445
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   Picture         =   "Form1.frx":23D2
   ScaleHeight     =   6210
   ScaleWidth      =   8445
   StartUpPosition =   2  '屏幕中心
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   200
      Left            =   360
      Top             =   2640
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000011&
      Height          =   615
      Left            =   1920
      TabIndex        =   4
      Top             =   2640
      Visible         =   0   'False
      Width           =   4455
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000011&
      Height          =   615
      Left            =   1920
      TabIndex        =   3
      Top             =   2160
      Visible         =   0   'False
      Width           =   4455
   End
   Begin VB.Image Image29 
      Height          =   2250
      Left            =   1920
      Picture         =   "Form1.frx":50D8
      Top             =   3360
      Visible         =   0   'False
      Width           =   4500
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   26.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   1095
      Left            =   1920
      TabIndex        =   2
      Top             =   720
      Visible         =   0   'False
      Width           =   4455
   End
   Begin VB.Image Image28 
      Height          =   1275
      Left            =   5040
      Picture         =   "Form1.frx":BE65
      Top             =   3960
      Visible         =   0   'False
      Width           =   1275
   End
   Begin VB.Image Image27 
      Height          =   1275
      Left            =   3480
      Picture         =   "Form1.frx":11837
      Top             =   3960
      Visible         =   0   'False
      Width           =   1275
   End
   Begin VB.Image Image26 
      Height          =   1275
      Left            =   1920
      Picture         =   "Form1.frx":17209
      Top             =   3960
      Visible         =   0   'False
      Width           =   1275
   End
   Begin VB.Image Image25 
      Height          =   1275
      Left            =   5040
      Picture         =   "Form1.frx":1CBDB
      Top             =   2400
      Visible         =   0   'False
      Width           =   1275
   End
   Begin VB.Image Image24 
      Height          =   1275
      Left            =   3480
      Picture         =   "Form1.frx":225AD
      Top             =   2400
      Visible         =   0   'False
      Width           =   1275
   End
   Begin VB.Image Image23 
      Height          =   1275
      Left            =   1920
      Picture         =   "Form1.frx":27F7F
      Top             =   2400
      Visible         =   0   'False
      Width           =   1275
   End
   Begin VB.Image Image22 
      Height          =   1275
      Left            =   5040
      Picture         =   "Form1.frx":2D951
      Top             =   840
      Visible         =   0   'False
      Width           =   1275
   End
   Begin VB.Image Image21 
      Height          =   1275
      Left            =   3480
      Picture         =   "Form1.frx":33323
      Top             =   840
      Visible         =   0   'False
      Width           =   1275
   End
   Begin VB.Image Image20 
      Height          =   1275
      Left            =   1920
      Picture         =   "Form1.frx":38CF5
      Top             =   840
      Visible         =   0   'False
      Width           =   1275
   End
   Begin VB.Image Image19 
      Height          =   1275
      Left            =   5040
      Picture         =   "Form1.frx":3E6C7
      Top             =   3960
      Visible         =   0   'False
      Width           =   1275
   End
   Begin VB.Image Image18 
      Height          =   1275
      Left            =   3480
      Picture         =   "Form1.frx":42BA9
      Top             =   3960
      Visible         =   0   'False
      Width           =   1275
   End
   Begin VB.Image Image17 
      Height          =   1275
      Left            =   1920
      Picture         =   "Form1.frx":4708B
      Top             =   3960
      Visible         =   0   'False
      Width           =   1275
   End
   Begin VB.Image Image16 
      Height          =   1275
      Left            =   5040
      Picture         =   "Form1.frx":4B56D
      Top             =   2400
      Visible         =   0   'False
      Width           =   1275
   End
   Begin VB.Image Image15 
      Height          =   1275
      Left            =   3480
      Picture         =   "Form1.frx":4FA4F
      Top             =   2400
      Visible         =   0   'False
      Width           =   1275
   End
   Begin VB.Image Image14 
      Height          =   1275
      Left            =   1920
      Picture         =   "Form1.frx":53F31
      Top             =   2400
      Visible         =   0   'False
      Width           =   1275
   End
   Begin VB.Image Image13 
      Height          =   1275
      Left            =   5040
      Picture         =   "Form1.frx":58413
      Top             =   840
      Visible         =   0   'False
      Width           =   1275
   End
   Begin VB.Image Image12 
      Height          =   1275
      Left            =   3480
      Picture         =   "Form1.frx":5C8F5
      Top             =   840
      Visible         =   0   'False
      Width           =   1275
   End
   Begin VB.Image Image11 
      Height          =   1275
      Left            =   1920
      Picture         =   "Form1.frx":60DD7
      Top             =   840
      Visible         =   0   'False
      Width           =   1275
   End
   Begin VB.Image Image10 
      Height          =   1275
      Left            =   5040
      Picture         =   "Form1.frx":652B9
      Top             =   3960
      Visible         =   0   'False
      Width           =   1275
   End
   Begin VB.Image Image9 
      Height          =   1275
      Left            =   3480
      Picture         =   "Form1.frx":68123
      Top             =   3960
      Visible         =   0   'False
      Width           =   1275
   End
   Begin VB.Image Image8 
      Height          =   1275
      Left            =   1920
      Picture         =   "Form1.frx":6AF8D
      Top             =   3960
      Visible         =   0   'False
      Width           =   1275
   End
   Begin VB.Image Image7 
      Height          =   1275
      Left            =   5040
      Picture         =   "Form1.frx":6DDF7
      Top             =   2400
      Visible         =   0   'False
      Width           =   1275
   End
   Begin VB.Image Image6 
      Height          =   1275
      Left            =   3480
      Picture         =   "Form1.frx":70C61
      Top             =   2400
      Visible         =   0   'False
      Width           =   1275
   End
   Begin VB.Image Image5 
      Height          =   1275
      Left            =   1920
      Picture         =   "Form1.frx":73ACB
      Top             =   2400
      Visible         =   0   'False
      Width           =   1275
   End
   Begin VB.Image Image4 
      Height          =   1275
      Left            =   5040
      Picture         =   "Form1.frx":76935
      Top             =   840
      Visible         =   0   'False
      Width           =   1275
   End
   Begin VB.Image Image3 
      Height          =   1275
      Left            =   3480
      Picture         =   "Form1.frx":7979F
      Top             =   840
      Visible         =   0   'False
      Width           =   1275
   End
   Begin VB.Image Image2 
      Height          =   1275
      Left            =   1920
      Picture         =   "Form1.frx":7C609
      Top             =   840
      Visible         =   0   'False
      Width           =   1275
   End
   Begin VB.Image Image1 
      Height          =   2250
      Left            =   1920
      Picture         =   "Form1.frx":7F473
      Top             =   3360
      Width           =   4500
   End
   Begin VB.Label Label2 
      Caption         =   "Alpha CAN'T Go"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   21.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000015&
      Height          =   975
      Left            =   2400
      TabIndex        =   1
      Top             =   1080
      Width           =   3495
   End
   Begin VB.Label Label1 
      Caption         =   "欢迎调戏"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   495
      Left            =   3720
      TabIndex        =   0
      Top             =   600
      Width           =   975
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim sta(1 To 9) As Integer
Dim a1 As Integer, a2 As Integer, a3 As Integer, a4 As Integer, a5 As Integer, a6 As Integer, a7 As Integer, a8 As Integer, a9 As Integer, fina As Integer
Dim checkres As Integer, dadada As Integer
Dim yes As Long, nono As Long, allall As Long
Dim isthis(1 To 9) As Double
Public Function check() As Integer
Dim ifsta As Integer
ifsta = 0 '平局
For b = 1 To 9
 If sta(b) = 0 Then
  ifsta = 1 '继续
  Exit For
 End If
Next b
If (sta(1) = 1 And sta(2) = 1) And sta(3) = 1 Then
 ifsta = 2 '用户赢
ElseIf (sta(1) = 1 And sta(4) = 1) And sta(7) = 1 Then
 ifsta = 2 '用户赢
ElseIf (sta(1) = 1 And sta(5) = 1) And sta(9) = 1 Then
 ifsta = 2 '用户赢
ElseIf (sta(2) = 1 And sta(5) = 1) And sta(8) = 1 Then
 ifsta = 2 '用户赢
ElseIf (sta(3) = 1 And sta(5) = 1) And sta(7) = 1 Then
 ifsta = 2 '用户赢
ElseIf (sta(3) = 1 And sta(6) = 1) And sta(9) = 1 Then
 ifsta = 2 '用户赢
ElseIf (sta(4) = 1 And sta(5) = 1) And sta(6) = 1 Then
 ifsta = 2 '用户赢
ElseIf (sta(7) = 1 And sta(8) = 1) And sta(9) = 1 Then
 ifsta = 2 '用户赢
ElseIf (sta(1) = 2 And sta(2) = 2) And sta(3) = 2 Then
 ifsta = 3 '电脑赢
ElseIf (sta(1) = 2 And sta(4) = 2) And sta(7) = 2 Then
 ifsta = 3 '电脑赢
ElseIf (sta(1) = 2 And sta(5) = 2) And sta(9) = 2 Then
 ifsta = 3 '电脑赢
ElseIf (sta(2) = 2 And sta(5) = 2) And sta(8) = 2 Then
 ifsta = 3 '电脑赢
ElseIf (sta(3) = 2 And sta(5) = 2) And sta(7) = 2 Then
 ifsta = 3 '电脑赢
ElseIf (sta(3) = 2 And sta(6) = 2) And sta(9) = 2 Then
 ifsta = 3 '电脑赢
ElseIf (sta(4) = 2 And sta(5) = 2) And sta(6) = 2 Then
 ifsta = 3 '电脑赢
ElseIf (sta(7) = 2 And sta(8) = 2) And sta(9) = 2 Then
 ifsta = 3 '电脑赢
End If
check = ifsta
End Function
Public Function endend2() As Integer
 Timer1.Enabled = True
End Function
Public Function endend3() As Integer
 Timer1.Enabled = False
 Image1.Visible = False
 Image2.Visible = False
 Image3.Visible = False
 Image4.Visible = False
 Image5.Visible = False
 Image6.Visible = False
 Image7.Visible = False
 Image8.Visible = False
 Image9.Visible = False
 Image10.Visible = False
 Image11.Visible = False
 Image12.Visible = False
 Image13.Visible = False
 Image14.Visible = False
 Image15.Visible = False
 Image16.Visible = False
 Image17.Visible = False
 Image18.Visible = False
 Image19.Visible = False
 Image20.Visible = False
 Image21.Visible = False
 Image22.Visible = False
 Image23.Visible = False
 Image24.Visible = False
 Image25.Visible = False
 Image26.Visible = False
 Image27.Visible = False
 Image28.Visible = False
 Label3.Visible = True
 Label4.Visible = True
 Label5.Visible = True
 dadada = check()
 If dadada = 0 Then
 Label3.Caption = "平局诶"
 Label4.Caption = "苟利国家生死以"
 Label5.Caption = "岂能不再苟一局"
 ElseIf dadada = 2 Then
 Label3.Caption = "你赢啦"
 Label4.Caption = "苟富贵"
 Label5.Caption = "要再来"
 ElseIf dadada = 3 Then
 Label3.Caption = "电脑赢了2333"
 Label4.Caption = "苟非吾之所有"
 Label5.Caption = "也要再苟一次"
 End If
 Image29.Visible = True
End Function
Public Function main(typeid As Integer, sdnum As Long) As Integer
Dim maxmax As Integer
maxmax = 0
If typeid = 0 Then '现在电脑下
 checkres = check()
 If checkres = 1 Then
  For a = 1 To 9
   If sta(a) = 0 Then
    sta(a) = 2
    If sdnum = 1 Then
     a1 = a
    ElseIf sdnum = 2 Then
     a2 = a
    ElseIf sdnum = 3 Then
     a3 = a
    ElseIf sdnum = 4 Then
     a4 = a
    ElseIf sdnum = 5 Then
     a5 = a
    ElseIf sdnum = 6 Then
     a6 = a
    ElseIf sdnum = 7 Then
     a7 = a
    ElseIf sdnum = 8 Then
     a8 = a
    ElseIf sdnum = 9 Then
     a9 = a
    End If
    a = 1
    fina = main(1, sdnum + 1)
    If sdnum = 1 Then
     a = a1
    ElseIf sdnum = 2 Then
     a = a2
    ElseIf sdnum = 3 Then
     a = a3
    ElseIf sdnum = 4 Then
     a = a4
    ElseIf sdnum = 5 Then
     a = a5
    ElseIf sdnum = 6 Then
     a = a6
    ElseIf sdnum = 7 Then
     a = a7
    ElseIf sdnum = 8 Then
     a = a8
    ElseIf sdnum = 9 Then
     a = a9
    End If
    sta(a) = 0
   End If
   If sdnum = 1 Then
   If allall = 0 Then
     allall = 1
    End If
    If yes = 0 Then
     yes = 1
    End If
    If nono = 0 Then
     nono = 1
    End If
    isthis(a) = (nono / allall) - (yes / nono) + (nono / yes) + (yes / allall)
   End If
  Next a
  If sdnum = 1 Then
    For c = 1 To 9
     If c = 1 Then
      maxmax = c
     ElseIf isthis(c) > isthis(maxmax) Then
      maxmax = c
     End If
    Next c
    If sta(maxmax) <> 0 Then
     For y = 1 To 9
      If sta(y) = 0 Then
       maxmax = y
      End If
     Next y
    End If
    allall = 0
    yes = 0
    nono = 0
   End If
 ElseIf checkres = 0 Then
  main = 0
  allall = allall + 1
 ElseIf checkres = 2 Then
  main = 0
  allall = allall + 1
  nono = nono + 1
 ElseIf checkres = 3 Then
  main = 0
  allall = allall + 1
  yes = yes + 1
 End If
End If
If typeid = 1 Then '现在模拟用户下
 checkres = check()
 If checkres = 1 Then
  For a = 1 To 9
   If sta(a) = 0 Then
    sta(a) = 1
    If sdnum = 1 Then
     a1 = a
    ElseIf sdnum = 2 Then
     a2 = a
    ElseIf sdnum = 3 Then
     a3 = a
    ElseIf sdnum = 4 Then
     a4 = a
    ElseIf sdnum = 5 Then
     a5 = a
    ElseIf sdnum = 6 Then
     a6 = a
    ElseIf sdnum = 7 Then
     a7 = a
    ElseIf sdnum = 8 Then
     a8 = a
    ElseIf sdnum = 9 Then
     a9 = a
    End If
    a = 1
    fina = main(0, sdnum + 1)
    If sdnum = 1 Then
     a = a1
    ElseIf sdnum = 2 Then
     a = a2
    ElseIf sdnum = 3 Then
     a = a3
    ElseIf sdnum = 4 Then
     a = a4
    ElseIf sdnum = 5 Then
     a = a5
    ElseIf sdnum = 6 Then
     a = a6
    ElseIf sdnum = 7 Then
     a = a7
    ElseIf sdnum = 8 Then
     a = a8
    ElseIf sdnum = 9 Then
     a = a9
    End If
   sta(a) = 0
   End If
   If sdnum = 1 Then
    If allall = 0 Then
     allall = 1
    End If
    If yes = 0 Then
     yes = 1
    End If
    If nono = 0 Then
     nono = 1
    End If
    isthis(a) = (nono / allall) - (yes / nono) + (nono / yes) + (yes / allall)
   End If
  Next a
   If sdnum = 1 Then
    For c = 1 To 9
     If c = 1 Then
      maxmax = c
     ElseIf isthis(c) > isthis(maxmax) Then
      maxmax = c
     End If
    Next c
     If sta(maxmax) <> 0 Then
     For y = 1 To 9
      If sta(y) = 0 Then
       maxmax = y
      End If
     Next y
    End If
    allall = 0
    yes = 0
    nono = 0
   End If
 ElseIf checkres = 0 Then
  main = 0
  allall = allall + 1
 ElseIf checkres = 2 Then
  main = 0
  allall = allall + 1
  nono = nono + 1
 ElseIf checkres = 3 Then
  main = 0
  allall = allall + 1
  yes = yes + 1
 End If
End If
main = maxmax
End Function
Private Sub Image1_Click()
yes = 0: nono = 0: alllall = 0
Label1.Visible = False
Label2.Visible = False
Image1.Visible = False
Image2.Visible = True
Image3.Visible = True
Image4.Visible = True
Image5.Visible = True
Image6.Visible = True
Image7.Visible = True
Image8.Visible = True
Image9.Visible = True
Image10.Visible = True
For i = 1 To 9
 sta(i) = 0
Next i
End Sub
Private Sub Image2_Click()
Dim com As Integer, endend As Integer
Image11.Visible = True
sta(1) = 1
com = main(1, 1)
If com = 0 Then
 endend = check()
 showshow = endend2()
Else
 sta(com) = 2
 If com = 1 Then
  Image20.Visible = True
 ElseIf com = 2 Then
  Image21.Visible = True
 ElseIf com = 3 Then
  Image22.Visible = True
 ElseIf com = 4 Then
  Image23.Visible = True
 ElseIf com = 5 Then
  Image24.Visible = True
 ElseIf com = 6 Then
  Image25.Visible = True
 ElseIf com = 7 Then
  Image26.Visible = True
 ElseIf com = 8 Then
  Image27.Visible = True
 ElseIf com = 9 Then
  Image28.Visible = True
 End If
 endend = check()
 If endend = 3 Then
  showshow = endend2()
End If
End If
End Sub
Private Sub Image29_Click()
 Label3.Visible = False
 Image29.Visible = False
 Image10.Visible = True
 Image2.Visible = True
 Image3.Visible = True
 Image4.Visible = True
 Image5.Visible = True
 Image6.Visible = True
 Image7.Visible = True
 Image8.Visible = True
 Image9.Visible = True
 Label4.Visible = False
 Label5.Visible = False
 For i = 1 To 9
  sta(i) = 0
 Next i
End Sub
Private Sub Image3_Click()
Dim com As Integer, endend As Integer
Image12.Visible = True
sta(2) = 1
com = main(1, 1)
If com = 0 Then
 endend = check()
 showshow = endend2()
Else
 sta(com) = 2
 If com = 1 Then
  Image20.Visible = True
 ElseIf com = 2 Then
  Image21.Visible = True
 ElseIf com = 3 Then
  Image22.Visible = True
 ElseIf com = 4 Then
  Image23.Visible = True
 ElseIf com = 5 Then
  Image24.Visible = True
 ElseIf com = 6 Then
  Image25.Visible = True
 ElseIf com = 7 Then
  Image26.Visible = True
 ElseIf com = 8 Then
  Image27.Visible = True
 ElseIf com = 9 Then
  Image28.Visible = True
 End If
 endend = check()
 If endend = 3 Then
  showshow = endend2()
End If
End If
End Sub
Private Sub Image4_Click()
Dim com As Integer, endend As Integer
Image13.Visible = True
sta(3) = 1
com = main(1, 1)
If com = 0 Then
 endend = check()
 showshow = endend2()
Else
 sta(com) = 2
 If com = 1 Then
  Image20.Visible = True
 ElseIf com = 2 Then
  Image21.Visible = True
 ElseIf com = 3 Then
  Image22.Visible = True
 ElseIf com = 4 Then
  Image23.Visible = True
 ElseIf com = 5 Then
  Image24.Visible = True
 ElseIf com = 6 Then
  Image25.Visible = True
 ElseIf com = 7 Then
  Image26.Visible = True
 ElseIf com = 8 Then
  Image27.Visible = True
 ElseIf com = 9 Then
  Image28.Visible = True
 End If
 endend = check()
 If endend = 3 Then
  showshow = endend2()
End If
End If
End Sub
Private Sub Image5_Click()
Dim com As Integer, endend As Integer
Image14.Visible = True
sta(4) = 1
com = main(1, 1)
If com = 0 Then
 endend = check()
 showshow = endend2()
Else
 sta(com) = 2
If com = 1 Then
  Image20.Visible = True
 ElseIf com = 2 Then
  Image21.Visible = True
 ElseIf com = 3 Then
  Image22.Visible = True
 ElseIf com = 4 Then
  Image23.Visible = True
 ElseIf com = 5 Then
  Image24.Visible = True
 ElseIf com = 6 Then
  Image25.Visible = True
 ElseIf com = 7 Then
  Image26.Visible = True
 ElseIf com = 8 Then
  Image27.Visible = True
 ElseIf com = 9 Then
  Image28.Visible = True
 End If
 endend = check()
 If endend = 3 Then
  showshow = endend2()
End If
End If
End Sub
Private Sub Image6_Click()
Dim com As Integer, endend As Integer
Image15.Visible = True
sta(5) = 1
com = main(1, 1)
If com = 0 Then
 endend = check()
 showshow = endend2()
Else
 sta(com) = 2
 If com = 1 Then
  Image20.Visible = True
 ElseIf com = 2 Then
  Image21.Visible = True
 ElseIf com = 3 Then
  Image22.Visible = True
 ElseIf com = 4 Then
  Image23.Visible = True
 ElseIf com = 5 Then
  Image24.Visible = True
 ElseIf com = 6 Then
  Image25.Visible = True
 ElseIf com = 7 Then
  Image26.Visible = True
 ElseIf com = 8 Then
  Image27.Visible = True
 ElseIf com = 9 Then
  Image28.Visible = True
 End If
 endend = check()
 If endend = 3 Then
  showshow = endend2()
End If
End If
End Sub
Private Sub Image7_Click()
Dim com As Integer, endend As Integer
Image16.Visible = True
sta(6) = 1
com = main(1, 1)
If com = 0 Then
 endend = check()
 showshow = endend2()
Else
 sta(com) = 2
 If com = 1 Then
  Image20.Visible = True
 ElseIf com = 2 Then
  Image21.Visible = True
 ElseIf com = 3 Then
  Image22.Visible = True
 ElseIf com = 4 Then
  Image23.Visible = True
 ElseIf com = 5 Then
  Image24.Visible = True
 ElseIf com = 6 Then
  Image25.Visible = True
 ElseIf com = 7 Then
  Image26.Visible = True
 ElseIf com = 8 Then
  Image27.Visible = True
 ElseIf com = 9 Then
  Image28.Visible = True
 End If
 endend = check()
 If endend = 3 Then
  showshow = endend2()
End If
End If
End Sub
Private Sub Image8_Click()
Dim com As Integer, endend As Integer
Image17.Visible = True
sta(7) = 1
com = main(1, 1)
If com = 0 Then
 endend = check()
 showshow = endend2()
Else
 sta(com) = 2
 If com = 1 Then
  Image20.Visible = True
 ElseIf com = 2 Then
  Image21.Visible = True
 ElseIf com = 3 Then
  Image22.Visible = True
 ElseIf com = 4 Then
  Image23.Visible = True
 ElseIf com = 5 Then
  Image24.Visible = True
 ElseIf com = 6 Then
  Image25.Visible = True
 ElseIf com = 7 Then
  Image26.Visible = True
 ElseIf com = 8 Then
  Image27.Visible = True
 ElseIf com = 9 Then
  Image28.Visible = True
 End If
 endend = check()
 If endend = 3 Then
  showshow = endend2()
End If
End If
End Sub
Private Sub Image9_Click()
Dim com As Integer, endend As Integer
Image18.Visible = True
sta(8) = 1
com = main(1, 1)
If com = 0 Then
 endend = check()
 showshow = endend2()
Else
 sta(com) = 2
 If com = 1 Then
  Image20.Visible = True
 ElseIf com = 2 Then
  Image21.Visible = True
 ElseIf com = 3 Then
  Image22.Visible = True
 ElseIf com = 4 Then
  Image23.Visible = True
 ElseIf com = 5 Then
  Image24.Visible = True
 ElseIf com = 6 Then
  Image25.Visible = True
 ElseIf com = 7 Then
  Image26.Visible = True
 ElseIf com = 8 Then
  Image27.Visible = True
 ElseIf com = 9 Then
  Image28.Visible = True
 End If
 endend = check()
 If endend = 3 Then
  showshow = endend2()
End If
End If
End Sub
Private Sub Image10_Click()
Dim com As Integer, endend As Integer
Image19.Visible = True
sta(9) = 1
com = main(1, 1)
If com = 0 Then
 endend = check()
 showshow = endend2()
Else
 sta(com) = 2
 If com = 1 Then
  Image20.Visible = True
 ElseIf com = 2 Then
  Image21.Visible = True
 ElseIf com = 3 Then
  Image22.Visible = True
 ElseIf com = 4 Then
  Image23.Visible = True
 ElseIf com = 5 Then
  Image24.Visible = True
 ElseIf com = 6 Then
  Image25.Visible = True
 ElseIf com = 7 Then
  Image26.Visible = True
 ElseIf com = 8 Then
  Image27.Visible = True
 ElseIf com = 9 Then
  Image28.Visible = True
 End If
 endend = check()
 If endend = 3 Then
  showshow = endend2()
End If
End If
End Sub
Private Sub Timer1_Timer()
 Dim end3 As Integer
 end3 = endend3()
End Sub
