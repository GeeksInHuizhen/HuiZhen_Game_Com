VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "Form2"
   ClientHeight    =   5805
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   7950
   LinkTopic       =   "Form2"
   ScaleHeight     =   5805
   ScaleWidth      =   7950
   StartUpPosition =   3  '´°¿ÚÈ±Ê¡
   Begin VB.CommandButton Command4 
      Caption         =   "¹ºÂò"
      Height          =   615
      Left            =   5520
      TabIndex        =   4
      Top             =   3600
      Width           =   1095
   End
   Begin VB.CommandButton Command3 
      Caption         =   "¹ºÂò"
      Height          =   615
      Left            =   3840
      TabIndex        =   3
      Top             =   3600
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "¹ºÂò"
      Height          =   615
      Left            =   2280
      TabIndex        =   2
      Top             =   3600
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "¹ºÂò"
      Height          =   615
      Left            =   720
      TabIndex        =   1
      Top             =   3600
      Width           =   1215
   End
   Begin VB.Label Label4 
      Caption         =   "250"
      Height          =   255
      Index           =   3
      Left            =   5640
      TabIndex        =   10
      Top             =   3240
      Width           =   975
   End
   Begin VB.Label Label4 
      Caption         =   "600"
      Height          =   255
      Index           =   2
      Left            =   3960
      TabIndex        =   9
      Top             =   3240
      Width           =   975
   End
   Begin VB.Label Label4 
      Caption         =   "100"
      Height          =   255
      Index           =   1
      Left            =   2280
      TabIndex        =   8
      Top             =   3240
      Width           =   975
   End
   Begin VB.Label Label4 
      Caption         =   "800"
      Height          =   255
      Index           =   0
      Left            =   840
      TabIndex        =   7
      Top             =   3240
      Width           =   975
   End
   Begin VB.Label Label3 
      Caption         =   "money:"
      BeginProperty Font 
         Name            =   "ËÎÌå"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5640
      TabIndex        =   6
      Top             =   120
      Width           =   975
   End
   Begin VB.Label Label2 
      Caption         =   "Label2"
      BeginProperty Font 
         Name            =   "ËÎÌå"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   6720
      TabIndex        =   5
      Top             =   120
      Width           =   1215
   End
   Begin VB.Image Image4 
      Height          =   735
      Left            =   5520
      Picture         =   "Form2.frx":0000
      Stretch         =   -1  'True
      Top             =   2160
      Width           =   1095
   End
   Begin VB.Image Image3 
      Height          =   855
      Left            =   3960
      Picture         =   "Form2.frx":4AB2
      Stretch         =   -1  'True
      Top             =   2160
      Width           =   960
   End
   Begin VB.Image Image2 
      Height          =   915
      Left            =   2280
      Picture         =   "Form2.frx":C5E7
      Stretch         =   -1  'True
      Top             =   2160
      Width           =   1035
   End
   Begin VB.Image Image1 
      Height          =   855
      Left            =   1080
      Picture         =   "Form2.frx":1290F
      Stretch         =   -1  'True
      Top             =   2160
      Width           =   615
   End
   Begin VB.Label Label1 
      Caption         =   "ÉÌµê"
      BeginProperty Font 
         Name            =   "ËÎÌå"
         Size            =   42
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   3000
      TabIndex        =   0
      Top             =   240
      Width           =   1695
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
If Form1.money >= 800 Then
 Form1.sd = 30
 Form1.money = Form1.money - 800
 Label2.Caption = CStr(Form1.money)
End If
End Sub

Private Sub Command2_Click()
If Form1.money >= 100 Then
 Form1.zy = Form1.zy + 1
 Form1.money = Form1.money - 100
 Label2.Caption = CStr(Form1.money)
End If
End Sub

Private Sub Command3_Click()
If Form1.money >= 600 Then
 Form1.stq = Form1.stq * 2
 Form1.money = Form1.money - 600
 Label2.Caption = CStr(Form1.money)
End If
End Sub

Private Sub Command4_Click()
If Form1.money >= 250 Then
 Form1.Timer4.Interval = Form1.Timer4.Interval + 1000
 Form1.money = Form1.money - 250
 Label2.Caption = CStr(Form1.money)
End If
End Sub

Private Sub Form_Load()
Label2.Caption = CStr(Form1.money)
End Sub
