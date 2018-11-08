VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "µÿ Û"
   ClientHeight    =   3030
   ClientLeft      =   -60
   ClientTop       =   270
   ClientWidth     =   4560
   LinkTopic       =   "Form2"
   ScaleHeight     =   3030
   ScaleWidth      =   4560
   Begin VB.CommandButton Command1 
      Caption         =   "¥Ú"
      BeginProperty Font 
         Name            =   "ÀŒÃÂ"
         Size            =   27.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2895
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4335
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Form2.Visible = False
End Sub
