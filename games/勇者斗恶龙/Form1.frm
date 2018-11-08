VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   7110
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   9615
   LinkTopic       =   "Form1"
   ScaleHeight     =   7110
   ScaleWidth      =   9615
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command2 
      Caption         =   "冰箭"
      Height          =   495
      Left            =   3120
      TabIndex        =   4
      Top             =   6240
      Width           =   1215
   End
   Begin VB.ListBox List1 
      Height          =   4920
      Left            =   3600
      TabIndex        =   1
      Top             =   720
      Width           =   3255
   End
   Begin VB.CommandButton Command1 
      Caption         =   "攻击"
      Height          =   495
      Left            =   1440
      TabIndex        =   0
      Top             =   6240
      Width           =   1335
   End
   Begin VB.Label Label3 
      Caption         =   " 3/3"
      Height          =   255
      Left            =   3120
      TabIndex        =   5
      Top             =   6840
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "dragon:1000/1000"
      Height          =   375
      Left            =   1200
      TabIndex        =   3
      Top             =   1440
      Width           =   2055
   End
   Begin VB.Label Label1 
      Caption         =   "hero:100/100"
      Height          =   375
      Left            =   1200
      TabIndex        =   2
      Top             =   1920
      Width           =   2055
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim bjcn As Integer
Dim bd As Boolean

Private Sub Command1_Click()
Dim hero, dragon As Integer
Dim HP_h, HP_d As Integer
hero = Len(Label1.Caption)
HP_h = Val(Mid(Label1.Caption, 6, hero - 4))
dragon = Len(Label2.Caption)
HP_d = Val(Mid(Label2.Caption, 8, dragon - 5))
Dim ATK, DMG As Integer
Dim doubleATK, doubleDMG As Integer
Randomize
ATK = Int(Rnd * 10) + 50
DMG = Int(Rnd * 10) + 5
doubleATK = Int(Rnd * 10) - 4
doubleDMG = Int(Rnd * 100) - 5
If doubleATK < 0 Then
ATK = ATK * 2
HP_d = HP_d - ATK
Label2.Caption = "dragon:" + Str(HP_d) + "/1000"
List1.AddItem "勇者会心一击，对巨龙造成" + Str(ATK) + "点伤害！"
Else
HP_d = HP_d - ATK
Label2.Caption = "dragon:" + Str(HP_d) + "/1000"
List1.AddItem "勇者对巨龙造成" + Str(ATK) + "点伤害"
End If
If bd = False Then
If doubleDMG < 0 Then
DMG = DMG * 2
HP_h = HP_h - DMG
Label1.Caption = "hero:" + Str(HP_h) + "/100"
List1.AddItem "巨龙猛击勇者，对勇者造成" + Str(DMG) + "点伤害！"
Else
HP_h = HP_h - DMG
Label1.Caption = "hero:" + Str(HP_h) + "/100"
List1.AddItem "巨龙吐出烈焰，对勇者造成" + Str(DMG) + "点伤害"
End If
Else
List1.AddItem "巨龙被冻住了，没有攻击"
End If
bd = False
If Not bjcn = 0 Then
bjcn = bjcn - 1
End If
If HP_h <= 0 Or HP_d <= 0 Then
HP_h = 100
HP_d = 1000
MsgBox "game again"
bjcn = 0
Label1.Caption = "hero:" + Str(HP_h) + "/100"
Label2.Caption = "dragon:" + Str(HP_d) + "/1000"
List1.Clear
End If
Label3.Caption = Str(3 - bjcn) + "/3"
End Sub

Private Sub Command2_Click()
Dim hero, dragon As Integer
Dim HP_h, HP_d As Integer
hero = Len(Label1.Caption)
HP_h = Val(Mid(Label1.Caption, 6, hero - 4))
dragon = Len(Label2.Caption)
HP_d = Val(Mid(Label2.Caption, 8, dragon - 5))
Dim ATK, DMG As Integer
Dim doubleATK, doubleDMG As Integer
Randomize
ATK = Int(Rnd * 10) + 50
DMG = Int(Rnd * 10) + 5
doubleATK = Int(Rnd * 10) - 1
doubleDMG = Int(Rnd * 100) - 5
If bjcn = 0 Then
If doubleATK < 0 Then
ATK = ATK * 2
HP_d = HP_d - ATK
Label2.Caption = "dragon:" + Str(HP_d) + "/1000"
List1.AddItem "勇者冰箭打中巨龙要害，对巨龙造成" + Str(ATK) + "点伤害！"
Else
HP_d = HP_d - ATK
Label2.Caption = "dragon:" + Str(HP_d) + "/1000"
List1.AddItem "勇者冰箭射中巨龙,造成" + Str(ATK) + "点伤害"
End If
If doubleDMG < 0 Then
DMG = DMG * 2
HP_h = HP_h - DMG
Label1.Caption = "hero:" + Str(HP_h) + "/100"
List1.AddItem "巨龙猛击勇者，对勇者造成" + Str(DMG) + "点伤害！"
Else
HP_h = HP_h - DMG
Label1.Caption = "hero:" + Str(HP_h) + "/100"
List1.AddItem "巨龙吐出烈焰，对勇者造成" + Str(DMG) + "点伤害"
End If
bd = True
bjcn = 3
If HP_h <= 0 Or HP_d <= 0 Then
HP_h = 100
HP_d = 1000
MsgBox "game again"
bjcn = 0
Label1.Caption = "hero:" + Str(HP_h) + "/100"
Label2.Caption = "dragon:" + Str(HP_d) + "/1000"
List1.Clear
End If
Else
MsgBox "冰箭充能未完毕"
End If
Label3.Caption = Str(3 - bjcn) + "/3"
End Sub

Private Sub Form_Load()
List1.Clear
bjlq = 0
bd = False
End Sub
