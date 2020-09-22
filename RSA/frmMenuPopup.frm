VERSION 5.00
Begin VB.Form Form2 
   BackColor       =   &H00FF00FF&
   BorderStyle     =   0  'None
   ClientHeight    =   2295
   ClientLeft      =   3600
   ClientTop       =   2460
   ClientWidth     =   1500
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmMenuPopup.frx":0000
   ScaleHeight     =   2295
   ScaleWidth      =   1500
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   225
      TabIndex        =   1
      Text            =   "0"
      Top             =   1440
      Width           =   645
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   315
      TabIndex        =   0
      Text            =   "0"
      Top             =   570
      Width           =   645
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim ShapeMenu As clsTransForm 'clsDrag 'make a reference to the class


Private Sub Form_Click()
Unload Me

End Sub

Private Sub Form_Load()
Set ShapeMenu = New clsTransForm 'clsDrag 'instantiate the object from the class
ShapeMenu.ShapeMe Me, RGB(255, 0, 255), False, App.Path & "\MenuRegionData.dat"
Text1.text = frmMain.Top
Text2.text = frmMain.Left

   HScroll1.Min = 0   ' Initialize scroll bar.
   HScroll1.Max = 1000
   HScroll1.LargeChange = 1
   HScroll1.SmallChange = 1

   VScroll1.Min = 0   ' Initialize scroll bar.
   VScroll1.Max = 1000
   VScroll1.LargeChange = 1
   VScroll1.SmallChange = 1

End Sub

Private Sub HScroll1_Change()
Text2.text = HScroll1.Value + CSng(Text2.text)
Me.Move CSng(Text2.text)
End Sub

Private Sub Text1_Change()
Me.Move CSng(Text2.text), CSng(Text1.text)
End Sub

Private Sub Text2_Change()
Me.Move CSng(Text2.text)
End Sub

Private Sub VScroll1_Change()
Text1.text = VScroll1.Value + CSng(Text1.text)

Me.Move CSng(Text2.text), CSng(Text1.text)
End Sub
