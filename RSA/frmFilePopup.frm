VERSION 5.00
Begin VB.Form frmFilePopup 
   BackColor       =   &H00FF00FF&
   BorderStyle     =   0  'None
   ClientHeight    =   2340
   ClientLeft      =   3135
   ClientTop       =   2985
   ClientWidth     =   1515
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   Picture         =   "frmFilePopup.frx":0000
   ScaleHeight     =   156
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   101
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdExit 
      BackColor       =   &H00FF00FF&
      DisabledPicture =   "frmFilePopup.frx":AB42
      DownPicture     =   "frmFilePopup.frx":CC04
      Height          =   510
      Left            =   30
      MaskColor       =   &H00FF00FF&
      Picture         =   "frmFilePopup.frx":EF56
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   1695
      Width           =   1380
   End
End
Attribute VB_Name = "frmFilePopup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim ShapeMenu As clsTransForm 'clsDrag 'make a reference to the class

Private Sub cmdExit_Click()
Unload Me
End Sub

Private Sub Form_Load()
Set ShapeMenu = New clsTransForm 'clsDrag 'instantiate the object from the class
ShapeMenu.ShapeMe Me, RGB(255, 0, 255), True, App.Path & "\Data\PopupRegionData.dat"
ShapeMenu.ShapeMe cmdExit, RGB(83, 100, 157), True, App.Path & "\Data\Button2RegionData.dat"
Me.Top = frmMain.Top + 135 'offset from main form
Me.Left = frmMain.Left + 270
Set ShapeMenu = Nothing
End Sub

Private Sub Form_Deactivate()
Unload Me
End Sub




