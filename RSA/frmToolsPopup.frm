VERSION 5.00
Begin VB.Form frmToolsPopup 
   BackColor       =   &H00FF00FF&
   BorderStyle     =   0  'None
   ClientHeight    =   2340
   ClientLeft      =   7545
   ClientTop       =   2985
   ClientWidth     =   1530
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   Picture         =   "frmToolsPopup.frx":0000
   ScaleHeight     =   156
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   102
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdPing 
      DownPicture     =   "frmToolsPopup.frx":AB42
      Height          =   555
      Left            =   60
      Picture         =   "frmToolsPopup.frx":CC74
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1230
      Width           =   1305
   End
   Begin VB.CommandButton cmdNetSend 
      DownPicture     =   "frmToolsPopup.frx":EDA6
      Height          =   555
      Left            =   60
      Picture         =   "frmToolsPopup.frx":10ED8
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   585
      Width           =   1305
   End
End
Attribute VB_Name = "frmToolsPopup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim ShapeMenu As clsTransForm 'clsDrag 'make a reference to the class


Private Sub cmdNetSend_Click()
Me.Hide
frmNetSend.Show 1
End Sub

Private Sub cmdPing_Click()
Me.Hide
frmPing.Show 1
End Sub

Private Sub Form_Load()
Set ShapeMenu = New clsTransForm 'clsDrag 'instantiate the object from the class
ShapeMenu.ShapeMe Me, RGB(255, 0, 255), True, App.Path & "\Data\PopupRegionData.dat"
ShapeMenu.ShapeMe cmdNetSend, RGB(83, 100, 157), True, App.Path & "\Data\Button1RegionData.dat"
ShapeMenu.ShapeMe cmdPing, RGB(83, 100, 157), True, App.Path & "\Data\Button1RegionData.dat"
Me.Top = frmMain.Top + 135 'offset from main form
Me.Left = frmMain.Left + 4680
Set ShapeMenu = Nothing
End Sub
Private Sub Form_Deactivate()
Unload Me
End Sub


