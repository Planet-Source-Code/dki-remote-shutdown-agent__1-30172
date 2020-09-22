VERSION 5.00
Begin VB.Form frmHelpPopup 
   BackColor       =   &H00FF00FF&
   BorderStyle     =   0  'None
   ClientHeight    =   2310
   ClientLeft      =   9015
   ClientTop       =   2985
   ClientWidth     =   1485
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   Picture         =   "frmHelpPopup.frx":0000
   ScaleHeight     =   154
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   99
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdInfo 
      DownPicture     =   "frmHelpPopup.frx":AB42
      Height          =   555
      Left            =   60
      Picture         =   "frmHelpPopup.frx":CC74
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1245
      Width           =   1305
   End
   Begin VB.CommandButton cmdAbout 
      DownPicture     =   "frmHelpPopup.frx":EDA6
      Height          =   555
      Left            =   60
      Picture         =   "frmHelpPopup.frx":10ED8
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   600
      Width           =   1305
   End
End
Attribute VB_Name = "frmHelpPopup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim ShapeMenu As clsTransForm 'clsDrag 'make a reference to the class

Private Sub cmdabout_Click()
Me.Hide
frmAbout.Show 1

End Sub

Private Sub cmdInfo_Click()
Me.Hide
frmInfo.Show 1

End Sub

Private Sub Form_Load()
Set ShapeMenu = New clsTransForm 'clsDrag 'instantiate the object from the class
ShapeMenu.ShapeMe Me, RGB(255, 0, 255), True, App.Path & "\Data\PopupRegionData.dat"
ShapeMenu.ShapeMe cmdInfo, RGB(83, 100, 157), True, App.Path & "\Data\Button1RegionData.dat"
ShapeMenu.ShapeMe cmdAbout, RGB(83, 100, 157), True, App.Path & "\Data\Button1RegionData.dat"
Me.Top = frmMain.Top + 135 'offset from main form
Me.Left = frmMain.Left + 6150
Set ShapeMenu = Nothing
End Sub
Private Sub Form_Deactivate()
Unload Me
End Sub


