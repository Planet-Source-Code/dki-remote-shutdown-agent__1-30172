VERSION 5.00
Begin VB.Form frmOptionsPopup 
   BackColor       =   &H00FF00FF&
   BorderStyle     =   0  'None
   ClientHeight    =   2400
   ClientLeft      =   6075
   ClientTop       =   2985
   ClientWidth     =   1500
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   Picture         =   "frmOptionsPopup.frx":0000
   ScaleHeight     =   160
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   100
   ShowInTaskbar   =   0   'False
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H009D6453&
      BackStyle       =   0  'Transparent
      Caption         =   "Intentionally Blank"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   555
      Left            =   165
      TabIndex        =   0
      Top             =   1005
      Width           =   1095
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "frmOptionsPopup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim ShapeMenu As clsTransForm 'clsDrag 'make a reference to the class


Private Sub Form_Load()
Set ShapeMenu = New clsTransForm 'clsDrag 'instantiate the object from the class
ShapeMenu.ShapeMe Me, RGB(255, 0, 255), True, App.Path & "\Data\PopupRegionData.dat"
Me.Top = frmMain.Top + 135 'offset from main form
Me.Left = frmMain.Left + 3210
Set ShapeMenu = Nothing
End Sub
Private Sub Form_Deactivate()
Unload Me
End Sub

