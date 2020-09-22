VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   0  'None
   ClientHeight    =   3975
   ClientLeft      =   2295
   ClientTop       =   1605
   ClientWidth     =   6300
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Frmabout.frx":0000
   ScaleHeight     =   2743.616
   ScaleMode       =   0  'User
   ScaleWidth      =   5916.026
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdOK 
      BackColor       =   &H00FF00FF&
      DownPicture     =   "Frmabout.frx":52692
      Height          =   585
      Left            =   4815
      Picture         =   "Frmabout.frx":552CC
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   2970
      Width           =   1335
   End
   Begin VB.TextBox txtDisclaimer 
      Appearance      =   0  'Flat
      BackColor       =   &H009D6453&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFFF&
      Height          =   945
      Left            =   495
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   2
      Top             =   2715
      Width           =   4125
   End
   Begin VB.PictureBox picIcon 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H009D6453&
      BorderStyle     =   0  'None
      FillColor       =   &H009D6453&
      ForeColor       =   &H009D6453&
      Height          =   480
      Left            =   5385
      Picture         =   "Frmabout.frx":57C4E
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   4
      Top             =   420
      Width           =   480
   End
   Begin VB.TextBox txtDescription 
      Appearance      =   0  'Flat
      BackColor       =   &H009D6453&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFFF&
      Height          =   1260
      Left            =   1065
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Top             =   1095
      Width           =   4125
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      BorderStyle     =   6  'Inside Solid
      Index           =   1
      X1              =   422.573
      X2              =   5647.457
      Y1              =   1718.642
      Y2              =   1718.642
   End
   Begin VB.Label lblTitle 
      BackColor       =   &H009D6453&
      Caption         =   "Application Title"
      ForeColor       =   &H00FFFFFF&
      Height          =   480
      Left            =   1050
      TabIndex        =   3
      Top             =   240
      Width           =   3885
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      Index           =   0
      X1              =   422.573
      X2              =   5633.371
      Y1              =   1718.642
      Y2              =   1718.642
   End
   Begin VB.Label lblVersion 
      BackColor       =   &H009D6453&
      Caption         =   "Version"
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Left            =   1050
      TabIndex        =   0
      Top             =   780
      Width           =   3885
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim ShapeTheForm As clsTransForm 'clsDrag 'make a reference to the class
Private Sub cmdOK_Click()
  Unload Me
End Sub


Private Sub Form_Load()
Set ShapeTheForm = New clsTransForm 'instantiate the object from the class
ShapeTheForm.ShapeMe Me, RGB(255, 0, 255), True, App.Path & "\Data\AboutRegionData.dat"
ShapeTheForm.ShapeMe cmdOK, RGB(83, 100, 157), True, App.Path & "\Data\OKRegionData.dat"

    lblVersion.Caption = "Version " & App.Major & "." & App.Minor & "." & App.Revision
    If lblVersion.Caption = "Version 1.0.0" Then
        lblVersion.Caption = lblVersion.Caption & " [ALPHA RELEASE]"
    End If
        
    lblTitle.Caption = App.Title
    txtDescription.text = App.FileDescription
    txtDisclaimer.text = App.LegalCopyright
    Set ShapeTheForm = Nothing
End Sub

