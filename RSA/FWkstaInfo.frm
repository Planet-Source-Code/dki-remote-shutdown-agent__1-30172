VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmInfo 
   BackColor       =   &H00800000&
   BorderStyle     =   0  'None
   ClientHeight    =   3975
   ClientLeft      =   8040
   ClientTop       =   1455
   ClientWidth     =   6300
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   FillColor       =   &H00FFFFFF&
   ForeColor       =   &H00FFFFFF&
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   Picture         =   "FWkstaInfo.frx":0000
   ScaleHeight     =   3975
   ScaleWidth      =   6300
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdOK 
      BackColor       =   &H00FF00FF&
      DownPicture     =   "FWkstaInfo.frx":52692
      Height          =   585
      Left            =   4755
      Picture         =   "FWkstaInfo.frx":552CC
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   555
      Width           =   1335
   End
   Begin MSComctlLib.ListView lstInfo 
      Height          =   3255
      Left            =   345
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   405
      Width           =   4290
      _ExtentX        =   7567
      _ExtentY        =   5741
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Item"
         Object.Width           =   4410
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Value"
         Object.Width           =   2540
      EndProperty
   End
End
Attribute VB_Name = "frmInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' *********************************************************
'  Copyright (C)1997, Karl E. Peterson
'  This routine will enumerate basic system information
'  about the machine. I use it to get the logged on domain
' *********************************************************
Option Explicit
Dim ShapeTheForm As clsTransForm 'clsDrag 'make a reference to the class

Private Sub cmdOK_Click()
  Unload Me
End Sub

Private Sub Form_Load()
Dim wks As New CNetWksta
Dim itmX As ListItem
Set ShapeTheForm = New clsTransForm 'instantiate the object from the class
ShapeTheForm.ShapeMe Me, RGB(255, 0, 255), True, App.Path & "\Data\AboutRegionData.dat"
ShapeTheForm.ShapeMe cmdOK, RGB(83, 100, 157), True, App.Path & "\Data\OKRegionData.dat"

   lstInfo.ListItems.Clear
   Set itmX = lstInfo.ListItems.Add(, , "WORKSTATION INFO")
   Set itmX = lstInfo.ListItems.Add(, , "Platform ID")
       itmX.SubItems(1) = wks.PlatformId
   Set itmX = lstInfo.ListItems.Add(, , "Machine")
       itmX.SubItems(1) = wks.ComputerName
   Set itmX = lstInfo.ListItems.Add(, , "Domain")
       itmX.SubItems(1) = wks.Domain
   Set itmX = lstInfo.ListItems.Add(, , "LanMan Version")
       itmX.SubItems(1) = wks.VerMajor & "." & wks.VerMinor
   Set itmX = lstInfo.ListItems.Add(, , "LanMan Root")
       itmX.SubItems(1) = wks.LanRoot
   Set itmX = lstInfo.ListItems.Add(, , "Logged-On Users")
       itmX.SubItems(1) = wks.LoggedOnUsers
   Set itmX = lstInfo.ListItems.Add(, , "")
   Set itmX = lstInfo.ListItems.Add(, , "WORKSTATION USER INFO")
   Set itmX = lstInfo.ListItems.Add(, , "User Name")
       itmX.SubItems(1) = wks.UserName
   Set itmX = lstInfo.ListItems.Add(, , "Logon Domain")
       itmX.SubItems(1) = wks.LogonDomain
   Set itmX = lstInfo.ListItems.Add(, , "Other Domains")
       itmX.SubItems(1) = wks.OtherDomains
   Set itmX = lstInfo.ListItems.Add(, , "Logon Server")
       itmX.SubItems(1) = wks.LogonServer
   Set itmX = Nothing
   Set ShapeTheForm = Nothing
   
End Sub





Private Sub lstInfo_DblClick()
 Dim wks As New CNetWksta
   Dim itmX                  As ListItem
   lstInfo.ListItems.Clear
   Set itmX = lstInfo.ListItems.Add(, , "WORKSTATION INFO")
   Set itmX = lstInfo.ListItems.Add(, , "Platform ID")
       itmX.SubItems(1) = wks.PlatformId
   Set itmX = lstInfo.ListItems.Add(, , "Machine")
       itmX.SubItems(1) = wks.ComputerName
   Set itmX = lstInfo.ListItems.Add(, , "Domain")
       itmX.SubItems(1) = wks.Domain
   Set itmX = lstInfo.ListItems.Add(, , "LanMan Version")
       itmX.SubItems(1) = wks.VerMajor & "." & wks.VerMinor
   Set itmX = lstInfo.ListItems.Add(, , "LanMan Root")
       itmX.SubItems(1) = wks.LanRoot
   Set itmX = lstInfo.ListItems.Add(, , "Logged-On Users")
       itmX.SubItems(1) = wks.LoggedOnUsers
   Set itmX = lstInfo.ListItems.Add(, , "")
   Set itmX = lstInfo.ListItems.Add(, , "WORKSTATION USER INFO")
   Set itmX = lstInfo.ListItems.Add(, , "User Name")
       itmX.SubItems(1) = wks.UserName
   Set itmX = lstInfo.ListItems.Add(, , "Logon Domain")
       itmX.SubItems(1) = wks.LogonDomain
   Set itmX = lstInfo.ListItems.Add(, , "Other Domains")
       itmX.SubItems(1) = wks.OtherDomains
   Set itmX = lstInfo.ListItems.Add(, , "Logon Server")
       itmX.SubItems(1) = wks.LogonServer
   Set itmX = Nothing
End Sub
