VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frmNetSend 
   BackColor       =   &H009D6453&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "NetSend"
   ClientHeight    =   4695
   ClientLeft      =   30
   ClientTop       =   255
   ClientWidth     =   6540
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4695
   ScaleWidth      =   6540
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin ComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   315
      Left            =   0
      TabIndex        =   4
      Top             =   4380
      Width           =   6540
      _ExtentX        =   11536
      _ExtentY        =   556
      Style           =   1
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   1
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Object.Tag             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton cmdSend 
      Caption         =   "Send"
      Height          =   336
      Left            =   4620
      TabIndex        =   3
      Top             =   240
      Width           =   1260
   End
   Begin VB.TextBox txtMsg 
      Height          =   300
      Index           =   2
      Left            =   852
      TabIndex        =   1
      ToolTipText     =   "Enter a Server Name (Optional) "
      Top             =   456
      Width           =   2892
   End
   Begin VB.TextBox txtMsg 
      Height          =   300
      Index           =   1
      Left            =   852
      TabIndex        =   0
      ToolTipText     =   "Enter User Name or Workstation Name"
      Top             =   96
      Width           =   2892
   End
   Begin VB.TextBox txtMsg 
      Height          =   3120
      Index           =   0
      Left            =   852
      MultiLine       =   -1  'True
      TabIndex        =   2
      Text            =   "netsend.frx":0000
      Top             =   936
      Width           =   5544
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H009D6453&
      Caption         =   "Message:"
      ForeColor       =   &H00FFFFFF&
      Height          =   264
      Index           =   2
      Left            =   60
      TabIndex        =   7
      Top             =   936
      Width           =   744
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H009D6453&
      Caption         =   "From:"
      ForeColor       =   &H00FFFFFF&
      Height          =   264
      Index           =   1
      Left            =   60
      TabIndex        =   6
      Top             =   468
      Width           =   744
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H009D6453&
      Caption         =   "To:"
      ForeColor       =   &H00FFFFFF&
      Height          =   264
      Index           =   0
      Left            =   60
      TabIndex        =   5
      Top             =   144
      Width           =   744
   End
End
Attribute VB_Name = "frmNetSend"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private WithEvents mNetSend As clsNetSend
Attribute mNetSend.VB_VarHelpID = -1
'****************************************************************
' This is a frontend for the Netsend command line executable
' It was downloaded from the internet and I haven't done extensive
' testing on it to see if its bullet proof.
'*****************************************************************


Private Sub cmdSend_Click()
    StatusBar1.SimpleText = "Sending Message..."
    With mNetSend
        .Message = txtMsg(0)
        .SendTo = txtMsg(1)
        .SendFromServer = txtMsg(2)
        .NetSendMessage
    End With
End Sub


Private Sub Form_Load()

    ' See Declaration Section and clsNetSend class module for more info
    
    ' App uses the mNetSend_Sent & mNetSend_Error events
    ' See code in cmdSend_Click event
    Set mNetSend = New clsNetSend
    Me.Icon = frmMain.Icon
    
       
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    ' Clean Up
    Set mNetSend = Nothing
    Set frmNetSend = Nothing
End Sub


Private Sub mNetSend_Error(ByVal lError As Long, ByVal ErrorText As String)
    StatusBar1.SimpleText = "Error " & lError & " - " & mNetSend.ErrorText
End Sub
Private Sub mNetSend_Sent()
    StatusBar1.SimpleText = "Message Sent"
End Sub



