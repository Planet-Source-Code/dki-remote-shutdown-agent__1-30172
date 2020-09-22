VERSION 5.00
Begin VB.Form frmPing 
   BackColor       =   &H009D6453&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Ping"
   ClientHeight    =   4800
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4800
   ScaleWidth      =   4560
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      BackColor       =   &H009D6453&
      Caption         =   "Ping by HostName..."
      ForeColor       =   &H00FFFFFF&
      Height          =   4590
      Left            =   75
      TabIndex        =   8
      Top             =   90
      Width           =   4395
      Begin VB.TextBox Text4 
         Height          =   330
         Index           =   5
         Left            =   1800
         TabIndex        =   19
         TabStop         =   0   'False
         Top             =   3480
         Width           =   2490
      End
      Begin VB.TextBox Text4 
         Height          =   330
         Index           =   4
         Left            =   1800
         TabIndex        =   18
         TabStop         =   0   'False
         Top             =   3120
         Width           =   2490
      End
      Begin VB.TextBox Text4 
         Height          =   330
         Index           =   3
         Left            =   1800
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   2760
         Width           =   2490
      End
      Begin VB.TextBox Text4 
         Height          =   330
         Index           =   2
         Left            =   1800
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   2400
         Width           =   2490
      End
      Begin VB.TextBox Text4 
         Height          =   330
         Index           =   1
         Left            =   1800
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   2040
         Width           =   2490
      End
      Begin VB.TextBox Text4 
         Height          =   330
         Index           =   0
         Left            =   1800
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   1680
         Width           =   2490
      End
      Begin VB.TextBox Text3 
         Height          =   330
         Left            =   1800
         TabIndex        =   1
         Text            =   "Echo This"
         Top             =   615
         Width           =   2490
      End
      Begin VB.TextBox Text2 
         Height          =   330
         Left            =   1800
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   1305
         Width           =   2490
      End
      Begin VB.TextBox Text1 
         Height          =   330
         Left            =   1800
         TabIndex        =   0
         Text            =   "localhost"
         Top             =   240
         Width           =   2490
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Ping Host"
         Height          =   375
         Left            =   2610
         TabIndex        =   2
         Top             =   4110
         Width           =   1620
      End
      Begin VB.Label Label1 
         BackColor       =   &H009D6453&
         Caption         =   "Data Pointer"
         ForeColor       =   &H00FFFFFF&
         Height          =   240
         Index           =   8
         Left            =   120
         TabIndex        =   17
         Top             =   3570
         Width           =   1665
      End
      Begin VB.Label Label1 
         BackColor       =   &H009D6453&
         Caption         =   "Data Returned"
         ForeColor       =   &H00FFFFFF&
         Height          =   240
         Index           =   7
         Left            =   120
         TabIndex        =   16
         Top             =   3165
         Width           =   1665
      End
      Begin VB.Label Label1 
         BackColor       =   &H009D6453&
         Caption         =   "Data Packet Size"
         ForeColor       =   &H00FFFFFF&
         Height          =   240
         Index           =   6
         Left            =   135
         TabIndex        =   15
         Top             =   2805
         Width           =   1665
      End
      Begin VB.Label Label1 
         BackColor       =   &H009D6453&
         Caption         =   "Round Trip Time"
         ForeColor       =   &H00FFFFFF&
         Height          =   240
         Index           =   5
         Left            =   150
         TabIndex        =   14
         Top             =   2445
         Width           =   1665
      End
      Begin VB.Label Label1 
         BackColor       =   &H009D6453&
         Caption         =   "Address (dec)"
         ForeColor       =   &H00FFFFFF&
         Height          =   240
         Index           =   4
         Left            =   135
         TabIndex        =   13
         Top             =   2085
         Width           =   1665
      End
      Begin VB.Label Label1 
         BackColor       =   &H009D6453&
         Caption         =   "Return Status"
         ForeColor       =   &H00FFFFFF&
         Height          =   240
         Index           =   3
         Left            =   150
         TabIndex        =   12
         Top             =   1725
         Width           =   1665
      End
      Begin VB.Label Label1 
         BackColor       =   &H009D6453&
         Caption         =   "Data To Send"
         ForeColor       =   &H00FFFFFF&
         Height          =   240
         Index           =   2
         Left            =   150
         TabIndex        =   11
         Top             =   645
         Width           =   1665
      End
      Begin VB.Label Label1 
         BackColor       =   &H009D6453&
         Caption         =   "Resolved IP Address"
         ForeColor       =   &H00FFFFFF&
         Height          =   240
         Index           =   1
         Left            =   150
         TabIndex        =   10
         Top             =   1365
         Width           =   1665
      End
      Begin VB.Label Label1 
         BackColor       =   &H009D6453&
         Caption         =   "Hostname to Ping"
         ForeColor       =   &H00FFFFFF&
         Height          =   240
         Index           =   0
         Left            =   150
         TabIndex        =   9
         Top             =   270
         Width           =   1665
      End
   End
End
Attribute VB_Name = "frmPing"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'****************************************************
' Copyright www.mvps.org
'VB interface to ping a computer by hostname or IP address
'*****************************************************
Private Sub Command1_Click()

   Dim ECHO As ICMP_ECHO_REPLY
   Dim pos As Long
   Dim success As Long
   Dim sIPAddress As String
   
   If SocketsInitialize() Then
   
     'convert the host name into an IP address
      sIPAddress = GetIPFromHostName(Text1.text)
      Text2.text = sIPAddress
      
     'ping the ip passing the address, text
     'to use, and the ECHO structure
      success = Ping(sIPAddress, (Text3.text), ECHO)
      
     'display the results
      Text4(0).text = GetStatusCode(success)
      Text4(1) = ECHO.Address
      Text4(2) = ECHO.RoundTripTime & " ms"
      Text4(3) = ECHO.DataSize & " bytes"
      
      If Left$(ECHO.Data, 1) <> Chr$(0) Then
         pos = InStr(ECHO.Data, Chr$(0))
         Text4(4) = Left$(ECHO.Data, pos - 1)
      End If
   
      Text4(5) = ECHO.DataPointer
      
      SocketsCleanup
      
   Else
   
        MsgBox "Windows Sockets for 32 bit Windows " & _
               "environments is not successfully responding."
   
   End If
   
End Sub

Private Sub Form_Load()
Me.Icon = frmMain.Icon

End Sub
