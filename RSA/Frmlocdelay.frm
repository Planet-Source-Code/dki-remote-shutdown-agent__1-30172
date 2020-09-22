VERSION 5.00
Begin VB.Form frmLocDelay 
   ClientHeight    =   3780
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6135
   ClipControls    =   0   'False
   LinkTopic       =   "Form3"
   ScaleHeight     =   3780
   ScaleWidth      =   6135
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   5400
      Top             =   600
   End
   Begin VB.TextBox txtExplain 
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      BorderStyle     =   0  'None
      Height          =   1215
      Left            =   1035
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   480
      Width           =   4215
   End
   Begin VB.Frame Frame1 
      Caption         =   "Message"
      Height          =   1215
      Left            =   120
      TabIndex        =   2
      Top             =   2400
      Width           =   4095
      Begin VB.TextBox txtMessage 
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         BorderStyle     =   0  'None
         Height          =   735
         Left            =   240
         MultiLine       =   -1  'True
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   360
         Width           =   3615
      End
   End
   Begin VB.CommandButton cmdSkipDelay 
      Caption         =   "DO IT NOW !"
      Height          =   615
      Left            =   4680
      TabIndex        =   0
      Top             =   2400
      Width           =   1215
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   4680
      TabIndex        =   1
      Top             =   3240
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   75
      Top             =   420
      Width           =   480
   End
   Begin VB.Label lblInitiatedTime 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "initiated at (12:40:00pm)"
      Height          =   195
      Left            =   2340
      TabIndex        =   7
      Top             =   120
      Width           =   1695
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Timed Shutdown "
      Height          =   195
      Left            =   960
      TabIndex        =   6
      Top             =   120
      Width           =   1245
   End
   Begin VB.Label lblTimeLeft 
      Caption         =   "Label2"
      Height          =   195
      Left            =   4560
      TabIndex        =   5
      Top             =   1920
      Width           =   1215
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Time before shutdown:"
      Height          =   195
      Left            =   2760
      TabIndex        =   4
      Top             =   1920
      Width           =   1620
   End
End
Attribute VB_Name = "frmLocDelay"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim iHours          As Integer
Dim iMinutes        As Integer
Dim iSeconds        As Integer
Dim lTimeLeft       As Long  'Timeout value in seconds
Dim modStorage      As Integer

Private Sub cmdCancel_Click()
    frmMain.SSTabLocalRemote.Enabled = True
    frmMain.frameDelay.Enabled = True
    frmMain.Show
    Unload Me
End Sub

Private Sub cmdSkipDelay_Click()
lTimeLeft = 1 'The shutdown take about 1 second to activate so set
              'timer value and shutdown
End Sub

Private Sub Form_Load()
    Me.Icon = LoadResPicture("DELAYICON", vbResIcon)
    Image1.Picture = Me.Icon

   If frmMain.chkLocalForce.Value = 1 Then
       frmLocDelay.txtExplain.text = stxtForced
   Else
       frmLocDelay.txtExplain.text = stxtNoForced
   End If
   
   txtMessage.text = frmMain.txtShutdownMessage.text
   
   Select Case True
    Case frmMain.optLocShutdown(0).Value = True
         Label3.Caption = "Timed Shutdown"
    Case frmMain.optLocShutdown(1).Value = True
         Label3.Caption = "Timed Reboot"
    Case frmMain.optLocShutdown(2).Value = True
         Label3.Caption = "Timed Logoff"
    End Select
    
    Me.Caption = Label3 & "in Progress"
    lblInitiatedTime.Caption = "initiated at " & Time & " ."
    
    If CLng(sTimeLeft) > 0 Then
        lTimeLeft = 0 + CLng(sTimeLeft)
        iHours = Int(CLng(sTimeLeft) / 3600)
        modStorage = Int(CLng(sTimeLeft) Mod 3600)
        iMinutes = Int(modStorage / 60)
        iSeconds = Int(modStorage Mod 60)
    Else
        iHours = 0
        iMinutes = 0
        iSeconds = 1
        
    End If
   
   lblTimeLeft.Caption = Format$(iHours & ":" & iMinutes & ":" & iSeconds, "h:mm:ss")
   bShutdown = False
    
    
End Sub


Private Sub Form_Unload(Cancel As Integer)
frmMain.frameDelay.Enabled = True
frmMain.SSTabLocalRemote.Enabled = True
frmMain.cmdExecute.Enabled = True
Timer1.Enabled = False

'bShutdown = True
End Sub

Private Sub Timer1_Timer()
   'Because the process of initiating the shutdown and then closing
   'the program takes we initiate the process 1 second before 0
 If lTimeLeft > 1 Then
      lTimeLeft = lTimeLeft - 1
      iHours = Int(lTimeLeft / 3600)
      modStorage = Int(lTimeLeft Mod 3600)
      iMinutes = Int(modStorage / 60)
      iSeconds = Int(modStorage Mod 60)
 Else
    If lTimeLeft <= 1 Then
        iHours = 0
        iMinutes = 0
        iSeconds = 0
        Timer1.Enabled = False
        Select Case True
            Case frmMain.optLocShutdown(0).Value = True
                 ShutDownNT (CBool(frmMain.chkLocalForce.Value))
            Case frmMain.optLocShutdown(1).Value = True
                 RebootNT (CBool(frmMain.chkLocalForce.Value))
            Case frmMain.optLocShutdown(2).Value = True
                 LogOffNT (CBool(frmMain.chkLocalForce.Value))
        End Select
        Unload Me
     End If
  End If

lblTimeLeft.Caption = Format$(iHours & ":" & iMinutes & ":" & iSeconds, "h:mm:ss")
    
If frmLocDelay.WindowState = vbMinimized Then
    frmLocDelay.Caption = lblTimeLeft & " Remaining"
Else
    frmLocDelay.Caption = Label3 & " in Progress"
End If
End Sub
