VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{753FEE6F-A545-4EAA-AAC8-87512ED29F21}#3.0#0"; "ccrpDtp6.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmMain 
   BackColor       =   &H00FF00FF&
   BorderStyle     =   0  'None
   ClientHeight    =   5835
   ClientLeft      =   2865
   ClientTop       =   2850
   ClientWidth     =   9270
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   Picture         =   "frmComputerSelector.frx":0000
   ScaleHeight     =   389
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   618
   ShowInTaskbar   =   0   'False
   Begin VB.Frame frameDelay 
      BackColor       =   &H009D6453&
      BorderStyle     =   0  'None
      Height          =   3195
      Left            =   5430
      TabIndex        =   27
      Top             =   525
      Width           =   3555
      Begin VB.TextBox txtCount 
         BackColor       =   &H009D6453&
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   285
         Locked          =   -1  'True
         TabIndex        =   28
         TabStop         =   0   'False
         Top             =   2010
         Width           =   3165
      End
      Begin VB.OptionButton optRelative 
         BackColor       =   &H00E6A997&
         Caption         =   "Relative"
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   2565
         TabIndex        =   18
         ToolTipText     =   "00:20:00 = 20 min from now"
         Top             =   675
         Width           =   915
      End
      Begin VB.OptionButton optAbsolute 
         BackColor       =   &H00E6A997&
         Caption         =   "Absolute"
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   2565
         TabIndex        =   19
         ToolTipText     =   "00:20:00 = 12:20 AM"
         Top             =   930
         Width           =   915
      End
      Begin VB.CheckBox chkNoDelay 
         BackColor       =   &H00E5A693&
         Caption         =   "No Delay or message"
         Height          =   240
         Left            =   285
         TabIndex        =   21
         Top             =   2355
         Width           =   2610
      End
      Begin CCRPDTP6.ccrpDtp ccrpDtp1 
         Height          =   375
         Left            =   750
         TabIndex        =   17
         Top             =   750
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         CustomFormat    =   "HH:MM:ss"
         CCRPVer         =   1
         Var             =   "frmComputerSelector.frx":AF122
         XD              =   "frmComputerSelector.frx":AF156
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Text            =   "12/31/01"
      End
      Begin RichTextLib.RichTextBox txtShutdownMessage 
         Height          =   735
         Left            =   285
         TabIndex        =   20
         Top             =   1230
         Width           =   3165
         _ExtentX        =   5583
         _ExtentY        =   1296
         _Version        =   393217
         TextRTF         =   $"frmComputerSelector.frx":AF1B2
      End
      Begin VB.Label lblDelayType 
         BackColor       =   &H00E6A997&
         Caption         =   "Delay"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   270
         TabIndex        =   29
         Top             =   840
         Width           =   465
      End
      Begin VB.Image Image1 
         Height          =   2745
         Left            =   15
         Picture         =   "frmComputerSelector.frx":AF27B
         Top             =   315
         Width           =   3585
      End
   End
   Begin VB.CommandButton cmdAbort 
      BackColor       =   &H009D6453&
      DisabledPicture =   "frmComputerSelector.frx":CF56D
      DownPicture     =   "frmComputerSelector.frx":D172F
      Height          =   510
      Left            =   1815
      MaskColor       =   &H00FF00FF&
      Picture         =   "frmComputerSelector.frx":D3C19
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   4080
      Width           =   1380
   End
   Begin VB.CommandButton cmdExit 
      BackColor       =   &H009D6453&
      DisabledPicture =   "frmComputerSelector.frx":D5DDB
      DownPicture     =   "frmComputerSelector.frx":D7E9D
      Height          =   525
      Left            =   3360
      MaskColor       =   &H00FF00FF&
      Picture         =   "frmComputerSelector.frx":DA1EF
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   4080
      Width           =   1410
   End
   Begin VB.CommandButton cmdExecute 
      BackColor       =   &H00FFFFFF&
      DisabledPicture =   "frmComputerSelector.frx":DC2B1
      DownPicture     =   "frmComputerSelector.frx":DE07B
      Height          =   555
      Left            =   345
      MaskColor       =   &H009D6453&
      Picture         =   "frmComputerSelector.frx":DFE45
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   4080
      Width           =   1410
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   6225
      Top             =   4620
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmComputerSelector.frx":E1F07
            Key             =   "NEWICON"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmComputerSelector.frx":E27E1
            Key             =   "BUSY"
         EndProperty
      EndProperty
   End
   Begin VB.Timer TimerLocalShutdown 
      Enabled         =   0   'False
      Interval        =   200
      Left            =   7230
      Top             =   4020
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   6180
      Top             =   4020
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DialogTitle     =   "Save Computer List"
      Filter          =   "*.lst"
   End
   Begin VB.Timer Timer 
      Interval        =   1000
      Left            =   6735
      Top             =   4035
   End
   Begin TabDlg.SSTab SSTabLocalRemote 
      Height          =   3300
      Left            =   255
      TabIndex        =   0
      Top             =   630
      Width           =   5115
      _ExtentX        =   9022
      _ExtentY        =   5821
      _Version        =   393216
      Tabs            =   2
      Tab             =   1
      TabsPerRow      =   2
      TabHeight       =   520
      BackColor       =   10314835
      TabCaption(0)   =   "Local"
      TabPicture(0)   =   "frmComputerSelector.frx":E2C33
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Frame3"
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Remote"
      TabPicture(1)   =   "frmComputerSelector.frx":E2C4F
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "Frame1"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
      Begin VB.Frame Frame1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   2745
         Left            =   60
         TabIndex        =   24
         Top             =   465
         Width           =   4965
         Begin VB.CheckBox chkRemForce 
            BackColor       =   &H80000004&
            Caption         =   "Force"
            ForeColor       =   &H00000000&
            Height          =   285
            Left            =   120
            TabIndex        =   15
            ToolTipText     =   "Force Applications to Close"
            Top             =   2415
            Width           =   1185
         End
         Begin VB.CheckBox chkRemReboot 
            BackColor       =   &H80000004&
            Caption         =   "Reboot"
            ForeColor       =   &H00000000&
            Height          =   285
            Left            =   3975
            TabIndex        =   16
            ToolTipText     =   "Shutdown and reboot"
            Top             =   2430
            Width           =   930
         End
         Begin VB.Frame Frame4 
            Appearance      =   0  'Flat
            BackColor       =   &H80000004&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   480
            Left            =   165
            TabIndex        =   25
            Top             =   150
            Width           =   3750
            Begin VB.CommandButton cmdremove 
               BackColor       =   &H80000004&
               Height          =   330
               Left            =   1200
               Style           =   1  'Graphical
               TabIndex        =   11
               ToolTipText     =   "Delete Selected Entries in List"
               Top             =   120
               UseMaskColor    =   -1  'True
               Width           =   330
            End
            Begin VB.CommandButton cmdload 
               BackColor       =   &H80000004&
               Height          =   330
               Left            =   480
               Style           =   1  'Graphical
               TabIndex        =   9
               ToolTipText     =   "Open Saved List"
               Top             =   120
               UseMaskColor    =   -1  'True
               Width           =   330
            End
            Begin VB.CommandButton cmdSave 
               BackColor       =   &H80000004&
               Height          =   330
               Left            =   840
               Style           =   1  'Graphical
               TabIndex        =   10
               ToolTipText     =   "Save Current List"
               Top             =   120
               UseMaskColor    =   -1  'True
               Width           =   330
            End
            Begin VB.CommandButton cmdSelect 
               BackColor       =   &H80000004&
               Height          =   330
               Left            =   105
               Style           =   1  'Graphical
               TabIndex        =   8
               ToolTipText     =   "New List"
               Top             =   120
               UseMaskColor    =   -1  'True
               Width           =   330
            End
            Begin VB.CommandButton cmdSortAZ 
               BackColor       =   &H80000004&
               Height          =   330
               Left            =   1560
               Style           =   1  'Graphical
               TabIndex        =   12
               ToolTipText     =   "Sort List A-Z"
               Top             =   120
               UseMaskColor    =   -1  'True
               Width           =   330
            End
            Begin VB.CommandButton cmdSortZA 
               BackColor       =   &H80000004&
               Height          =   330
               Left            =   1920
               Style           =   1  'Graphical
               TabIndex        =   13
               ToolTipText     =   "Sort List Z-A"
               Top             =   120
               UseMaskColor    =   -1  'True
               Width           =   330
            End
            Begin VB.CommandButton cmdClearListView 
               BackColor       =   &H80000004&
               Height          =   330
               Left            =   2280
               Style           =   1  'Graphical
               TabIndex        =   14
               ToolTipText     =   "Clear List Contents"
               Top             =   120
               UseMaskColor    =   -1  'True
               Width           =   330
            End
         End
         Begin MSComctlLib.ListView ListView1 
            Height          =   1620
            Left            =   15
            TabIndex        =   26
            TabStop         =   0   'False
            Top             =   750
            Width           =   4935
            _ExtentX        =   8705
            _ExtentY        =   2858
            View            =   3
            LabelEdit       =   1
            Sorted          =   -1  'True
            MultiSelect     =   -1  'True
            LabelWrap       =   0   'False
            HideSelection   =   -1  'True
            AllowReorder    =   -1  'True
            FullRowSelect   =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   1
            NumItems        =   3
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "Name"
               Object.Width           =   2249
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "Comment"
               Object.Width           =   3528
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Text            =   "Status"
               Object.Width           =   2734
            EndProperty
         End
      End
      Begin VB.Frame Frame3 
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   2490
         Left            =   -74790
         TabIndex        =   22
         Top             =   420
         Width           =   4635
         Begin VB.TextBox txtLocalMachineName 
            BackColor       =   &H00FFFFFF&
            Height          =   315
            Left            =   195
            Locked          =   -1  'True
            TabIndex        =   23
            TabStop         =   0   'False
            Top             =   405
            Width           =   2490
         End
         Begin VB.OptionButton optLocShutdown 
            BackColor       =   &H80000004&
            Caption         =   "Shutdown"
            ForeColor       =   &H00000000&
            Height          =   240
            Index           =   0
            Left            =   195
            TabIndex        =   4
            Top             =   975
            Value           =   -1  'True
            Width           =   1560
         End
         Begin VB.OptionButton optLocShutdown 
            BackColor       =   &H80000004&
            Caption         =   "Reboot"
            ForeColor       =   &H00000000&
            Height          =   240
            Index           =   1
            Left            =   195
            TabIndex        =   5
            Top             =   1260
            Width           =   1560
         End
         Begin VB.OptionButton optLocShutdown 
            BackColor       =   &H80000004&
            Caption         =   "Log off && Log on as a different user"
            ForeColor       =   &H00000000&
            Height          =   240
            Index           =   2
            Left            =   195
            TabIndex        =   6
            Top             =   1530
            Width           =   3435
         End
         Begin VB.CheckBox chkLocalForce 
            BackColor       =   &H80000004&
            Caption         =   "Force"
            ForeColor       =   &H00000000&
            Height          =   285
            Left            =   195
            TabIndex        =   7
            ToolTipText     =   "Check to Force Applications Closed CAUTION: User Will Loose Any Unsaved Data !"
            Top             =   1995
            Width           =   1185
         End
      End
   End
   Begin VB.Label lblStatus 
      BackColor       =   &H009D6453&
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   435
      TabIndex        =   30
      Top             =   5430
      Width           =   8490
   End
   Begin VB.Image mnuHelp 
      Height          =   330
      Left            =   6405
      Top             =   195
      Width           =   660
   End
   Begin VB.Image mnuTools 
      Height          =   330
      Left            =   4875
      Top             =   180
      Width           =   735
   End
   Begin VB.Image mnuOptions 
      Height          =   330
      Left            =   3255
      Top             =   195
      Width           =   990
   End
   Begin VB.Image mnuEdit 
      Height          =   330
      Left            =   2040
      Top             =   180
      Width           =   525
   End
   Begin VB.Image mnuFile 
      Height          =   330
      Left            =   525
      Top             =   195
      Width           =   525
   End
   Begin VB.Image mnuMinimize 
      Height          =   315
      Left            =   7920
      Top             =   180
      Width           =   225
   End
   Begin VB.Image mnuExit 
      Height          =   315
      Left            =   8220
      Top             =   180
      Width           =   225
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim iLinecount          As Integer
Dim iShutdownMode       As Integer
Dim bShutdownEnabled    As Boolean
Dim iShutdownIndex      As Integer
Dim MainKeyRoot         As String
Dim MainSubKey          As String
Dim BlankArray(2)       As Byte
Dim LocalRadiobuttons   As Variant
Dim wks                 As New CNetWksta
Dim ShapeTheForm        As clsTransForm 'clsDrag 'make a reference to the class
Dim reg                 As New RegistryRoutines
Private WithEvents gSysTray As clsSysTray
Attribute gSysTray.VB_VarHelpID = -1

Private Sub cmdClearListView_Click()
    ListView1.ListItems.Clear
End Sub

Private Sub cmdExecute_Click()
   Dim sTimeDelay As String
   sTimeDelay = Format(ccrpDtp1.Value, "hh:mm:ss AMPM")
   cmdExecute.Enabled = False
   SSTabLocalRemote.Enabled = False
   frameDelay.Enabled = False
  
    Select Case True
    
        Case SSTabLocalRemote.Tab = cLocal
            iShutdownMode = cLocal
                Select Case True
                    Case optLocShutdown.item(0).Value 'Local Shutdown
                         
                            If chkNoDelay.Value = vbChecked Then
                                 ShutDownNT chkLocalForce.Value
                            Else
                               cmdAbort.Enabled = True
                               If optAbsolute.Value = True Then
                                    If DateDiff("s", Format(Now, "hh:mm:ss AMPM"), sTimeDelay) < 0 Then
                                        'Shutdown is tommorrow sometime so subtract difference from 24 hours
                                        sTimeLeft = 86400 - DateDiff("s", sTimeDelay, Format(Now, "hh:mm:ss AMPM"))
                                    Else
                                        sTimeLeft = DateDiff("s", Format(Now, "hh:mm:ss AMPM"), sTimeDelay)
                                    End If
                               Else
                                  sTimeLeft = Format(sTimeDelay, "h") * 3600 + Format(sTimeDelay, "n") * 60 + Format(sTimeDelay, "s")
                               End If
                               frmMain.Hide
                               frmLocDelay.Show
                               
                            End If
                            
                    Case optLocShutdown.item(1).Value  'Local Reboot
                             
                            If chkNoDelay.Value = vbChecked Then
                                RebootNT chkLocalForce.Value
                            Else
                               If optAbsolute.Value = True Then
                                    If DateDiff("s", Format(Now, "hh:mm:ss AMPM"), sTimeDelay) < 0 Then
                                        'Reboot is tommorrow sometime so subtract difference from 24 hours
                                        sTimeLeft = 86400 - DateDiff("s", sTimeDelay, Format(Now, "hh:mm:ss AMPM"))
                                    Else
                                        sTimeLeft = DateDiff("s", Format(Now, "hh:mm:ss AMPM"), sTimeDelay)
                                    End If
                               Else
                                  sTimeLeft = Format(sTimeDelay, "h") * 3600 + Format(sTimeDelay, "n") * 60 + Format(sTimeDelay, "s")
                               End If
                               frmMain.Hide
                               frmLocDelay.Show
                            End If
                    Case optLocShutdown.item(2).Value  'Local Logoff
                             
                            If chkNoDelay.Value = vbChecked Then
                                LogOffNT chkLocalForce.Value
                            Else
                               If optAbsolute.Value = True Then
                                    If DateDiff("s", Format(Now, "hh:mm:ss AMPM"), sTimeDelay) < 0 Then
                                        'Logoff is tommorrow sometime so subtract difference from 24 hours
                                        sTimeLeft = 86400 - DateDiff("s", sTimeDelay, Format(Now, "hh:mm:ss AMPM"))
                                    Else
                                        sTimeLeft = DateDiff("s", Format(Now, "hh:mm:ss AMPM"), sTimeDelay)
                                    End If
                               Else
                                  sTimeLeft = Format(sTimeDelay, "h") * 3600 + Format(sTimeDelay, "n") * 60 + Format(sTimeDelay, "s")
                               End If
                               frmMain.Hide
                               frmLocDelay.Show
                            End If
                End Select
        
        
        Case SSTabLocalRemote.Tab = cRemote 'Remote Shutdown
            iShutdownMode = cRemote
            cmdAbort.Enabled = True
            
            If chkNoDelay.Value = vbChecked Then 'Shutdown NOW !
               For iShutdownIndex = 1 To ListView1.ListItems.Count Step 1
                    ListView1.ListItems.item(iShutdownIndex).SubItems(2) = TrimNull(NTShutdown(ListView1.ListItems.item(iShutdownIndex), chkRemForce, chkRemReboot, 0, txtShutdownMessage.text))
               Next
            Else
                If optAbsolute.Value = True Then 'Shutdown with Time Delay
                  'Determine if absolute time has already passed
                   If DateDiff("s", Format(Now, "hh:mm:ss AMPM"), sTimeDelay) < 0 Then
                     'Shutdown is tommorrow sometime so subtract difference from 24 hours
                      sTimeLeft = 86400 - DateDiff("s", sTimeDelay, Format(Now, "hh:mm:ss AMPM"))
                   Else
                      sTimeLeft = DateDiff("s", Format(Now, "hh:mm:ss AMPM"), sTimeDelay)
                   End If
                Else
                   sTimeLeft = Format(sTimeDelay, "h") * 3600 + Format(sTimeDelay, "n") * 60 + Format(sTimeDelay, "s")
                End If
               
                For iShutdownIndex = 1 To ListView1.ListItems.Count Step 1
                    ListView1.ListItems.item(iShutdownIndex).SubItems(2) = TrimNull(NTShutdown(ListView1.ListItems.item(iShutdownIndex), chkRemForce, chkRemReboot, CLng(sTimeLeft), txtShutdownMessage.text))
                Next
                   
           End If
                cmdExecute.Enabled = True
                SSTabLocalRemote.Enabled = True
                frameDelay.Enabled = True
    End Select
End Sub



Private Sub Form_Load()
Dim CurrentTime As Variant

Dim i As Integer
Set wks = New CNetWksta
Set ShapeTheForm = New clsTransForm 'instantiate the object from the class
Set gSysTray = New clsSysTray
Set reg = New RegistryRoutines
Set gSysTray.SourceWindow = Me
MainKeyRoot = "Software\Network Tools\Remote Shutdown"
MainSubKey = "Settings"


    
    'Form Transparency region files
    ShapeTheForm.ShapeMe Me, RGB(255, 0, 255), True, App.Path & "\Data\MainRegionData.dat"
    ShapeTheForm.ShapeMe cmdExecute, RGB(83, 100, 157), True, App.Path & "\Data\ExecuteRegionData.dat"
    ShapeTheForm.ShapeMe cmdExit, RGB(83, 100, 157), True, App.Path & "\Data\Button2RegionData.dat"
    ShapeTheForm.ShapeMe cmdAbort, RGB(83, 100, 157), True, App.Path & "\Data\abortRegionData.dat"
    gSysTray.ChangeIcon ImageList1.ListImages("NEWICON").Picture 'Set System tray icon

    cmdAbort.Enabled = False
    txtLocalMachineName.Enabled = False
    txtLocalMachineName.text = GetMyMachineName
    frmMain.lblStatus.Caption = wks.LogonDomain

    'Test Registry settings
    reg.hkey = HKEY_LOCAL_MACHINE
    reg.KeyRoot = MainKeyRoot
    reg.Subkey = MainSubKey
    
        If Not reg.KeyExists Then reg.CreateKey 'If settings doesn't exist create it
        Me.Top = reg.GetRegistryValue("Top", Me.Top)
        Me.Left = reg.GetRegistryValue("Left", Me.Left)
        LocalRadiobuttons = reg.GetRegistryValue("LocalShutdown", BlankArray())
        
        For i = 0 To UBound(LocalRadiobuttons) 'put data in the text box
            ReDim Preserve LocalRadiobuttons(UBound(LocalRadiobuttons))
            optLocShutdown.item(i).Value = LocalRadiobuttons(i)
        Next i
       
        chkLocalForce.Value = reg.GetRegistryValue("Local Force", vbUnchecked)
        chkRemForce.Value = reg.GetRegistryValue("Remote Force", vbUnchecked)
        chkRemReboot.Value = reg.GetRegistryValue("Remote Reboot", vbUnchecked)
        chkNoDelay.Value = reg.GetRegistryValue("No Delay", vbUnchecked)
        txtShutdownMessage.text = reg.GetRegistryValue("Message", "This Machine is being shut down. Please save your work and log off.")
        If reg.GetRegistryValue("Time Type", 1) = 1 Then 'Delay
            optRelative = True
            optAbsolute = False
            lblDelayType.Caption = "Delay"
            With ccrpDtp1
                .DisplayMode = dtpDisplayTime
                .TimeFormat = dtpTimeFormatHMS
            End With
        Else
            optRelative = False 'Time
            optAbsolute = True
            lblDelayType.Caption = "Time"
            With ccrpDtp1
                .DisplayMode = dtpDisplayTime
                .TimeFormat = dtpTimeFormatHMSAP
            End With
        End If
        CurrentTime = Format(Time, "Short Time")
        ccrpDtp1.Value = reg.GetRegistryValue("Time", CurrentTime)
        iLinecount = SendMessage(txtShutdownMessage.hwnd, EM_GETLINECOUNT, 0&, ByVal 0&)
        txtCount.text = "Chars: " & Len(txtShutdownMessage.text) & ", lines: " & iLinecount & " (max 3)"
    
    'Create a buffer
    sTempDir = String(100, Chr$(0))
    'Get the temporary path
    GetTempPath 100, sTempDir
    'strip the rest of the buffer
    sTempDir = Left$(sTempDir, InStr(sTempDir, Chr$(0)) - 1)

    'Load Button Images from Resource File
    cmdSelect.Picture = LoadResPicture("NEW", vbResBitmap)
    cmdload.Picture = LoadResPicture("OPEN", vbResBitmap)
    cmdSave.Picture = LoadResPicture("SAVE", vbResBitmap)
    cmdClearListView.Picture = LoadResPicture("CLEAR", vbResBitmap)
    cmdremove.Picture = LoadResPicture("DELETE", vbResBitmap)
    cmdSortAZ.Picture = LoadResPicture("SORTAZ", vbResBitmap)
    cmdSortZA.Picture = LoadResPicture("SORTZA", vbResBitmap)
    gSysTray.IconInSysTray 'Create System Tray Icon
    gSysTray.ToolTip = "Remote Shutdown Agent"
End Sub


Private Sub cmdAbort_Click()
    Dim i As Integer
    
    Select Case True
      Case iShutdownMode = cLocal  'Local Shutdown
        bShutdownEnabled = False
            frmLocDelay.Timer1.Enabled = False
            Unload frmLocDelay
      Case iShutdownMode = cRemote  'Remote Shutdown
        bShutdownEnabled = False
        For i = 1 To iShutdownIndex - 1 Step 1
         ListView1.ListItems.item(i).SubItems(2) = NTAbortRemoteShutdown(ListView1.ListItems.item(i).text)
        'Get index of current remote shutdown process
        'for loop abort each one
        'set status text
        Next
    End Select
   
   cmdExecute.Enabled = True
   SSTabLocalRemote.Enabled = True
   frameDelay.Enabled = True
End Sub

Private Sub cmdExit_Click()
Unload Me
End Sub

Private Sub cmdload_Click()

   ' CancelError is True.
   On Error GoTo ErrHandler
   With CommonDialog1
        'Set Path
        .InitDir = App.Path
        ' Set filters.
        .Filter = "Computer Lists (*.rsl)|*.rsl|"
        ' Specify default filter.
        .FilterIndex = 0
        ' Display the Open dialog box.
        .ShowOpen
   End With
   
   ' Call the open file procedure.
   OpenFile (CommonDialog1.FileName)
   Exit Sub

ErrHandler:
' User pressed Cancel button.
   Exit Sub
End Sub

Private Sub cmdremove_Click()
Dim iIndex As Integer
If ListView1.ListItems.Count = 0 Then Exit Sub
For iIndex = ListView1.ListItems.Count To 1 Step -1
    If ListView1.ListItems.item(iIndex).Selected = True Then
       ListView1.ListItems.Remove (iIndex)
    End If
Next iIndex
End Sub

Private Sub cmdSave_Click()
 ' CancelError is True.
   On Error GoTo ErrHandler
   With CommonDialog1
        'Set Path
        .InitDir = App.Path
        ' Set filters.
        .Filter = "Computer Lists (*.rsl)|*.rsl|"
        ' Specify default filter.
        .FilterIndex = 0
        ' Display the Save dialog box.
        .ShowSave
   End With
   
   ' Call the save file procedure.
   SaveFile (CommonDialog1.FileName)
   Exit Sub

ErrHandler:
' User pressed Cancel button.
   Exit Sub
End Sub

Private Sub cmdSelect_Click()
    frmComputersOnline.Show 1
End Sub

Private Sub cmdSortAZ_Click()
    ListView1.SortOrder = lvwAscending
End Sub

Private Sub cmdSortZA_Click()
    ListView1.SortOrder = lvwDescending
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ShapeTheForm.DragForm Me.hwnd, Button 'Allows moving of form
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    'Save Settings to Registry
    reg.SetRegistryValue "Top", Me.Top, REG_DWORD
    reg.SetRegistryValue "Left", Me.Left, REG_DWORD
    reg.SetRegistryValue "LocalShutdown", LocalRadiobuttons, REG_BINARY
    reg.SetRegistryValue "Local Force", chkLocalForce.Value, REG_DWORD
    reg.SetRegistryValue "Remote Force", chkRemForce.Value, REG_DWORD
    reg.SetRegistryValue "Remote Reboot", chkRemReboot.Value, REG_DWORD
    reg.SetRegistryValue "No Delay", chkNoDelay.Value, REG_DWORD
    reg.SetRegistryValue "Message", txtShutdownMessage.text, REG_SZ
    reg.SetRegistryValue "Time", ccrpDtp1.Value, REG_SZ
    reg.SetRegistryValue "No Delay", chkNoDelay.Value, REG_DWORD
    If optRelative.Value = True Then
       reg.SetRegistryValue "Time Type", 1, REG_DWORD
    Else
       reg.SetRegistryValue "Time Type", 0, REG_DWORD
    End If
    gSysTray.RemoveFromSysTray
End Sub

Private Sub Form_Resize()
    If Me.WindowState = vbMinimized Then
        gSysTray.MinToSysTray
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set wks = Nothing
    Set ShapeTheForm = Nothing
    Set gSysTray = Nothing
    Set reg = Nothing
End Sub
Private Sub ListView1_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
 'Don't change the sortorder the first time
 'Change sort order to opposite direction
 If ListView1.SortKey = ColumnHeader.index - 1 Then
    If ListView1.SortOrder = lvwAscending Then
        ListView1.SortOrder = lvwDescending
    Else
        ListView1.SortOrder = lvwAscending
    End If
 End If
 
ListView1.SortKey = ColumnHeader.index - 1

End Sub

Private Sub mnuAbout_Click()
    frmAbout.Show 1
End Sub

Private Sub mnuchangeDomain_Click()
    frmComputersOnline.Show 1
End Sub

Private Sub mnuEdit_Click()
    frmEditPopup.Show
End Sub

Private Sub mnuExit_Click()
    Unload Me
End Sub

Private Sub mnuInfo_Click()
    frmInfo.Show 1
End Sub

Private Sub mnuNetSend_Click()
    frmNetSend.Show
End Sub

Private Sub mnuPing_Click()
    frmPing.Show
End Sub

Private Sub mnuFile_Click()
    frmFilePopup.Show
End Sub

Private Sub mnuHelp_Click()
    frmHelpPopup.Show
End Sub

Private Sub mnuMinimize_Click()
    Me.WindowState = vbMinimized
End Sub

Private Sub mnuOptions_Click()
    frmOptionsPopup.Show
End Sub

Private Sub mnuTools_Click()
    frmToolsPopup.Show
End Sub

Private Sub optAbsolute_Click()
    With ccrpDtp1
        .DisplayMode = dtpDisplayTime
        .TimeFormat = dtpTimeFormatHMSAP
    End With
    lblDelayType.Caption = "Time"
End Sub

Private Sub optLocShutdown_Click(index As Integer)
LocalRadiobuttons(index) = optLocShutdown(index).Value
End Sub

Private Sub optRelative_Click()
    With ccrpDtp1
        .DisplayMode = dtpDisplayTime
        .TimeFormat = dtpTimeFormatHMS
    End With
    lblDelayType.Caption = "Delay"
End Sub


Private Sub TimerLocalShutdown_Timer()
'  While bShutdownEnabled = True
'
'  Wend
  bShutdownEnabled = False
  TimerLocalShutdown.Enabled = False
  
End Sub
