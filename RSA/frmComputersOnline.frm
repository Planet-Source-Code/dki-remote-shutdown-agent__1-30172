VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmComputersOnline 
   BackColor       =   &H00FF00FF&
   BorderStyle     =   0  'None
   ClientHeight    =   5850
   ClientLeft      =   120
   ClientTop       =   120
   ClientWidth     =   3870
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmComputersOnline.frx":0000
   ScaleHeight     =   5850
   ScaleWidth      =   3870
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdOk 
      DownPicture     =   "frmComputersOnline.frx":4AA92
      Height          =   555
      Left            =   2070
      Picture         =   "frmComputersOnline.frx":4CE44
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   600
      Width           =   1305
   End
   Begin VB.CommandButton cmdRescan 
      DownPicture     =   "frmComputersOnline.frx":4EF76
      Height          =   480
      Left            =   345
      Picture         =   "frmComputersOnline.frx":51178
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   630
      Width           =   1230
   End
   Begin VB.ComboBox cmbDomainList 
      Height          =   315
      Left            =   195
      Style           =   2  'Dropdown List
      TabIndex        =   2
      ToolTipText     =   "Click to Change Browse Domain"
      Top             =   1230
      Width           =   3405
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   3705
      Left            =   150
      TabIndex        =   3
      Top             =   1635
      Width           =   3465
      _ExtentX        =   6112
      _ExtentY        =   6535
      View            =   3
      LabelEdit       =   1
      Sorted          =   -1  'True
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Name"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Comment"
         Object.Width           =   5292
      EndProperty
   End
   Begin VB.Image cmdExit 
      Height          =   315
      Left            =   2925
      Top             =   135
      Width           =   255
   End
End
Attribute VB_Name = "frmComputersOnline"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'**********************************************************************
' Created By: Don Kiser
' Creation Date: 1/1/02
' Description:  When called this form will populate a listview with
'               all computers in the selected domain along with their
'               descriptions. There's
'               a routine to make sure that each value only gets put in the listview once
'               on the main form.
'**********************************************************************
Dim FilePath As String
Dim prevOrder As Integer
Dim Foundcomputers() As BrowseNetwork
Dim sFileName As String
Dim sRemotePath  As String
Dim c As Long
Dim wks As New CNetWksta
Dim ShapeTheForm As clsTransForm 'clsDrag 'make a reference to the class

Private Sub cmbDomainList_Click()
    ListServers SV_TYPE_ALL, cmbDomainList.text
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdRescan_Click()
    LoadControls
End Sub

Private Sub Form_Load()
Dim strret As String
Dim sRet As String
    
    Set wks = New CNetWksta
    Set ShapeTheForm = New clsTransForm 'instantiate the object from the class
        ShapeTheForm.ShapeMe Me, RGB(255, 0, 255), True, App.Path & "\Data\OnlineRegionData.dat"
        ShapeTheForm.ShapeMe cmdRescan, RGB(83, 100, 157), True, App.Path & "\Data\RescanRegionData.dat"
        ShapeTheForm.ShapeMe cmdOk, RGB(83, 100, 157), True, App.Path & "\Data\button1RegionData.dat"
           
    LoadControls

End Sub
Private Sub LoadControls()
Dim namespace As IADsContainer
Dim Domain As IADs
  cmbDomainList.Clear           'Clear contents
  ListView1.ListItems.Clear     'Clear contents

    'Loads Combobox with all the current domains
    Set namespace = GetObject("WinNT:")
    
    For Each Domain In namespace
        cmbDomainList.AddItem Domain.name
    Next
        cmbDomainList.ListIndex = 0
   'populate the listview with the first domain in the combo box
    ListServers SV_TYPE_ALL, cmbDomainList.text
End Sub

Private Sub vbAddFileItemViewToMainForm(index As Integer)
    Dim sComputerName   As String
    Dim itmX            As ListItem
    Dim itmY            As ListItem
    Dim i               As Integer
    Dim bExists         As Boolean
    'check to make sure it's not already in the list on the main form
    'If it is msg user and don't re-add it
      If frmMain.ListView1.ListItems.Count > 0 Then
        For i = 1 To frmMain.ListView1.ListItems.Count
            Select Case True
             
                Case ListView1.ListItems.item(index) = frmMain.ListView1.ListItems.item(i)
                   ' Already exists in list
                   MsgBox ListView1.ListItems.item(index) & " Already in List", vbOKOnly, "Duplicate Computer Found"
                   bExists = True
                Case bExists = False And ListView1.ListItems.item(index) <> frmMain.ListView1.ListItems.item(i) And frmMain.ListView1.ListItems.Count <> i And ListView1.ListItems.item(index) <> ""
                   ' Doesn't exist yet but may farther down
                   ' keep looking
                Case bExists = False And ListView1.ListItems.item(index) <> frmMain.ListView1.ListItems.item(i) And frmMain.ListView1.ListItems.Count = i And ListView1.ListItems.item(index) <> ""
                   ' Doesn't exist in list
                   ' Add it to list
                   Set itmX = ListView1.ListItems.item(index)
                   Set itmY = frmMain.ListView1.ListItems.Add(, , itmX)
                   itmY.SubItems(1) = itmX.SubItems(1)
            End Select
        Next
        
      Else
        Set itmX = ListView1.ListItems.item(index)
        Set itmY = frmMain.ListView1.ListItems.Add(, , itmX)
        itmY.SubItems(1) = itmX.SubItems(1)
      End If
  
    Set itmX = Nothing
    Set itmY = Nothing
  
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ShapeTheForm.DragForm Me.hwnd, Button
End Sub

Private Sub Form_Unload(Cancel As Integer)
    'clean up and exit
    On Error Resume Next
    Set ShapeTheForm = Nothing
End Sub

Private Sub ListView1_ColumnClick(ByVal ColumnHeader As ColumnHeader)
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

Private Sub cmdOK_Click()
'Go through the listview and check if item is selected
'if it is add it to the main form listview

Dim iIndex As Integer
    sDomainName = cmbDomainList.text
    frmMain.lblStatus.Caption = sDomainName
    If ListView1.ListItems.Count = 0 Then Exit Sub
    For iIndex = ListView1.ListItems.Count To 1 Step -1
        If ListView1.ListItems.item(iIndex).Selected = True Then
           vbAddFileItemViewToMainForm (iIndex)
        End If
    Next iIndex
    
    Unload Me
End Sub

Private Sub ListServers(lType As Long, sDomain As String)
Dim lvwitm As ListItem
Dim lReturn As Long
Dim Server_Info As Long
Dim lEntries As Long
Dim lTotal As Long
Dim lMax As Long
Dim vResume As Variant
Dim tServer_info_101 As SERVER_INFO_101
Dim sServer As String
Dim sDescription As String
Dim lServerInfo101StructPtr As Long
Dim X As Long, i As Long
Dim bBuffer(512) As Byte
    
    sDomain = StrConv(sDomain, vbUnicode)
    Me.ListView1.ListItems.Clear
        
    lReturn = NetServerEnum( _
        ByVal 0&, _
        101, _
        Server_Info, _
        lMax, _
        lEntries, _
        lTotal, _
        ByVal lType, _
        sDomain, _
        vResume)
    
    If lReturn <> 0 Then
        MsgBox "Error " + Str$(lReturn) + " when trying to obtain server List " + Str$(lTotal), vbOKOnly + vbExclamation
        Exit Sub
    End If
        
    X = 1
    lServerInfo101StructPtr = Server_Info
    
    Do While X <= lTotal
    
        RtlMoveMemory _
            tServer_info_101, _
            ByVal lServerInfo101StructPtr, _
            Len(tServer_info_101)
        
        lstrcpyW bBuffer(0), _
            tServer_info_101.ptr_name
            
        
        i = 0
        Do While bBuffer(i) <> 0
            sServer = sServer & _
                Chr$(bBuffer(i))
            i = i + 2
        Loop
               lstrcpyW bBuffer(0), _
            tServer_info_101.ptr_comment
            
        
        i = 0
        Do While bBuffer(i) <> 0
            sDescription = sDescription & _
                Chr$(bBuffer(i))
            i = i + 2
        Loop
        Set lvwitm = Me.ListView1.ListItems.Add(, , sServer)
            lvwitm.SubItems(1) = sDescription
            DoEvents
            Set lvwitm = Nothing
        X = X + 1
    sServer = ""
    sDescription = ""
        lServerInfo101StructPtr = _
            lServerInfo101StructPtr + _
            Len(tServer_info_101)
        
    Loop
    
    lReturn = NetApiBufferFree(Server_Info)
    

End Sub


