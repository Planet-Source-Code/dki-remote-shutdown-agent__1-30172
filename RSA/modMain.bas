Attribute VB_Name = "GeneralMod"
Option Explicit
Public Const sConstTimeType As String = 0
Public Const sConstDelayChkbox As String = 0
Public Declare Sub CoTaskMemFree Lib "ole32.dll" (ByVal pv As Long)
Public Declare Function GetTempPath Lib "Kernel32" Alias "GetTempPathA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function NetServerEnum Lib "netapi32.dll" (vServername As Any, ByVal lLevel As Long, vBufptr As Any, lPrefmaxlen As Long, lEntriesRead As Long, lTotalEntries As Long, vServerType As Any, ByVal sDomain As String, vResumeHandle As Any) As Long
Public Declare Sub RtlMoveMemory Lib "Kernel32" (dest As Any, Vsrc As Any, ByVal lSize&)
Public Declare Sub lstrcpyW Lib "Kernel32" (vDest As Any, ByVal sSrc As Any)
Public Declare Function NetApiBufferFree Lib "netapi32" (ByVal Buffer As Long) As Long
Public Const stxtNoForced = "Your Computer will be shut down when the timer runs out. Although a forced shut down is not in effect, it may be a good idea to save your work before the timer runs out or you MAY lose any unsaved data."
Public Const stxtForced = "Your Computer will be shut down when the timer runs out. A forced shut down IS in effect, it is a good idea to save your work before the timer runs out or you WILL lose any unsaved data."
Public sTimeLeft As String
Public bShutdown As Boolean
Public Const cRemote As Integer = 1
Public Const cLocal As Integer = 0
Public sDomainName  As String
Public sTempDir     As String
Public bSk
Public Const LB_SETTABSTOPS As Long = &H192
Public Const EM_GETLINECOUNT = &HBA
Public Type BrowseNetwork
    sComputerName As String
    sComment1 As String
    sComment2 As String
    sComment3 As String
    sComment4 As String
    sComment5 As String
    sComment6 As String
End Type

Public Type SERVER_INFO_101
    dw_platform_id As Long
    ptr_name As Long
    dw_ver_major As Long
    dw_ver_minor As Long
    dw_type As Long
    ptr_comment As Long
End Type
Public Const SV_TYPE_WORKSTATION = &H1
Public Const SV_TYPE_SERVER = &H2
Public Const SV_TYPE_SQLSERVER = &H4
Public Const SV_TYPE_DOMAIN_CTRL = &H8
Public Const SV_TYPE_DOMAIN_BACKUP = &H10
Public Const SV_TYPE_TIMESOURCE = &H20
Public Const SV_TYPE_AFP = &H40
Public Const SV_TYPE_NOVELL = &H80
Public Const SV_TYPE_NT = &H8000
Public Const SV_TYPE_ALL = &HFFFFFFFF


Public Function FileExists(FilePath$) As Boolean
    FileExists = (Dir(FilePath$) <> "")
End Function

Public Function DoTabs(lstListBox As ListBox, TabArray() As Long)
    'clear any existing tabs
    Call SendMessage(lstListBox.hwnd, LB_SETTABSTOPS, 0&, ByVal 0&)
    'set list tabstops
    Call SendMessage(lstListBox.hwnd, LB_SETTABSTOPS, _
    CLng(UBound(TabArray)) + 1, TabArray(0))
End Function


Public Function OpenFile(FileName As String)
Dim iFileRead As Integer
Dim sTextline As String
Dim itmX As ListItem
Dim sComputerName As String
Dim sComment As String

    iFileRead = FreeFile()  'Get next available file number
    Open FileName For Input As #iFileRead         'Global Const for File dest
        Do While Not EOF(iFileRead)
        Line Input #iFileRead, sTextline   ' Read line into variable.
            If sTextline <> "" Then
               Set itmX = frmMain.ListView1.ListItems.Add(, , UCase$(ParseString(sTextline, ",", 1)))
                   itmX.SubItems(1) = ParseString(sTextline, ",", 2)
                   Set itmX = Nothing
            End If
        Loop
    Close #iFileRead
End Function

Public Function SaveFile(FileName As String)
Dim iFileWrite As Integer
Dim sTextline As String
Dim sComputerName   As String
Dim itmX            As ListItem
Dim i               As Integer
  
iFileWrite = FreeFile()
Open FileName For Output As #iFileWrite
 
 If frmMain.ListView1.ListItems.Count > 0 Then
     For i = 1 To frmMain.ListView1.ListItems.Count
        Set itmX = frmMain.ListView1.ListItems.item(i)
        sTextline = Compress(itmX.text & "," & itmX.SubItems(1), " ")
        Print #iFileWrite, sTextline
        Set itmX = Nothing
    Next
     
 End If
 Close #iFileWrite
 
End Function

Sub Main()
    '**************************************************
    'Main startup allows us to process command line arguements
    'without loading any forms
    'Code is not currently implemented
    '**************************************************

 Dim sMachineName   As String
 Dim iDelay         As Integer
 Dim sMessage       As String
 Dim sListName      As String
 Dim sCommand       As String
 Dim bForce         As Boolean
 Dim bReboot        As Boolean
 Dim sLocalMode     As String
 Dim i              As Integer
 
 iDelay = 20
 Load frmMain
 frmMain.Icon = LoadResPicture("Main", vbResIcon)
 frmMain.Show
 
End Sub

Public Function TrimNull(item As String) As String

    Dim pos As Integer
    
    pos = InStr(item, Chr$(0))
    If pos Then item = Left$(item, pos - 1)

TrimNull = item
End Function
