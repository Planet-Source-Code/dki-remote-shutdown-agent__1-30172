Attribute VB_Name = "modShutdown"

' Shutdown Flags
Const EWX_LOGOFF = 0
Const EWX_SHUTDOWN = 1
Const EWX_REBOOT = 2
Const EWX_FORCE = 4

Private Const TOKEN_ASSIGN_PRIMARY = &H1
Private Const TOKEN_DUPLICATE = (&H2)
Private Const TOKEN_IMPERSONATE = (&H4)
Private Const TOKEN_QUERY = (&H8)
Private Const TOKEN_QUERY_SOURCE = (&H10)
Private Const TOKEN_ADJUST_PRIVILEGES = (&H20)
Private Const TOKEN_ADJUST_GROUPS = (&H40)
Private Const TOKEN_ADJUST_DEFAULT = (&H80)
Private Const SE_SHUTDOWN_NAME = "SeShutdownPrivilege"
Private Const SE_PRIVILEGE_ENABLED = &H2

Private Const READ_CONTROL = &H20000
Private Const STANDARD_RIGHTS_ALL = &H1F0000
Private Const STANDARD_RIGHTS_EXECUTE = (READ_CONTROL)
Private Const STANDARD_RIGHTS_READ = (READ_CONTROL)
Private Const STANDARD_RIGHTS_REQUIRED = &HF0000
Private Const STANDARD_RIGHTS_WRITE = (READ_CONTROL)
Private Const TOKEN_ALL_ACCESS = (STANDARD_RIGHTS_REQUIRED Or _
                        TOKEN_ASSIGN_PRIMARY Or _
                        TOKEN_DUPLICATE Or _
                        TOKEN_IMPERSONATE Or _
                        TOKEN_QUERY Or _
                        TOKEN_QUERY_SOURCE Or _
                        TOKEN_ADJUST_PRIVILEGES Or _
                        TOKEN_ADJUST_GROUPS Or _
                        TOKEN_ADJUST_DEFAULT)
Private Const TOKEN_READ = (STANDARD_RIGHTS_READ Or _
                        TOKEN_QUERY)
Private Const TOKEN_WRITE = (STANDARD_RIGHTS_WRITE Or _
                        TOKEN_ADJUST_PRIVILEGES Or _
                        TOKEN_ADJUST_GROUPS Or _
                        TOKEN_ADJUST_DEFAULT)
Private Const TOKEN_EXECUTE = (STANDARD_RIGHTS_EXECUTE)

Private Const TokenDefaultDacl = 6
Private Const TokenGroups = 2
Private Const TokenImpersonationLevel = 9
Private Const TokenOwner = 4
Private Const TokenPrimaryGroup = 5
Private Const TokenPrivileges = 3
Private Const TokenSource = 7
Private Const TokenStatistics = 10
Private Const TokenType = 8
Private Const TokenUser = 1

Const ANYSIZE_ARRAY = 1

Private Const VER_PLATFORM_WIN32_NT = 2
Private Const VER_PLATFORM_WIN32_WINDOWS = 1
Private Const VER_PLATFORM_WIN32s = 0
Private hToken As Long
Private tkpSaved As TOKEN_PRIVILEGES

' To Report API errors:
'Private Const FORMAT_MESSAGE_ALLOCATE_BUFFER = &H100
'Private Const FORMAT_MESSAGE_ARGUMENT_ARRAY = &H2000
'Private Const FORMAT_MESSAGE_FROM_HMODULE = &H800
'Private Const FORMAT_MESSAGE_FROM_STRING = &H400
Private Const FORMAT_MESSAGE_FROM_SYSTEM = &H1000
Private Const FORMAT_MESSAGE_IGNORE_INSERTS = &H200
'Private Const FORMAT_MESSAGE_MAX_WIDTH_MASK = &HFF
Private Declare Function FormatMessage Lib "kernel32" Alias "FormatMessageA" (ByVal dwFlags As Long, lpSource As Any, ByVal dwMessageId As Long, ByVal dwLanguageId As Long, ByVal lpBuffer As String, ByVal nSize As Long, Arguments As Long) As Long

Type OSVERSIONINFO
    dwOSVersionInfoSize As Long
    dwMajorVersion As Long
    dwMinorVersion As Long
    dwBuildNumber As Long
    dwPlatformId As Long
    szCSDVersion As String * 128
End Type
' ============================================================================================
' NT Only
Private Type LARGE_INTEGER
    LowPart As Long
    HighPart As Long
End Type

Private Type LUID
    LowPart As Long
    HighPart As Long
End Type

Private Type LUID_AND_ATTRIBUTES
    pLuid As LUID
    Attributes As Long
End Type

Private Type TOKEN_PRIVILEGES
    PrivilegeCount As Long
    Privileges(ANYSIZE_ARRAY) As LUID_AND_ATTRIBUTES
End Type
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Declare Function GetTokenInformation Lib "advapi32.dll" (ByVal TokenHandle As Long, TokenInformationClass As Integer, TokenInformation As Any, ByVal TokenInformationLength As Long, ReturnLength As Long) As Long
Private Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (lpVersionInformation As OSVERSIONINFO) As Long
Private Declare Function AbortSystemShutdown Lib "advapi32.dll" Alias "AbortSystemShutdownA" (ByVal lpMachineName As String) As Long
Private Declare Function InitiateSystemShutdown Lib "advapi32.dll" Alias "InitiateSystemShutdownA" (ByVal lpMachineName As String, ByVal lpMessage As String, ByVal dwTimeout As Long, ByVal bForceAppsClosed As Long, ByVal bRebootAfterShutdown As Long) As Long
Private Declare Function ExitWindowsEx Lib "user32" (ByVal uFlags As Long, ByVal dwReserved As Long) As Long
Private Declare Function OpenProcessToken Lib "advapi32.dll" (ByVal ProcessHandle As Long, ByVal DesiredAccess As Long, TokenHandle As Long) As Long
Private Declare Function GetCurrentProcess Lib "kernel32" () As Long
Private Declare Function LookupPrivilegeValue Lib "advapi32" Alias "LookupPrivilegeValueA" (ByVal lpSystemName As String, ByVal lpName As String, lpLuid As LUID) As Long
Private Declare Function AdjustTokenPrivileges Lib "advapi32.dll" (ByVal TokenHandle As Long, ByVal DisableAllPrivileges As Long, NewState As TOKEN_PRIVILEGES, ByVal BufferLength As Long, PreviousState As TOKEN_PRIVILEGES, ReturnLength As Long) As Long
Private Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Private Declare Function GetLastError Lib "kernel32" () As Long
Public Function NTShutdown( _
    Optional ByVal sMachineNetworkName As String = vbNullString, _
    Optional ByVal bForceAppsToClose As Boolean = False, _
    Optional ByVal bReboot As Boolean = False, _
    Optional ByVal lTimeOut As Long = -1, _
    Optional ByVal sMsg As String = "" _
    ) As String
    
Dim lR As Long
    
    If IsNT Then
    ' Make sure we have enabled the privilege to shutdown
    ' for this process if we're running NT:
        If Not (NTEnableShutDown(sMsg)) Then
            Exit Function 'User doesn't have permission to shut down
        End If
    
        ' This is the code to do a shutdown:
        lR = InitiateSystemShutdown(sMachineNetworkName, sMsg, lTimeOut, bForceAppsToClose, bReboot)
    
        If (lR = 0) Then
            If Err.LastDllError = 1115 Then
                frmMain.cmdAbort.Enabled = True
                frmMain.SSTabLocalRemote.Enabled = True
                NTShutdown = "Already Initiated"
            Else
                 'MsgBox "Error Initiating Shutdown", vbOKOnly + vbInformation, "Error"
                 'Err.Raise eeSSDErrorBase + 2, App.EXEName & ".mShutDown", "InitiateSystemShutdown failed: " & WinError(Err.LastDllError)
                 NTShutdown = WinError(Err.LastDllError)
            End If
        End If
    Else
       MsgBox "Function only available under Windows NT.", vbOKOnly + vbInformation, "Wrong OS"
       NTShutdown = "Failed"
    End If
    
    If lR = 1 Then
        If lTimeOut <> 0 Then
            NTShutdown = "Initiated"
        Else
            NTShutdown = "Success"
        End If
    End If
    
End Function

Public Function NTAbortRemoteShutdown( _
                Optional ByVal sMachineNetworkName As String = vbNullString)
    
    Dim lR As Long
        
    lR = AbortSystemShutdown(sMachineNetworkName)
    
    If (lR = 0) Then
       'Err.Raise eeSSDErrorBase + 2, App.EXEName & ".mShutDown", "InitiateSystemShutdown failed: " & WinError(Err.LastDllError)
       NTAbortRemoteShutdown = WinError(Err.LastDllError)
    Else
       NTAbortRemoteShutdown = "Aborted"
    End If
    
End Function
Public Function WinError(ByVal lLastDLLError As Long) As String
Dim sBuff As String
Dim lCount As Long
    
    ' Return the error message associated with LastDLLError:
    sBuff = String$(256, 0)
    lCount = FormatMessage( _
    FORMAT_MESSAGE_FROM_SYSTEM Or FORMAT_MESSAGE_IGNORE_INSERTS, _
    0, lLastDLLError, 0&, sBuff, Len(sBuff), ByVal 0)
    sBuff = TrimNull$(sBuff)
    WinError = RTrim(sBuff)
    
End Function

Public Function IsNT() As Boolean
Static bOnce As Boolean
Static bValue As Boolean
    ' Return whether the system is running NT or not:
    If Not (bOnce) Then
    Dim tVI As OSVERSIONINFO
    tVI.dwOSVersionInfoSize = Len(tVI)
    If (GetVersionEx(tVI) <> 0) Then
        bValue = (tVI.dwPlatformId = VER_PLATFORM_WIN32_NT)
        bOnce = True
    End If
    End If
    IsNT = bValue
End Function
'set the shut down privilege for the current application
Private Function NTEnableShutDown(ByRef sMsg As String) As Boolean
Dim tLUID As LUID
Dim hProcess As Long
Dim hToken As Long
Dim tTP As TOKEN_PRIVILEGES, tTPOld As TOKEN_PRIVILEGES
Dim lTpOld As Long
Dim lR As Long

    ' Under NT we must enable the SE_SHUTDOWN_NAME privilege in the
    ' process we're trying to shutdown from, otherwise a call to
    ' try to shutdown has no effect!

    ' Find the LUID of the Shutdown privilege token:
    lR = LookupPrivilegeValue(vbNullString, SE_SHUTDOWN_NAME, tLUID)
    
    ' If we get it:
    If (lR <> 0) Then
                
    ' Get the current process handle:
    hProcess = GetCurrentProcess()
    If (hProcess <> 0) Then
        ' Open the token for adjusting and querying (if we can - user may not have rights):
        lR = OpenProcessToken(hProcess, TOKEN_ADJUST_PRIVILEGES Or TOKEN_QUERY, hToken)
        If (lR <> 0) Then
                    
            ' Ok we can now adjust the shutdown priviledges:
            With tTP
                .PrivilegeCount = 1
                With .Privileges(0)
                .Attributes = SE_PRIVILEGE_ENABLED
                .pLuid.HighPart = tLUID.HighPart
                .pLuid.LowPart = tLUID.LowPart
                End With
            End With
            
            ' Now allow this process to shutdown the system:
            lR = AdjustTokenPrivileges(hToken, 0, tTP, Len(tTP), tTPOld, lTpOld)
            
            If (lR <> 0) Then
                NTEnableShutDown = True
            Else
                
                'Err.Raise eeSSDErrorBase + 6, App.EXEName & ".mShutDown", "Can't enable shutdown: You do not have the privileges to shutdown this system. [" & WinError(Err.LastDllError) & "]"
                WinError (Err.LastDllError)
            End If
            
            ' Remember to close the handle when finished with it:
            CloseHandle hToken
        Else
            'Err.Raise eeSSDErrorBase + 6, App.EXEName & ".mShutDown", "Can't enable shutdown: You do not have the privileges to shutdown this system. [" & WinError(Err.LastDllError) & "]"
            WinError (Err.LastDllError)
        End If
    Else
        'Err.Raise eeSSDErrorBase + 5, App.EXEName & ".mShutDown", "Can't enable shutdown: Can't determine the current process. [" & WinError(Err.LastDllError) & "]"
        WinError (Err.LastDllError)
    End If
    Else
        'Err.Raise eeSSDErrorBase + 4, App.EXEName & ".mShutDown", "Can't enable shutdown: Can't find the SE_SHUTDOWN_NAME privilege value. [" & WinError(Err.LastDllError) & "]"
        WinError (Err.LastDllError)
    End If

End Function
       
' Shut Down NT
Public Sub ShutDownNT(Force As Boolean)
    Dim ret As Long
    Dim Flags As Long
    Flags = EWX_SHUTDOWN
    If Force Then Flags = Flags + EWX_FORCE
    'If IsNT Then NTEnableShutDown "Local Shutdown"
    ExitWindowsEx Flags, 0
End Sub
'Restart NT
Public Sub RebootNT(Force As Boolean)
    Dim ret As Long
    Dim Flags As Long
    Flags = EWX_REBOOT
    If Force Then Flags = Flags + EWX_FORCE
    'If IsNT Then NTEnableShutDown "Local Reboot"
    ExitWindowsEx Flags, 0
End Sub
'Log off the current user
Public Sub LogOffNT(Force As Boolean)
    Dim ret As Long
    Dim Flags As Long
    Flags = EWX_LOGOFF
    If Force Then Flags = Flags + EWX_FORCE
    ExitWindowsEx Flags, 0
End Sub

Function GetMyMachineName() As String
    Dim sLen As Long
    'create a buffer
    GetMyMachineName = Space(100)
    sLen = 100
    'retrieve the computer name
    If GetComputerName(GetMyMachineName, sLen) Then
        GetMyMachineName = Left(GetMyMachineName, sLen)
    End If
End Function

