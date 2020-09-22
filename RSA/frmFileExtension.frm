VERSION 5.00
Begin VB.Form frmFileExtension 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "frmFileExtension"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function fCreateShellLink Lib "STKIT432.DLL" (ByVal _
    lpstrFolderName As String, ByVal lpstrLinkName As String, ByVal _
    lpstrLinkPath As String, ByVal lpstrLinkArgs As String) As Long
Private Sub Form_Load()
Dim strString As String
    Dim lngDword As Long

'    If Command$ <> "%1" Then
'        MsgBox (Command$ & " is the file you need to open!"), vbInformation
'        'Add to Recent file folder
'        lReturn = fCreateShellLink("..\..\Recent", _
'        Command$, Command$, "")
'    End If
    'create an entry in the class key
    Call savestring(HKEY_CLASSES_ROOT, "\.rsl", "", "RSA Computer List")
    'content type
    Call savestring(HKEY_CLASSES_ROOT, "\.rsl", "Content Type", "text/plain")
    'name
    Call savestring(HKEY_CLASSES_ROOT, "\rslfile", "", "RSA Computer List")
    'edit flags
    Call SaveDword(HKEY_CLASSES_ROOT, "\rslfile", "EditFlags", "0000")
    'file's icon (can be an icon file, or an
    '     icon located within a dll file)
    Call savestring(HKEY_CLASSES_ROOT, "\rslfile\DefaultIcon", "", App.Path & "\RSA.ico")
    'Shell
    Call savestring(HKEY_CLASSES_ROOT, "\rslfile\Shell", "", "")
    'Shell Open
    Call savestring(HKEY_CLASSES_ROOT, "\rslfile\Shell\Open", "", "")
    'Shell open command
    Call savestring(HKEY_CLASSES_ROOT, "\rslfile\Shell\Open\command", "", App.Path & "\Remote Shutdown Agent.exe %1")

End Sub
