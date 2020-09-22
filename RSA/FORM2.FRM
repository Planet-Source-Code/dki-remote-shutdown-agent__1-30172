VERSION 5.00
Begin VB.Form frmOnline 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form2"
   ClientHeight    =   5370
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5985
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5370
   ScaleWidth      =   5985
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox lstOnline 
      Height          =   5130
      Left            =   120
      MultiSelect     =   2  'Extended
      Sorted          =   -1  'True
      TabIndex        =   2
      ToolTipText     =   "Click Then Type To Search By Computer Name"
      Top             =   135
      Width           =   4455
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   420
      Left            =   4890
      TabIndex        =   1
      Top             =   780
      Width           =   1035
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   420
      Left            =   4890
      TabIndex        =   0
      Top             =   255
      Width           =   1035
   End
End
Attribute VB_Name = "frmOnline"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Tabs(1) As Long
Dim Foundcomputers() As BrowseNetwork
Dim sFilename As String
Dim sRemotePath  As String
Dim c As Long
Dim rcText As RECT
Dim newWidth As Long
Dim itemWidth As Long
Dim sysScrollWidth As Long
  
Sub ParseFile(Filename As String)
Dim iFile As Integer
Dim sTextline As String
Dim sFilename As String
Dim iIndex As Integer
Dim stemp As String
Dim iCount As Integer
Dim str As String
Static x As Long
Dim c As Long


iFile = FreeFile()

Open Filename For Input As #iFile
iIndex = 0
Do While Not EOF(iFile)   ' Loop until end of file.
   
   
   Line Input #iFile, sTextline   ' Read line into variable.
   sTextline = Compress(sTextline, " ")
   sTextline = ReplaceText(sTextline, " ", ",")
   If InStr(1, sTextline, "\\") > 0 Then
        sTextline = Right(sTextline, Len(sTextline) - 2)
        iIndex = iIndex + 1
        ReDim Preserve Foundcomputers(iIndex)
        stemp = ParseString(sTextline, ",", 1)
        Foundcomputers(iIndex).sComputerName = stemp
        stemp = ParseString(sTextline, ",", 2)
        Foundcomputers(iIndex).sComment1 = stemp
        stemp = ParseString(sTextline, ",", 3)
        Foundcomputers(iIndex).sComment2 = stemp
        stemp = ParseString(sTextline, ",", 4)
        Foundcomputers(iIndex).sComment3 = stemp
        stemp = ParseString(sTextline, ",", 5)
        Foundcomputers(iIndex).sComment4 = stemp
        stemp = ParseString(sTextline, ",", 6)
        Foundcomputers(iIndex).sComment5 = stemp
        stemp = ParseString(sTextline, ",", 7)
        Foundcomputers(iIndex).sComment6 = stemp
        lstOnline.AddItem Foundcomputers(iIndex).sComputerName & vbTab & Foundcomputers(iIndex).sComment1 & " " & Foundcomputers(iIndex).sComment2 & " " & Foundcomputers(iIndex).sComment3 & " " & Foundcomputers(iIndex).sComment4 & " " & Foundcomputers(iIndex).sComment5 & " " & Foundcomputers(iIndex).sComment6 & " "
   End If
Loop
Close #iFile
End Sub

Private Sub cmdCancel_Click()
Unload Me

End Sub

Private Sub cmdOK_Click()
Dim iIndex As Integer
'starting with the last element of the array and work our way backwards to 0
'the first element of the array representing the ListBox.
'Working backwards through the array is necessary because if we started
'from the 0 element of the array, Visual Basic would re-sequence the
'elements in the Array, and eventually we would attempt to remove an item
'number that no longer exists.

If lstOnline.ListIndex = -1 Then Exit Sub

For iIndex = lstOnline.ListCount - 1 To 0 Step -1

If lstOnline.Selected(iIndex) = True Then

frmMain.lstSelected.AddItem lstOnline.List(iIndex)

lstOnline.RemoveItem iIndex

End If

Next iIndex
   sysScrollWidth = GetSystemMetrics(SM_CXVSCROLL)
   For c = 0 To frmMain.lstSelected.ListCount - 1
  
      Call DrawText(frmMain.hDC, (frmMain.lstSelected.List(c)), -1&, rcText, DT_CALCRECT)
      itemWidth = rcText.Right + sysScrollWidth + Tabs(1) 'need to account for tab spacing
         
      If itemWidth >= newWidth Then
         newWidth = itemWidth
      End If
      
   Next
   Call SendMessage(frmMain.lstSelected.hwnd, LB_SETHORIZONTALEXTENT, newWidth, ByVal 0&)
   
Unload Me
End Sub

Sub NetView()
Dim filenum As Integer

Dim sBatContents
Dim sHostLoc As String
Dim sMsg As String

Dim dResults As Double

    On Error GoTo ErrLoad:
    ChDir App.Path
    sRemotePath = "c:\temp"
    sBatContents = "net view > " & sRemotePath & "\netview.txt" 'contents of batch file
            sHostLoc = "netview.bat" 'batch file name
            sHostLoc = sRemotePath & "\netview.bat"
            Screen.MousePointer = vbHourglass
            filenum = FreeFile
            sMsg = Dir$(sHostLoc)
            If UCase(sMsg) = UCase("netview.bat") Then
                'File exists, kill it
                Kill sHostLoc
                Open sHostLoc For Output As #filenum
            Else
                Open sHostLoc For Output As #filenum
            End If
                Print #filenum, sBatContents
                Close filenum

    ShellAndWait sHostLoc, vbHide, 3000
    sFilename = sRemotePath & "\netview.txt"
    Screen.MousePointer = vbDefault
    Me.Caption = lstOnline.ListCount & " Domain Users found"
    Exit Sub
ErrLoad:
    MsgBox Err.Description
    Screen.MousePointer = vbDefault

End Sub



Private Sub Form_Load()
Me.Icon = frmMain.Icon
Me.Caption = "Computers in Domain"
Tabs(0) = 0
Tabs(1) = 100
NetView
ParseFile (sRemotePath & "\netview.txt")
'assure that the form font is the same as the
'list font to assure the DrawText method
'calculates the width correctly.
   frmOnline.Font.Name = lstOnline.Font.Name
   frmOnline.Font.Bold = lstOnline.Font.Bold
   frmOnline.Font.Size = lstOnline.Font.Size
   
  'get the width of the system scrollbar
   sysScrollWidth = GetSystemMetrics(SM_CXVSCROLL)
   
  'loop through the list items, using DrawText
  'with DT_CALCRECT to determine the longest item.
   For c = 0 To lstOnline.ListCount - 1
   
      Call DrawText(frmOnline.hDC, (lstOnline.List(c)), -1&, rcText, DT_CALCRECT)
      
     'calc the required width to display the
     'widest list item by adding the rect
     'width needed to display the item
     'with the width of the system scroll bar
      itemWidth = rcText.Right + sysScrollWidth + Tabs(1) 'need to account for tab spacing
         
     'if this width is wider than a previous
     'value, save the longer width
      If itemWidth >= newWidth Then
         newWidth = itemWidth
      End If
      
   Next
   
  'add a horizontal scrollbar wide enough
  'to display the longest list item. If the
  'scrollbar is not needed, its not shown.
   Call SendMessage(lstOnline.hwnd, LB_SETHORIZONTALEXTENT, newWidth, ByVal 0&)
   

DoTabs lstOnline, Tabs
End Sub

