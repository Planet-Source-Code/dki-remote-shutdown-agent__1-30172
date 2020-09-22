VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Test String parsing Module"
   ClientHeight    =   4170
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8535
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4170
   ScaleWidth      =   8535
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   645
      Left            =   4500
      TabIndex        =   8
      Top             =   3330
      Width           =   1725
   End
   Begin VB.TextBox txtDelimitor 
      Height          =   285
      Left            =   1440
      TabIndex        =   6
      Top             =   720
      Width           =   2655
   End
   Begin VB.CommandButton cmdDone 
      Caption         =   "&Done"
      Default         =   -1  'True
      Height          =   615
      Left            =   6480
      TabIndex        =   4
      Top             =   2400
      Width           =   1815
   End
   Begin VB.ListBox lstResult 
      Height          =   1815
      Left            =   120
      TabIndex        =   3
      Top             =   1200
      Width           =   3975
   End
   Begin VB.TextBox txtInput 
      Height          =   285
      Left            =   1440
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Top             =   240
      Width           =   6855
   End
   Begin VB.CommandButton cmdParse 
      Caption         =   "&Parse the text"
      Height          =   615
      Left            =   4440
      TabIndex        =   0
      Top             =   2400
      Width           =   1815
   End
   Begin VB.Label lblNumber 
      Height          =   255
      Left            =   4440
      TabIndex        =   7
      Top             =   1320
      Width           =   3855
   End
   Begin VB.Label Label2 
      Caption         =   "Delimitor list:"
      Height          =   255
      Left            =   240
      TabIndex        =   5
      Top             =   720
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "Text to parse:"
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   240
      Width           =   6135
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Foundcomputers() As BrowseNetwork

Private Sub cmdDone_Click()
Unload Me
End Sub

Private Sub Command1_Click()
Dim sFileName As String

Me.MousePointer = vbHourglass
'Clean the resultlist
lstResult.Clear
Me.MousePointer = vbDefault

Dim sTextLine As String
Dim iFileRead As Integer    ' FreeFile variable
Dim sTemp As String
Dim iIndex As Integer
Dim lUpper As Long
Dim lcount As Long
sFileName = "C:\temp\net.txt"
iFileRead = FreeFile()  'Get next available file number
iIndex = 0
Open sFileName For Input As #iFileRead         'Global Const for File dest
Do While Not EOF(iFileRead)   ' Loop until end of file.
    iIndex = iIndex + 1
    ReDim Preserve Foundcomputers(iIndex)
    Line Input #iFileRead, sTextLine   ' Read line into variable.
    sTextLine = MakeCSV.Compress(sTextLine, " ", vbTextCompare) 'remove extra spaces
    sTextLine = MakeCSV.ReplaceText(sTextLine, " ", ",")        'make into a csv
    lstResult.AddItem (sTextLine)                               'display
    Foundcomputers(iIndex).sComputerName = ParseString(sTextLine, ",", 1)
    Foundcomputers(iIndex).sComment1 = ParseString(sTextLine, ",", 2)
    Foundcomputers(iIndex).sComment2 = ParseString(sTextLine, ",", 3)
    Foundcomputers(iIndex).sComment3 = ParseString(sTextLine, ",", 4)
    Foundcomputers(iIndex).sComment4 = ParseString(sTextLine, ",", 5)
    Foundcomputers(iIndex).sComment5 = ParseString(sTextLine, ",", 6)
    Foundcomputers(iIndex).sComment6 = ParseString(sTextLine, ",", 7)
    Debug.Print Foundcomputers(iIndex).sComputerName, Foundcomputers(iIndex).sComment1, Foundcomputers(iIndex).sComment2, Foundcomputers(iIndex).sComment3, Foundcomputers(iIndex).sComment4, Foundcomputers(iIndex).sComment5, Foundcomputers(iIndex).sComment6
Loop
Close #iFileRead   ' Close file.

'Kill sFileName
End Sub
