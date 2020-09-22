Attribute VB_Name = "modTextManip"
Option Explicit

Function ReplaceText(text As String, charfind As _
String, charchange As String)
'I do not remember where this code came from

Dim lReplace As Long
Dim bSameLength As Boolean
Dim lLenFind As Long
Dim lLenReplace As Long
lLenFind = Len(charfind)
lLenReplace = Len(charchange)
bSameLength = (lLenFind = lLenReplace)
ReplaceText = text
lReplace = InStr(1, ReplaceText, charfind)
    Do While lReplace <> 0
         If bSameLength Then
          Mid$(ReplaceText, lReplace, lLenFind) = charchange
        Else
          ReplaceText = Left$(ReplaceText, lReplace - 1) & charchange & Mid$(ReplaceText, lReplace + lLenFind)
        End If
    lReplace = InStr(lReplace + lLenReplace, ReplaceText, charfind)
    Loop

End Function
Public Function Compress( _
                sExpression As String, _
                sCompress As String, _
                Optional Compare As VbCompareMethod = vbBinaryCompare _
                ) As String

' by Peter Nierop, pnierop.pnc@inter.nl.net, 20001217
'This routine will remove multiple occurences of text and replace them down to one occurence
'I use it to get rid of multiple spaces "    " > " "
  Dim aOrg() As Byte
  Dim aFind() As Byte
  Dim aOut() As Byte

  Dim lMaxOrg&, lCurOrg&, lMaxFind&, lCurFind&, lFind&, lCurOut&, lMaxOut&
  Dim lFoundTwice&, lCopy&, lComp&

  'prepare Original String
  If Len(sExpression) = 0 Then Exit Function
  aOrg = sExpression
  lMaxOrg = UBound(aOrg)

  'prepare Compressed Output
  ReDim aOut(lMaxOrg)

  'Character or String to find
  If Len(sCompress) = 0 Then
    Compress = sExpression
    Exit Function
  End If

  aFind = sCompress
  lMaxFind = UBound(aFind)

  ' if test capitals
  If Compare = vbBinaryCompare Then
    lComp = &HFF
  Else
    lComp = &HDF   'to uppercase
    For lCurFind = 0 To lMaxFind Step 2
      aFind(lCurFind) = aFind(lCurFind) And &HDF
    Next
    lCurFind = 0
  End If

  ' preload the first character to find
  lFind = aFind(0)

  '==========  With one character to find -> shorter loop =====================
  If lMaxFind = 1 Then

    ' step through lowest bytes of unicode string array
    For lCurOrg = 0 To lMaxOrg Step 2

      'look for match with find character
      If lFind = (aOrg(lCurOrg) And lComp) Then
        lFoundTwice = lFoundTwice + 1
      Else
        lFoundTwice = 0
      End If

      'copy only if not a second match
      If lFoundTwice < 2 Then
        aOut(lCurOut) = aOrg(lCurOrg)
        lCurOut = lCurOut + 2
      End If

    Next

  Else
  '============ Longer loop if multiple characters to find ======================

    ' step through lowest bytes of unicode string array
    For lCurOrg = 0 To lMaxOrg Step 2

      'look for match with current find character
      If lFind = (aOrg(lCurOrg) And lComp) Then

        lCurFind = lCurFind + 2
        ' if no more characters to test -> match with string happened
        If lCurFind >= lMaxFind Then
          lFoundTwice = lFoundTwice + 1
          lCurFind = 0  'and start over
        End If
        ' now load next character from string to find
        lFind = aFind(lCurFind)

      Else
        ' no match so clean up
        lFoundTwice = 0
        lCurFind = 0
        lFind = aFind(lCurFind)
      End If

      ' copy character in output
      aOut(lCurOut) = aOrg(lCurOrg)
      lCurOut = lCurOut + 2

      'reset pointer to copy to if we had a second match
      If lFoundTwice = 2 Then
        lFoundTwice = 1
        lCurOut = lCurOut - lMaxFind - 1
      End If


    Next

  End If

  'shorten compressed string to real length
  ReDim Preserve aOut(lCurOut - 1)
  Compress = aOut  'array to string conversion
End Function

Public Function ParseString(ByVal vsString As _
String, ByVal vsDelimiter As String, ByVal _
viNumber As Integer)

Dim iFoundat As Integer
Dim iFoundatold As Integer
Dim iCurrentSection As Integer
Dim sText As String
If Len(vsString) > 0 And InStr(vsString, vsDelimiter) > 0 And viNumber > 0 Then
  iFoundat = 1
  iFoundatold = 1
  Do While InStr(iFoundatold + 1, vsString, vsDelimiter) > 0
    iFoundatold = iFoundat
    iFoundat = InStr(iFoundat + 1, vsString, vsDelimiter)
    iCurrentSection = iCurrentSection + 1
  Loop
  If viNumber > iCurrentSection Then
    Exit Function
  End If
  iFoundat = 1
  iCurrentSection = 0
  Do
    iFoundatold = iFoundat
    iFoundat = InStr(iFoundat + 1, vsString, vsDelimiter)
    If Trim(sText) = "" Then
      sText = Mid(vsString, 1, iFoundat - 1)
      iCurrentSection = iCurrentSection + 1
    Else
      If iFoundat > 0 Then
        sText = Mid(vsString, iFoundatold + 1, (iFoundat - 1) - iFoundatold)
      Else
        sText = Mid(vsString, iFoundatold + 1)
      End If
      iCurrentSection = iCurrentSection + 1
    End If
    If iCurrentSection = viNumber Then
      ParseString = sText
      Exit Do
    End If
  Loop
End If
ParseString = sText
End Function


