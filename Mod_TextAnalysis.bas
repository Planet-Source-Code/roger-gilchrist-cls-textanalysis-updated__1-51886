Attribute VB_Name = "Mod_TextAnalysis"
Option Explicit
Public Function CharDesc(AsciiNo As Long) As String
'just a support to allow you to show low ascii characters with Code name
'Note: using ByVal means that the original string is not affected by case chages in the routines
'      and saves using a temporary variable while manipulating the data
  If AsciiNo > 32 Then
    CharDesc = Chr$(AsciiNo)
   Else
    CharDesc = Chr$(AsciiNo) & Array(" NUL", " SCH", " STX", " ETX", " EOI", " ENQ", " ACK", " BEL", " BS", " HT", " LF", " VT", " FF", " CR", " SCI", " SI", " SLE", " CS1", " DC2", " DC3", " DC4", " NAK", " SYN", " ETB", " CAN", " EM", " SUB", " ESC", " FS", " GS", " RS", " US", " SP")(AsciiNo)
  End If
End Function
Public Function CountCharSplit(ByVal strSearch As String, _
                                ByVal strFind As String, _
                                Optional ByVal CaseSensitive As Boolean = True) As Long
  If Len(strSearch) Then
'the missing error test (Thanks to Merlin for reminding me to check the extreme cases)
'without it a blank strSource produces a return of 1 for any strFind except a zero length one
'No need to check strFind a zero length always returns 0
    If CaseSensitive Then
      CountCharSplit = UBound(Split(strSearch, strFind))
     Else
      CountCharSplit = UBound(Split(LCase$(strSearch), LCase$(strFind)))
    End If
  End If
End Function
Public Function CountNumerals(strSearch As String) As Long
  Dim I As Long
  If Len(strSearch) Then
    For I = 0 To 9
      CountNumerals = CountNumerals + UBound(Split(strSearch, CStr(I)))
    Next I
  End If
End Function
Public Function CountPunctuation(strSearch As String, Optional bExcludeLayout As Boolean = False) As Long
'bExcludeLayout ignore characters below chr$(33) so no spaces, tabs etc
  Dim I       As Long
  Dim strChr  As String
  Dim StartAt As Long

  If Len(strSearch) Then
    If bExcludeLayout Then
      StartAt = 33
    End If
    For I = StartAt To 255
      strChr = Chr$(I)
      If Not IsNumeric(strChr) Then
        If UCase$(strChr) = LCase$(strChr) Then
          CountPunctuation = CountPunctuation + UBound(Split(strSearch, Chr$(I)))
        End If
      End If
    Next I
  End If
End Function
Public Function CountWholeWord(ByVal strSearch As String, ByVal strFind As String, Optional CaseSensitive As Boolean = True, Optional AcceptInternalPunctuation As String = vbNullString) As Long
' ignores embedded matches ie 'the' in 'there'
'CaseSensitive = True  'The' and 'the' don't count as same word
'CaseSensitive = False 'The' and 'the' count as same word
  Dim Words()  As String
  Dim I        As Long
  If Len(strSearch) Then
    Words() = Split(AlphaNumericOnlyString(strSearch, AcceptInternalPunctuation))
    For I = LBound(Words) To UBound(Words)
      If Words(I) = strFind Then
        CountWholeWord = CountWholeWord + 1
      End If
    Next I
  End If
End Function
Public Function CountWords(ByVal strSearch As String, Optional AcceptInternalPunctuation As String = vbNullString) As Long
  If Len(strSearch) Then
    CountWords = UBound(ArrayNoBlanks(Split(AlphaNumericOnlyString(strSearch, AcceptInternalPunctuation))))
  End If
End Function
Public Function UniqueWordArray(ByVal strSearch As String, Optional CaseSensitive As Boolean = True, Optional AcceptInternalPunctuation As String = vbNullString) As String()
'Returns a String Array of the Unique words in strSearch
'CaseSensitive = True  'The' and 'the' don't count as same word
'CaseSensitive = False 'The' and 'the' count as same word
  If Len(strSearch) Then
    If Not CaseSensitive Then
      strSearch = LCase$(strSearch)
    End If
    UniqueWordArray = ArrayNoBlanks(ArrayNoDuplicates(Split(AlphaNumericOnlyString(strSearch, AcceptInternalPunctuation))))
  End If
End Function
Public Function CountUniqueWords(ByVal strSearch As String, Optional CaseSensitive As Boolean = True, Optional AcceptInternalPunctuation As String = vbNullString) As Long
'CaseSensitive = True  'The' and 'the' don't count as same word
'CaseSensitive = False 'The' and 'the' count as same word
  If Len(strSearch) Then
    CountUniqueWords = UBound(UniqueWordArray(strSearch, CaseSensitive, AcceptInternalPunctuation)) + 1
  End If
End Function
Public Function ArrayNoDuplicates(arrIn As Variant) As String()
  Dim I As Long
  Dim J As Long
'Dim K As Long
'Dim ArrTmp As Variant
  For I = LBound(arrIn) To UBound(arrIn)
    If Len(arrIn(I)) Then
      For J = LBound(arrIn) To UBound(arrIn)
        If I <> J Then
          If arrIn(I) = arrIn(J) Then
            arrIn(J) = ""
          End If
        End If
      Next J
    End If
  Next I
  ArrayNoDuplicates = arrIn
End Function
Public Function ArrayNoBlanks(arrIn As Variant) As String()
  Dim I                       As Long
  Dim K                       As Long
  Dim ArrTmp()                As String
  ReDim ArrTmp(UBound(arrIn)) As String
  For I = LBound(arrIn) To UBound(arrIn)
    If Len(arrIn(I)) Then
      ArrTmp(K) = arrIn(I)
      K = K + 1
    End If
  Next I
  ReDim Preserve ArrTmp(K - 1) As String
  ArrayNoBlanks = ArrTmp
End Function
Private Function AlphaNumericOnlyString(ByVal strSearch As String, Optional AcceptInternalPunctuation As String = vbNullString) As String
'AcceptInternalPunctation = if you want to allow hyphanated words or underscore characters in words
'or any other embedded punctuation add them to this parameter
'NOTE this will also count free standing incidences of the characters so the count may be inaccurate.
'I have not fully implemented this in the code just put it there if you want it
  Dim I        As Long
  Dim Words()  As String
  If Len(strSearch) Then
    AlphaNumericOnlyString = strSearch
    For I = 1 To 255 'remove all characters that are not letters or numerals
      If Not IsNumeric(Chr$(I)) Then
        If Not LCase$(Chr$(I)) <> UCase$(Chr$(I)) Then
          If Len(AcceptInternalPunctuation) Then
            If InStr(AcceptInternalPunctuation, Chr$(I)) Then
              GoTo SkipAcceptable
            End If
          End If
          Words() = Split(AlphaNumericOnlyString, Chr$(I))
          AlphaNumericOnlyString = Join(Words())
SkipAcceptable:
        End If
      End If
    Next I
  End If
End Function
Private DummyToKeepDecCommentsInDeclarations        As Boolean

