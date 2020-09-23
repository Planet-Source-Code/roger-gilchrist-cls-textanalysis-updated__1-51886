VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cls_TextAnalysis"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private CFm_CaseSensitive                   As Boolean
Private CFm_SourceString                    As String
Private cfm_LCaseSourceString               As String
Private CFm_AlphaNumericOnlyString          As String
Private CFm_UniqueWordArray()               As String
Private CFm_UniqueWordArrayLCase()          As String
Private bHasWorkString                      As Boolean
Private CFm_AcceptInternalPunctuation       As String

Public Property Get AcceptableInternalPunctuation() As String

  AcceptableInternalPunctuation = CFm_AcceptInternalPunctuation

End Property

Public Property Let AcceptableInternalPunctuation(ByVal propVal As String)

  CFm_AcceptInternalPunctuation = propVal
'refresh the internal data
  SourceString = CFm_SourceString

End Property

Private Function AlphaNumericOnlyString(ByVal strSearch As String) As String

'AcceptInternalPunctation = if you want to allow hyphanated words or underscore characters in words
'or any other embedded punctuation add them to this parameter
'NOTE this will also count free standing incidences of the characters so the count may be inaccurate.
'I have not fully implemented this in the code just put it there if you want it

  Dim I        As Long

  If bHasWorkString Then
    AlphaNumericOnlyString = strSearch
    For I = 1 To 255 'remove all characters that are not letters or numerals
      If Not IsNumeric(Chr$(I)) Then
        If Not LCase$(Chr$(I)) <> UCase$(Chr$(I)) Then
          If Len(CFm_AcceptInternalPunctuation) Then
            If InStr(CFm_AcceptInternalPunctuation, Chr$(I)) Then
              GoTo SkipAcceptable
            End If
          End If
          AlphaNumericOnlyString = Join(Split(AlphaNumericOnlyString, Chr$(I)))
SkipAcceptable:
        End If
      End If
    Next I
  End If

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

Public Function ArrayNoDuplicates(arrIn As Variant) As String()

  Dim I As Long
  Dim J As Long

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

Public Property Get CaseSensitive() As Boolean

  CaseSensitive = CFm_CaseSensitive

End Property

Public Property Let CaseSensitive(ByVal propVal As Boolean)

  CFm_CaseSensitive = propVal
  SourceString = CFm_SourceString

End Property

Public Function CharDesc(AsciiNo As Long) As String

'just a support to allow you to show low ascii characters with Code name

  If AsciiNo > 32 Then
    CharDesc = Chr$(AsciiNo)
   Else
    CharDesc = Chr$(AsciiNo) & Array(" NUL", " SCH", " STX", " ETX", " EOI", " ENQ", " ACK", " BEL", " BS", " HT", " LF", " VT", " FF", " CR", " SCI", " SI", " SLE", " CS1", " DC2", " DC3", " DC4", " NAK", " SYN", " ETB", " CAN", " EM", " SUB", " ESC", " FS", " GS", " RS", " US", " SP")(AsciiNo)
  End If

End Function

Private Sub Class_Initialize()

  CFm_CaseSensitive = True

End Sub

Public Function CountChar(ByVal strFind As String, _
                          Optional ByVal CaseSensitive As Boolean = True) As Long

  If bHasWorkString Then
'Note: using ByVal means that the original string is not affected by case chages in the routines
'      and saves using a temporary variable while manipulating the data
'the missing error test (Thanks to Merlin for reminding me to check the extreme cases)
'without it a blank strSource produces a return of 1 for any strFind except a zero length one
'No need to check strFind a zero length always returns 0
    If CaseSensitive Then
      CountChar = UBound(Split(CFm_SourceString, strFind))
     Else
      CountChar = UBound(Split(cfm_LCaseSourceString, LCase$(strFind)))
    End If
  End If

End Function

Public Property Get CountFormat(Optional bExcludeSpace As Boolean = False) As Long

  CountFormat = CountPunctuation(False) - CountPunctuation()
  If bExcludeSpace Then
    CountFormat = CountFormat - CountChar(" ")
  End If

End Property

Public Function CountNumerals() As Long

  Dim I As Long

  If bHasWorkString Then
    For I = 0 To 9
      CountNumerals = CountNumerals + UBound(Split(cfm_LCaseSourceString, CStr(I)))
    Next I
  End If

End Function

Public Property Get CountPunctuation(Optional bExcludeLayout As Boolean = True) As Long

'bExcludeLayout ignore characters below chr$(33) so no spaces, tabs etc

  Dim I       As Long
  Dim StartAt As Long

  If bHasWorkString Then
    If bExcludeLayout Then
      StartAt = 33
    End If
    For I = StartAt To 255
      If IsPunctuation(Chr$(I)) Then
        CountPunctuation = CountPunctuation + UBound(Split(CFm_SourceString, Chr$(I)))
      End If
    Next I
  End If

End Property

Public Property Get CountUCase() As Long

'Not used in demo just to show how to build it

  Dim I    As Long

  For I = 1 To Len(CFm_SourceString)
    If IsUpperAlpha(Mid$(CFm_SourceString, I, 1)) Then
      CountUCase = CountUCase + 1
    End If
  Next I

End Property

Public Function CountUniqueWords(Optional bForceCaseAwareness As Boolean = True) As Long

'CaseSensitive = True  'The' and 'the' don't count as same word
'CaseSensitive = False 'The' and 'the' count as same word

  If bHasWorkString Then
    If bForceCaseAwareness Then
      CountUniqueWords = UBound(CFm_UniqueWordArray) + 1
     Else
      CountUniqueWords = UBound(CFm_UniqueWordArrayLCase) + 1
    End If
  End If

End Function

Public Function CountWholeWord(ByVal strFind As String) As Long

' ignores embedded matches ie 'the' in 'there'
'CaseSensitive = True  'The' and 'the' don't count as same word
'CaseSensitive = False 'The' and 'the' count as same word

  Dim Words()  As String
  Dim I        As Long

  If Len(cfm_LCaseSourceString) Then
    Words() = Split(CFm_AlphaNumericOnlyString)
    For I = LBound(Words) To UBound(Words)
      If Words(I) = strFind Then
        CountWholeWord = CountWholeWord + 1
      End If
    Next I
  End If

End Function

Public Function CountWords() As Long

  If bHasWorkString Then
    CountWords = UBound(ArrayNoBlanks(Split(CFm_AlphaNumericOnlyString))) + 1
'Thanks to Min Thant Sin for finding this bug
'forgot about zero based array needing the + 1
'bHasWorkString stops the procedure hitting if the answer is 0
  End If

End Function

Private Sub GenerateAlphaNumericOnlyString()

'AcceptInternalPunctation = if you want to allow hyphanated words or underscore characters in words
'or any other embedded punctuation add them to this parameter
'NOTE this will also count free standing incidences of the characters so the count may be inaccurate.
'I have not fully implemented this in the code just put it there if you want it

  Dim I        As Long

  If bHasWorkString Then
    CFm_AlphaNumericOnlyString = cfm_LCaseSourceString
    For I = 1 To 255 'remove all characters that are not letters or numerals
      If Not IsNumeric(Chr$(I)) Then
        If Not LCase$(Chr$(I)) <> UCase$(Chr$(I)) Then
          If Len(CFm_AcceptInternalPunctuation) Then
            If InStr(CFm_AcceptInternalPunctuation, Chr$(I)) Then
              GoTo SkipAcceptable
            End If
          End If
          CFm_AlphaNumericOnlyString = Join(Split(CFm_AlphaNumericOnlyString, Chr$(I)))
SkipAcceptable:
        End If
      End If
    Next I
  End If

End Sub

Private Sub GenerateUniqueWordArray()

  CFm_UniqueWordArray = ArrayNoBlanks(ArrayNoDuplicates(Split(AlphaNumericOnlyString(CFm_SourceString))))
  CFm_UniqueWordArrayLCase = ArrayNoBlanks(ArrayNoDuplicates(Split(AlphaNumericOnlyString(cfm_LCaseSourceString))))

End Sub

Public Function HasChar(ByVal strFind As String, _
                        Optional ByVal CaseSensitive As Boolean = True) As Long

  If bHasWorkString Then
'the missing error test (Thanks to Merlin for reminding me to check the extreme cases)
'without it a blank strSource produces a return of 1 for any strFind except a zero length one
'No need to check strFind a zero length always returns 0
    If CaseSensitive Then
      HasChar = InStr(CFm_SourceString, strFind)
     Else
      HasChar = InStr(cfm_LCaseSourceString, LCase$(strFind))
    End If
  End If

End Function

Public Function IsAlpha(strChr As String) As Boolean

  IsAlpha = LCase$(strChr) <> UCase$(strChr)

End Function

Private Function IsPunctuation(ByVal strChr As String) As Boolean

  If Not IsNumeric(strChr) Then
    If UCase$(strChr) = LCase$(strChr) Then
      IsPunctuation = True
    End If
  End If

End Function

Public Function IsUpper(strChr As String) As Boolean

  IsUpper = strChr = UCase$(strChr)

End Function

Private Function IsUpperAlpha(strChr As String) As Boolean

  If IsAlpha(strChr) Then
    IsUpperAlpha = IsUpper(strChr)
  End If

End Function

Public Property Get SourceString() As String

  SourceString = CFm_SourceString

End Property

Public Property Let SourceString(ByVal propVal As String)

  CFm_SourceString = propVal
  cfm_LCaseSourceString = LCase$(CFm_SourceString)
  bHasWorkString = Len(cfm_LCaseSourceString)
  If bHasWorkString Then
'Thanks to Min Thant Sin for finding this bug
'crashed if you deleted the whole text
    GenerateUniqueWordArray
    GenerateAlphaNumericOnlyString
  End If

End Property

Public Property Get UniqueWordArray(Optional CaseSenitive As Boolean = True) As String()

'Returns a String Array of the Unique words in strSearch

  If bHasWorkString Then
    If CaseSenitive Then
      UniqueWordArray = CFm_UniqueWordArray
     Else
      UniqueWordArray = CFm_UniqueWordArrayLCase
    End If
  End If

End Property

':)Roja's VB Code Fixer V1.1.93 (23/02/2004 4:16:29 PM) 9 + 346 = 355 Lines Thanks Ulli for inspiration and lots of code.

