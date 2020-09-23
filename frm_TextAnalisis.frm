VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "ComDlg32.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "richtx32.ocx"
Begin VB.Form frm_TextAnalisis 
   Caption         =   "Count CharacterDemo"
   ClientHeight    =   5160
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   10380
   LinkTopic       =   "Form1"
   ScaleHeight     =   5160
   ScaleWidth      =   10380
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtAcceptablelPunctuation 
      Height          =   375
      Left            =   7800
      TabIndex        =   19
      Top             =   4680
      Width           =   2415
   End
   Begin VB.CommandButton cmdStatistics 
      Caption         =   "Layout Characters(No Spaces)"
      Height          =   375
      Index           =   7
      Left            =   7440
      TabIndex        =   18
      Top             =   3240
      Width           =   2895
   End
   Begin VB.CommandButton cmdStatistics 
      Caption         =   "Layout  Characters Used"
      Height          =   375
      Index           =   6
      Left            =   7440
      TabIndex        =   17
      Top             =   2760
      Width           =   2895
   End
   Begin VB.CheckBox chkCaseSensitive 
      Caption         =   "Case Sensitive"
      Height          =   255
      Left            =   4320
      TabIndex        =   16
      Top             =   4440
      Width           =   1455
   End
   Begin VB.CommandButton WordUsage 
      Caption         =   "Word"
      Height          =   255
      Left            =   5880
      TabIndex        =   14
      Top             =   4800
      Width           =   975
   End
   Begin VB.CommandButton cmdStatistics 
      Caption         =   "Punctuation Used (No Layout)"
      Height          =   375
      Index           =   5
      Left            =   7440
      TabIndex        =   13
      Top             =   2235
      Width           =   2895
   End
   Begin VB.CommandButton cmdStatistics 
      Caption         =   "Punctuation Used"
      Height          =   375
      Index           =   4
      Left            =   7440
      TabIndex        =   12
      Top             =   1860
      Width           =   2895
   End
   Begin VB.CommandButton cmdStatistics 
      Caption         =   "Numerals Used"
      Height          =   375
      Index           =   3
      Left            =   7440
      TabIndex        =   11
      Top             =   1485
      Width           =   2895
   End
   Begin VB.CommandButton cmdStatistics 
      Caption         =   "Unique Words (Any Case)"
      Height          =   375
      Index           =   2
      Left            =   7440
      TabIndex        =   10
      Top             =   1110
      Width           =   2895
   End
   Begin VB.CommandButton cmdStatistics 
      Caption         =   "Unique Word (Case Aware)"
      Height          =   375
      Index           =   1
      Left            =   7440
      TabIndex        =   9
      Top             =   735
      Width           =   2895
   End
   Begin VB.CommandButton cmdStatistics 
      Caption         =   "Word Count"
      Height          =   375
      Index           =   0
      Left            =   7440
      TabIndex        =   8
      Top             =   360
      Width           =   2895
   End
   Begin RichTextLib.RichTextBox rtbDemo 
      Height          =   3975
      Left            =   120
      TabIndex        =   7
      Top             =   120
      Width           =   3495
      _ExtentX        =   6165
      _ExtentY        =   7011
      _Version        =   393217
      TextRTF         =   $"frm_TextAnalisis.frx":0000
   End
   Begin MSComDlg.CommonDialog cdlDemo 
      Left            =   3480
      Top             =   720
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.OptionButton optSort 
      Caption         =   "Count Sort"
      Height          =   195
      Index           =   1
      Left            =   5400
      TabIndex        =   6
      Top             =   4200
      Width           =   1095
   End
   Begin VB.OptionButton optSort 
      Caption         =   "Alpha Sort"
      Height          =   195
      Index           =   0
      Left            =   4320
      TabIndex        =   5
      Top             =   4200
      Value           =   -1  'True
      Width           =   1095
   End
   Begin VB.CommandButton cmdFind 
      Caption         =   "Find (Any Case)"
      Height          =   375
      Index           =   1
      Left            =   120
      TabIndex        =   4
      Top             =   4560
      Width           =   1575
   End
   Begin VB.CommandButton cmdUsage 
      Caption         =   "Character"
      Height          =   255
      Left            =   4440
      TabIndex        =   3
      Top             =   4800
      Width           =   975
   End
   Begin VB.ListBox lstUsage 
      Columns         =   3
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   9
         Charset         =   255
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3840
      Left            =   3840
      Sorted          =   -1  'True
      TabIndex        =   2
      Top             =   120
      Width           =   3495
   End
   Begin VB.TextBox txtFind 
      Height          =   375
      Left            =   1800
      TabIndex        =   1
      Text            =   "txtFind"
      Top             =   4320
      Width           =   1575
   End
   Begin VB.CommandButton cmdFind 
      Caption         =   "Find (Case Aware)"
      Height          =   375
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   4200
      Width           =   1575
   End
   Begin VB.Label lblAcceptableInternal 
      Caption         =   $"frm_TextAnalisis.frx":008B
      Height          =   855
      Left            =   7560
      TabIndex        =   20
      Top             =   3840
      Width           =   2775
   End
   Begin VB.Label lblListCount 
      Caption         =   "List count = 0"
      Height          =   255
      Left            =   3960
      TabIndex        =   15
      Top             =   3960
      Width           =   3375
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileOpt 
         Caption         =   "&Open"
         Index           =   0
      End
      Begin VB.Menu mnuFileOpt 
         Caption         =   "-"
         Index           =   1
      End
      Begin VB.Menu mnuFileOpt 
         Caption         =   "E&xit"
         Index           =   2
      End
   End
End
Attribute VB_Name = "frm_TextAnalisis"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private TextAnalisis     As New cls_TextAnalysis

Private Sub cmdFind_Click(Index As Integer)

  DocmdFindClick Index

End Sub

Private Sub cmdStatistics_Click(Index As Integer)

  Select Case Index
   Case 0
    cmdStatistics(0).Caption = "Word Count = " & TextAnalisis.CountWords
   Case 1
    cmdStatistics(1).Caption = "Unique Words (Case Aware) = " & TextAnalisis.CountUniqueWords
   Case 2
    cmdStatistics(2).Caption = "Unique Words (Any Case) = " & TextAnalisis.CountUniqueWords(False)
   Case 3
    cmdStatistics(3).Caption = "Numerals Used = " & TextAnalisis.CountNumerals
   Case 4
    cmdStatistics(4).Caption = "Punctuation Used = " & TextAnalisis.CountPunctuation()
   Case 5
    cmdStatistics(5).Caption = "Punctuation Used (Layout) = " & TextAnalisis.CountPunctuation(False)
   Case 6
    cmdStatistics(6).Caption = "Layout Characters Used = " & TextAnalisis.CountFormat
   Case 7
    cmdStatistics(7).Caption = "Layout Characters(No Spaces) = " & TextAnalisis.CountFormat(True)
  End Select

End Sub

Private Sub cmdUsage_Click()

  Dim I                                 As Long
  Dim CaseAware                         As Boolean

  lstUsage.Clear
  CaseAware = chkCaseSensitive.Value = vbChecked
  For I = 1 To 255
    If TextAnalisis.HasChar(Chr$(I), CaseAware) Then
'this trap is to stop letters double hitting on caseless tests
      If Not CaseAware Then
        If TextAnalisis.IsAlpha(Chr$(I)) Or IsNumeric(Chr$(I)) Then
          If TextAnalisis.IsUpper(Chr$(I)) And Not IsNumeric(Chr$(I)) Then
            GoTo skipUcase
          End If
        End If
      End If
      Select Case optSort(0).Value
       Case True
        lstUsage.AddItem " " & TextAnalisis.CharDesc(I) & " = " & TextAnalisis.CountChar(Chr$(I), CaseAware)
       Case False
        lstUsage.AddItem " " & Format$(TextAnalisis.CountChar(Chr$(I), CaseAware), "000") & " = " & TextAnalisis.CharDesc(I)
      End Select
skipUcase:
    End If
  Next I
  lblListCount.Caption = "List count = " & lstUsage.ListCount

End Sub

Private Sub DocmdFindClick(ByVal lngIndex As Long)

  Select Case lngIndex
   Case 0
    cmdFind(0).Caption = "Case Find =" & TextAnalisis.CountChar(txtFind.Text)
   Case 1
    cmdFind(1).Caption = "Any Case Find =" & TextAnalisis.CountChar(txtFind.Text, False)
  End Select

End Sub

Private Sub Form_Load()

  rtbDemo.Text = "The quick brown fox jumed over the lazy dog." & vbNewLine & "You can also use the menu to load txt and rtf files for more experiments." & vbNewLine & _
   "Note 1 Characters below ASCII 33 (Space or less) are shown with their Code name." & vbNewLine & "Note 2 ListBox sort doesn't cope properly with the low ASCII characters (Space (32) lists before the lower characters)"
  txtFind.Text = "the"
  TextAnalisis.SourceString = rtbDemo.Text

End Sub

Private Sub Form_Unload(Cancel As Integer)

  End

End Sub

Private Sub mnuFileOpt_Click(Index As Integer)

  Select Case Index
   Case 0
    With cdlDemo
      .Filter = "Doc files|*.txt;*.rtf"
      .FilterIndex = 1
      .ShowOpen
      If Len(.FileName) Then
        rtbDemo.LoadFile (.FileName)
      End If
    End With
'Case 1
   Case 2
    End
  End Select

End Sub

Private Sub rtbDemo_Change()

'this keeps the class upto date with the current text

  TextAnalisis.SourceString = rtbDemo.Text

End Sub

Private Sub txtAcceptablelPunctuation_Change()

  TextAnalisis.AcceptableInternalPunctuation = txtAcceptablelPunctuation.Text

End Sub

Private Sub txtFind_KeyPress(KeyAscii As Integer)

  If KeyAscii = 13 Then
    KeyAscii = 0
    DocmdFindClick 0
    DocmdFindClick 1
  End If

End Sub

Private Sub WordUsage_Click()

  Dim ArrTmp()  As String
  Dim I         As Long

  lstUsage.Clear
  TextAnalisis.CaseSensitive = chkCaseSensitive.Value = vbChecked
  ArrTmp = TextAnalisis.UniqueWordArray(chkCaseSensitive.Value = vbChecked)
  For I = LBound(ArrTmp) To UBound(ArrTmp)
    Select Case optSort(0).Value
     Case True
      lstUsage.AddItem " " & ArrTmp(I) & " = " & TextAnalisis.CountWholeWord(ArrTmp(I))
     Case False
'NOTE format is used to force listbox Sort to order correctly
      lstUsage.AddItem " " & Format$(TextAnalisis.CountWholeWord(ArrTmp(I)), "000") & " = " & ArrTmp(I)
    End Select
  Next I
  lblListCount.Caption = "List count = " & lstUsage.ListCount

End Sub

':)Roja's VB Code Fixer V1.1.93 (23/02/2004 4:16:31 PM) 2 + 151 = 153 Lines Thanks Ulli for inspiration and lots of code.

