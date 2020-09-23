VERSION 5.00
Begin VB.Form FrmCWL 
   BorderStyle     =   1  'Fest Einfach
   Caption         =   "Create Wordlist"
   ClientHeight    =   5400
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5355
   Icon            =   "FrmCWL.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5400
   ScaleWidth      =   5355
   StartUpPosition =   3  'Windows-Standard
   Begin VB.CommandButton CmdQuit 
      Caption         =   "Quit"
      Height          =   495
      Left            =   1920
      TabIndex        =   7
      Top             =   4800
      Width           =   1215
   End
   Begin VB.CommandButton CmdConv 
      Caption         =   "Create new Wordlist"
      Enabled         =   0   'False
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   1200
      Width           =   5055
   End
   Begin VB.TextBox TxtSpezial 
      Height          =   285
      Left            =   120
      TabIndex        =   2
      Top             =   840
      Width           =   5055
   End
   Begin VB.CommandButton CmdOpen 
      Caption         =   "Open Wordlist"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   2055
   End
   Begin VB.TextBox Text1 
      Height          =   2175
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertikal
      TabIndex        =   0
      Top             =   2160
      Width           =   5055
   End
   Begin VB.Label LblStat 
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   120
      TabIndex        =   6
      Top             =   4440
      Width           =   5055
   End
   Begin VB.Label Label2 
      Caption         =   "Status / Results"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   1920
      Width           =   4575
   End
   Begin VB.Label Label1 
      Caption         =   "Special Chars (For German you should write ÄÖÜß)"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   600
      Width           =   5055
   End
End
Attribute VB_Name = "FrmCWL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Create Wordlists for WordPuz
'© 2006 ScytheVB
'This code is optimized for speed not for Size

Option Explicit
'For Hyperjump
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Dim FName As String             'Holds the name of the Original Wordlist

Private Type WL
 SortWord As String             'The Original Word
 InList As Boolean              'Should this one be saved ?
 Words() As String
End Type

'Get a Filename
Private Sub CmdOpen_Click()
 FName = GetOpenName("Open Wordlist", App.Path, "Textfiles|*.txt|All Files|*.*", , OFN_EXPLORER, Me.hWnd)
 CmdConv.Enabled = Not (FName = "")
 Text1 = vbNullString
End Sub

'Convert Original to WordPuz Wordlist
Private Sub CmdConv_Click()
 'Some Temporary Variables
 Dim I As Long
 Dim F As Long
 Dim X As Long
 Dim Y As Long
 Dim Z As Long
 Dim Tmp As String
 Dim TmpS As String
 Dim TmpC As String
 Dim TmpB As Boolean
 Dim B5 As Boolean
 Dim B4 As Boolean
 Dim B3 As Boolean
 Dim LoadedWords As Long
 Dim PossibleWords As Long

 Dim Wordlist() As WL            'Our Test Wordlist
 Dim WLCtr As Long

 Dim SpezialC() As Long          'Spezial chars needed for some languages
 Dim SpCtr As Long

 Dim SearchList() As String     'Holds the 6 Letter words
 
 Dim SaveList() As String       'The Scrambled result

 Dim Hitlist() As Long          'We temporaly store our right words here
 Dim HitCtr As Long
 Dim RealHits As Long

 'Set Variables and get some Values
 ReDim Wordlist(0)
 ReDim Wordlist(0).Words(0)
 ReDim SearchList(0)

 SpCtr = Len(TxtSpezial)
 ReDim SpezialC(SpCtr)
 For I = 1 To SpCtr
  SpezialC(I) = AscW(Mid$(TxtSpezial, I, 1))
 Next I

 'Open the file and get all Words between 3 and 6 letters in UpperCase
 'Check if its Only letters (incl. Spezial Chars)
 'Check for double Words
 'Sort letters for faster compare
 'Find double LetterWords (ITS, SIT, both are IST/ If one match the second also does)
 Text1 = "Open " & CutAfter(FName, "\") & vbCrLf
 Me.MousePointer = 11
 Open FName For Input As #1
  Do Until EOF(1)
   Line Input #1, Tmp
   LoadedWords = LoadedWords + 1
   Tmp = Trim$(UCase$(Tmp))
   Z = Len(Tmp)
   'Is the size OK ?
   
   If Z < 7 And Z > 2 Then
    'Check for Letters and sort
    For I = 1 To Z
     TmpB = False
     TmpS = Mid$(Tmp, I, 1)
     X = AscW(TmpS)
     'Normal Char
     If X > 64 Then
      If X < 91 Then
       TmpB = True
      End If
     End If
     'No Normal so check for Spezial
     If TmpB = False Then
      For F = 1 To SpCtr
       If X = SpezialC(F) Then TmpB = True
      Next F
     End If
     'Done all Char checks so go if this is none
     If TmpB = False Then Exit For
     'Sort the word (THIS should be HIST)
     If I = 1 Then
      Wordlist(WLCtr).SortWord = Wordlist(WLCtr).SortWord & TmpS
     Else
      Y = Len(Wordlist(WLCtr).SortWord)
      For F = 1 To Y
       TmpC = Mid$(Wordlist(WLCtr).SortWord, F, 1)
       If TmpC > TmpS Then
        Wordlist(WLCtr).SortWord = Left$(Wordlist(WLCtr).SortWord, F - 1) & TmpS & Right$(Wordlist(WLCtr).SortWord, Y - F + 1)
        Exit For
       End If
      Next F
      'Maybe it was the Biggest Char
      If F > Y Then
       Wordlist(WLCtr).SortWord = Wordlist(WLCtr).SortWord & TmpS
      End If
     End If
    Next I
    'Now check for double Words
    If TmpB = True Then
     For I = 0 To WLCtr - 1
      If Wordlist(I).SortWord = Wordlist(WLCtr).SortWord Then
       For F = 0 To UBound(Wordlist(I).Words)
        If Wordlist(I).Words(F) = Tmp Then
         TmpB = False
        End If
       Next F
       'Found a new LetterWord
       If TmpB Then
        ReDim Preserve Wordlist(I).Words(F)
        Wordlist(I).Words(F) = Tmp
        PossibleWords = PossibleWords + 1
        TmpB = False
       End If
       Exit For
      End If
     Next I
     'found a complete new word so store it
     If TmpB Then
      Wordlist(WLCtr).Words(0) = Tmp
      PossibleWords = PossibleWords + 1
      'Fond a new Searchword
      If Z = 6 Then
       Y = UBound(SearchList)
       SearchList(Y) = Wordlist(WLCtr).SortWord
       ReDim Preserve SearchList(Y + 1)
      End If
      WLCtr = WLCtr + 1
      ReDim Preserve Wordlist(WLCtr)
      ReDim Preserve Wordlist(WLCtr).Words(0)
     Else
      Wordlist(WLCtr).SortWord = vbNullString
     End If
    End If
   End If
   LblStat = "Loading and Sorting words " & LoadedWords
   DoEvents
  Loop
 Close

 'set the right lenght for our Wordlist
 WLCtr = WLCtr - 1
 ReDim Preserve Wordlist(WLCtr)
 ReDim Hitlist(WLCtr)


 Text1 = Text1 & LoadedWords & " words in list" & vbCrLf
 Text1 = Text1 & PossibleWords & " words found for check" & vbCrLf
 Text1 = Text1 & UBound(SearchList) & " words with 6 Letters" & vbCrLf

 PossibleWords = 0
 'Now Search the Words that matches our 6 letter words
 For I = 0 To UBound(SearchList) - 1
  LblStat = "Test word " & I & "  Found " & PossibleWords & " good"

  'no hits
  HitCtr = 0
  B3 = False
  B4 = False
  B5 = False
  RealHits = 0

  'Search
  For X = 0 To WLCtr
   Z = 1
   For F = 1 To Len(Wordlist(X).SortWord)
    Y = InStr(Z, SearchList(I), Mid$(Wordlist(X).SortWord, F, 1))
    If Y = 0 Then Exit For
    Z = Y + 1
   Next F

   'Found a word
   If Y <> 0 Then
    Hitlist(HitCtr) = X
    HitCtr = HitCtr + 1
    RealHits = RealHits + UBound(Wordlist(X).Words) + 1
    Select Case Len(Wordlist(X).SortWord)
    Case 5
     B5 = True
    Case 4
     B4 = True
    Case 3
     B3 = True
    End Select
   End If
   DoEvents
  Next X

  If RealHits > 17 And RealHits < 61 Then
   If B5 Then
    If B4 Then
     If B3 Then
      For F = 0 To HitCtr - 1
       If Wordlist(Hitlist(F)).InList = False Then
        PossibleWords = PossibleWords + UBound(Wordlist(Hitlist(F)).Words) + 1
        Wordlist(Hitlist(F)).InList = True
       End If
      Next F
     End If
    End If
   End If
  End If
 Next I
 Text1 = Text1 & PossibleWords & " words in new list" & vbCrLf


 'Now we have to scramble wordlist
 LblStat = "Scramble Wordlist"
 ReDim SearchList(PossibleWords)
 ReDim SaveList(PossibleWords)
 X = 0

 For I = 0 To WLCtr
  If Wordlist(I).InList Then
   For F = 0 To UBound(Wordlist(I).Words)
    SearchList(X) = Wordlist(I).Words(F)
    X = X + 1
   Next F
  End If
 Next I

 Randomize Time
 X = 0
 Do Until X = PossibleWords
  I = Rnd * PossibleWords
  If SearchList(I) <> vbNullChar Then
   SaveList(X) = SearchList(I)
   SearchList(I) = vbNullChar
   X = X + 1
  End If
 Loop
 LblStat = vbNullChar
 Text1 = Text1 & "All tests are done" & vbCrLf
 
 Me.MousePointer = 0
 'Now Save
 LoadedWords = 0
 FName = GetSaveName("Save new Wordlist", App.Path, FName, "Textfiles|*.txt|All Files|*.*", "*.txt", OFN_EXPLORER, Me.hWnd)
 If FName <> "" Then
  Open FName For Output As #1
   For I = 0 To PossibleWords
      Print #1, SaveList(I)
   Next I
  Close
  Text1 = Text1 & vbCrLf & "Saved as " & CutAfter(FName, "\")
 Else
  Text1 = Text1 & vbCrLf & "Error:" & vbCrLf & "No Filename given"
 End If
 If MsgBox("The Game is free" & vbCrLf & "Please Support it" & vbCrLf & "and send me the Wordlist", vbExclamation + vbOKCancel, "Created new Language Pack") = vbOK Then
   ShellExecute 0&, vbNullString, "mailto:scythe@cablenet.de?subject=New%20Wordlist%20for%20WordPuz", vbNullString, vbNullString, vbNormalFocus
   MsgBox "Thanks for supporting the Game"
 End If
 If LenB(FName) <> 0 Then
  MsgBox "Done" & vbCrLf & "Dont forget to create a Language File" & vbCrLf & "Named " & Left$(FName, Len(FName) - 4), vbInformation
 End If
End Sub

Public Function CutAfter(ByVal StringToCut As String, ByVal CutString As String) As String
 Dim Cutlenght As Long
 Cutlenght = Len(StringToCut) - Len(CutString) + 1
 Do Until Mid$(StringToCut, Cutlenght, Len(CutString)) = CutString
  Cutlenght = Cutlenght - 1
 Loop
 CutAfter = Right$(StringToCut, Len(StringToCut) - Cutlenght)
End Function

Private Sub CmdQuit_Click()
 Unload Me
 End
End Sub

'Take only uppercase Letters
Private Sub TxtSpezial_Change()
 TxtSpezial = UCase$(TxtSpezial)
 TxtSpezial.SelStart = Len(TxtSpezial)
End Sub



