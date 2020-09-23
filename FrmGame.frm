VERSION 5.00
Begin VB.Form FrmGame 
   BackColor       =   &H00000000&
   BorderStyle     =   3  'Fester Dialog
   Caption         =   "WordPuz © ScytheVB  2006 "
   ClientHeight    =   8505
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9405
   FillColor       =   &H00FFFFFF&
   ForeColor       =   &H00FFFFFF&
   Icon            =   "FrmGame.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   567
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   627
   StartUpPosition =   2  'Bildschirmmitte
   Begin VB.CommandButton CmdLng 
      BackColor       =   &H00000000&
      Height          =   495
      Left            =   7080
      Picture         =   "FrmGame.frx":12FA
      Style           =   1  'Grafisch
      TabIndex        =   35
      Top             =   7920
      Width           =   615
   End
   Begin VB.PictureBox PicSound2 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'Kein
      Height          =   480
      Left            =   8280
      Picture         =   "FrmGame.frx":170D
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   34
      Top             =   8520
      Width           =   480
   End
   Begin VB.PictureBox PicSound1 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'Kein
      Height          =   480
      Left            =   7800
      Picture         =   "FrmGame.frx":1C5F
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   33
      Top             =   8520
      Width           =   480
   End
   Begin VB.PictureBox PicMusic2 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'Kein
      Height          =   480
      Left            =   7320
      Picture         =   "FrmGame.frx":2112
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   32
      Top             =   8520
      Width           =   480
   End
   Begin VB.PictureBox PicMusic1 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'Kein
      Height          =   480
      Left            =   6840
      Picture         =   "FrmGame.frx":254F
      ScaleHeight     =   480
      ScaleWidth      =   420
      TabIndex        =   31
      Top             =   8520
      Width           =   420
   End
   Begin VB.CommandButton CmdSound 
      BackColor       =   &H00000000&
      Height          =   495
      Left            =   7800
      Style           =   1  'Grafisch
      TabIndex        =   30
      Top             =   7920
      Width           =   615
   End
   Begin VB.CommandButton cmdMusic 
      BackColor       =   &H00000000&
      Height          =   495
      Left            =   8520
      Style           =   1  'Grafisch
      TabIndex        =   29
      Top             =   7920
      Width           =   615
   End
   Begin VB.Frame FrmDone 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'Kein
      Height          =   5775
      Left            =   240
      TabIndex        =   20
      Top             =   720
      Visible         =   0   'False
      Width           =   8895
      Begin VB.CommandButton CmdNextLevel 
         Caption         =   "OK"
         Height          =   375
         Left            =   3840
         Style           =   1  'Grafisch
         TabIndex        =   22
         Top             =   5400
         Width           =   855
      End
      Begin VB.Label LblNL 
         Alignment       =   2  'Zentriert
         BackStyle       =   0  'Transparent
         Caption         =   "Weiter geht es mit Level"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   -120
         TabIndex        =   27
         Top             =   5040
         Width           =   8895
      End
      Begin VB.Label LblPos3 
         BackColor       =   &H00000000&
         ForeColor       =   &H00FFFFFF&
         Height          =   4095
         Left            =   6240
         TabIndex        =   26
         Top             =   840
         Width           =   2415
      End
      Begin VB.Label LblPos2 
         BackColor       =   &H00000000&
         ForeColor       =   &H00FFFFFF&
         Height          =   4095
         Left            =   3120
         TabIndex        =   25
         Top             =   840
         Width           =   2415
      End
      Begin VB.Label LblPos1 
         BackColor       =   &H00000000&
         ForeColor       =   &H00FFFFFF&
         Height          =   4095
         Left            =   120
         TabIndex        =   24
         Top             =   840
         Width           =   2415
      End
      Begin VB.Label LblPW 
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Mögliche Wörter waren:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   23
         Top             =   600
         Width           =   2175
      End
      Begin VB.Label LblGO 
         Alignment       =   2  'Zentriert
         BackStyle       =   0  'Transparent
         Caption         =   "Gut gemacht"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   0
         TabIndex        =   21
         Top             =   0
         Width           =   8775
      End
   End
   Begin VB.Timer TimerSound 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   8760
      Top             =   8520
   End
   Begin VB.FileListBox File1 
      Height          =   1260
      Left            =   7200
      Pattern         =   "*.mid"
      TabIndex        =   15
      Top             =   9120
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   8640
      Top             =   6480
   End
   Begin VB.CommandButton CmdNewGame 
      Caption         =   "Neues Spiel"
      Height          =   375
      Left            =   6840
      Style           =   1  'Grafisch
      TabIndex        =   11
      Top             =   6000
      Width           =   2175
   End
   Begin VB.CommandButton CmdReset 
      Caption         =   "Reset"
      Enabled         =   0   'False
      Height          =   375
      Left            =   3600
      Style           =   1  'Grafisch
      TabIndex        =   10
      Top             =   6000
      Width           =   2175
   End
   Begin VB.CommandButton CmdOK 
      Caption         =   "OK"
      DisabledPicture =   "FrmGame.frx":3011
      Enabled         =   0   'False
      Height          =   375
      Left            =   360
      Picture         =   "FrmGame.frx":5019
      Style           =   1  'Grafisch
      TabIndex        =   9
      Top             =   6000
      Width           =   2055
   End
   Begin VB.PictureBox PicFront 
      BorderStyle     =   0  'Kein
      Height          =   615
      Left            =   240
      Picture         =   "FrmGame.frx":7021
      ScaleHeight     =   41
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   587
      TabIndex        =   1
      Top             =   5160
      Width           =   8805
   End
   Begin VB.PictureBox PicBlock 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'Kein
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   72
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1395
      Index           =   5
      Left            =   7560
      ScaleHeight     =   93
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   92
      TabIndex        =   7
      Top             =   3900
      Width           =   1380
   End
   Begin VB.PictureBox PicBlock 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'Kein
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   72
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1395
      Index           =   4
      Left            =   6120
      ScaleHeight     =   93
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   92
      TabIndex        =   6
      Top             =   3900
      Width           =   1380
   End
   Begin VB.PictureBox PicBlock 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'Kein
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   72
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1395
      Index           =   3
      Left            =   4680
      ScaleHeight     =   93
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   92
      TabIndex        =   5
      Top             =   3900
      Width           =   1380
   End
   Begin VB.PictureBox PicBlock 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'Kein
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   72
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1395
      Index           =   2
      Left            =   3240
      ScaleHeight     =   93
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   92
      TabIndex        =   4
      Top             =   3900
      Width           =   1380
   End
   Begin VB.PictureBox PicBlock 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'Kein
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   72
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1395
      Index           =   1
      Left            =   1800
      ScaleHeight     =   93
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   92
      TabIndex        =   3
      Top             =   3900
      Width           =   1380
   End
   Begin VB.PictureBox PicBlock 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'Kein
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   72
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1395
      Index           =   0
      Left            =   360
      Picture         =   "FrmGame.frx":ACEF
      ScaleHeight     =   93
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   92
      TabIndex        =   2
      Top             =   3900
      Width           =   1380
   End
   Begin VB.PictureBox PicBack 
      BorderStyle     =   0  'Kein
      Height          =   1380
      Left            =   240
      Picture         =   "FrmGame.frx":C5F8
      ScaleHeight     =   92
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   587
      TabIndex        =   0
      Top             =   3840
      Width           =   8805
   End
   Begin VB.Frame FrmHigh 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'Kein
      Height          =   2655
      Left            =   3360
      TabIndex        =   36
      Top             =   1080
      Width           =   2655
      Begin VB.Label LblHighScore 
         Alignment       =   1  'Rechts
         BackColor       =   &H000000FF&
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00FFFFFF&
         Height          =   2415
         Left            =   1680
         TabIndex        =   39
         Top             =   240
         Width           =   975
      End
      Begin VB.Label LblHighName 
         BackColor       =   &H000000FF&
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00FFFFFF&
         Height          =   2415
         Left            =   0
         TabIndex        =   38
         Top             =   240
         Width           =   1575
      End
      Begin VB.Label LblHighTit 
         Alignment       =   2  'Zentriert
         BackColor       =   &H000000FF&
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   0
         TabIndex        =   37
         Top             =   0
         Width           =   2655
      End
   End
   Begin VB.Label LblLink 
      Alignment       =   2  'Zentriert
      BackStyle       =   0  'Transparent
      Caption         =   "www.scythe-tools.de"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   405
      Left            =   3240
      MouseIcon       =   "FrmGame.frx":172F3
      MousePointer    =   99  'Benutzerdefiniert
      TabIndex        =   28
      Top             =   8040
      Width           =   2895
   End
   Begin VB.Label LblFnd3 
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   975
      Left            =   6600
      TabIndex        =   19
      Top             =   6840
      Width           =   2415
   End
   Begin VB.Label LblFnd2 
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   975
      Left            =   3480
      TabIndex        =   18
      Top             =   6840
      Width           =   2415
   End
   Begin VB.Label LblFound 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Gefundene Wörter"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   360
      TabIndex        =   17
      Top             =   6600
      Width           =   2175
   End
   Begin VB.Label LblFnd1 
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   975
      Left            =   360
      TabIndex        =   16
      Top             =   6840
      Width           =   2415
   End
   Begin VB.Label LblPT 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   240
      TabIndex        =   14
      Top             =   480
      Width           =   1695
   End
   Begin VB.Label LblTime 
      Alignment       =   2  'Zentriert
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   420
      Left            =   120
      TabIndex        =   13
      Top             =   120
      Width           =   9210
   End
   Begin VB.Label LblDone 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   300
      Left            =   6120
      TabIndex        =   12
      Top             =   120
      Width           =   1410
   End
   Begin VB.Label LblMax 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   300
      Left            =   240
      TabIndex        =   8
      Top             =   120
      Width           =   90
   End
End
Attribute VB_Name = "FrmGame"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'WordPuz
'© 2006 ScytheVB



'Also included for this Project

'Create Wordlist
'A tool to convert any wordlist to a Wordpuz Wordlist
'See Directory "Wordlist Creation Tool"

'Create LNG
'Create your own Language Pack
'See Directory "Language Pack Creation Tool" or use a Texteditor :-)



Option Explicit

'For Hyperjump
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long



Dim Words() As String           'All words
Dim Mixed() As String           'All words with 6 Letters
Dim ResultWord As String        'The you typed shown on screen
Dim ResultWords() As String     'List of possible Words
Dim Possible As Long            'Number of possible Words
Dim Done As Long                'Number of found words for this Level
Dim Wordcounter As Long         'Startnumber for our 6 Words list
Dim StartTime As Variant        'Starttime of the level
Dim Level As Long               'Level
Dim OldWCTR As Long             'Check if we need a ned wordsearch
Dim MidiSnd As String           'Songnames
Dim Music As Boolean            'Songs On/Off
Dim Sound As Boolean            'Sound On/Off

'Change the Gamelanguage
Private Sub CmdLng_Click()
 CmdOK.Enabled = False
 Timer1.Enabled = False
 'Show Language selector
 FrmLoadLng.Show 1
 'Restart
 Form_Load
End Sub

'Continue game
Private Sub CmdNextLevel_Click()
 'Hide "well done" or "Game Over"
 FrmDone.Visible = False
 'Select a new Word
 Wordcounter = Wordcounter + 1
 'Set everything to Zero
 Done = 0
 LblPT = ""
 LblFnd1 = ""
 LblFound = ""
 LblFnd2 = ""
 'If Next level
 If LblGO = Lng(5) Then
  InitWord
  StartTime = DateAdd("n", 2, Time)
 Else 'Quit Game
  PrintText "SCYTHE"
 End If
End Sub

'Check if our word is in the list
Private Sub CmdOK_Click()
 Dim i As Long
 Dim X As Long
 Dim Tmp As String

 'Search all possible Words
 For i = 0 To UBound(ResultWords)
  'Found one
  If ResultWords(i) = ResultWord Then
   ResultWords(i) = ""
   GamePoints = GamePoints + Len(ResultWord)
   Done = Done + 1
   'Show the new word on the screen
   If Done > 5 Then
    If Done > 10 Then
     LblFnd3 = LblFnd3 & ResultWord & vbCrLf
    End If
    LblFnd2 = LblFnd2 & ResultWord & vbCrLf
   Else
    LblFnd1 = LblFnd1 & ResultWord & vbCrLf
   End If
   'Increase Points and found words on screen
   LblDone = Lng(1) & Done & " " & Lng(2) & Int(Level * 0.5) + 1
   LblPT = Lng(3) & GamePoints
   'Have we completed the level
   If Done = Int(Level * 0.5) + 1 Then
    Timer1.Enabled = False
    Level = Level + 1
    LblGO = Lng(5)
    LblNL = Lng(8) & Level
    If Sound Then PlayWaveRes 102, soundASYNC
    FrmDone.Visible = True
    Exit Sub
   End If
   'Set the letters back
   InitWord
   Exit Sub
  End If
 Next i
 'wrong word
 If Sound Then PlayWaveRes 100, soundASYNC
 If Timer1.Enabled Then
  MsgBox Lng(13)
  InitWord
 End If
End Sub

'Set the letters back
Private Sub CmdReset_Click()
 InitWord
End Sub

Private Sub CmdNewGame_Click()
 'Start Time
 Timer1.Enabled = True
 'Hide Highscores
 FrmHigh.Visible = False
 CmdReset.Enabled = True
 Wordcounter = Wordcounter + 1
 Done = 0
 Level = 1
 LblFnd1 = ""
 LblFound = ""
 LblFnd2 = ""
 InitWord
 GamePoints = 0
 StartTime = DateAdd("n", 2, Time)
End Sub
'En/Disable Sound & Music
Private Sub CmdSound_Click()
 If Sound Then
  CmdSound.Picture = PicSound2.Picture
 Else
  CmdSound.Picture = PicSound1.Picture
 End If
 Sound = Not Sound
End Sub
Private Sub cmdMusic_Click()
 If Music Then
  cmdMusic.Picture = PicMusic2.Picture
  StopMIDI MidiSnd
 Else
  PlayMIDI MidiSnd
  cmdMusic.Picture = PicMusic1.Picture
 End If
 Music = Not Music
 TimerSound.Enabled = Music
End Sub

'Start the Game
Private Sub Form_Load()
 Dim i As Long
 Dim f As Long
 Dim Tmp As String

 'Language ="" if the game is loaded
 If Language = "" Then
  Me.Caption = Me.Caption & "   Version " & App.Major & "." & App.Minor & "." & App.Revision
  If Dir("WordPuz.ini") <> "" Then
   Open "WordPuz.ini" For Input As #1
    Line Input #1, Language
    Line Input #1, Tmp
    If Tmp = "False" Then Sound = True
    Line Input #1, Tmp
    If Tmp = "False" Then Music = True
    cmdMusic_Click
    CmdSound_Click
   Close
  Else
   Sound = True
   Music = True
   CmdLng_Click
   cmdMusic_Click
   CmdSound_Click
  End If
 End If
 LoadLanguage

 'Load GFX
 For i = 1 To 5
  PicBlock(i).Picture = PicBlock(0).Picture
  CmdNewGame.Picture = CmdOK.Picture
  CmdReset.Picture = CmdOK.Picture
  CmdReset.DisabledPicture = CmdOK.DisabledPicture
  CmdNextLevel.Picture = CmdOK.Picture
 Next i
 'Set Language to our Buttons and labels
 LblFound = Lng(7)
 CmdNewGame.Caption = Lng(9)
 CmdNextLevel.Caption = Lng(10)
 CmdOK.Caption = Lng(11)
 CmdReset.Caption = Lng(12)
 LblPW = Lng(0)

 'Select a random 6 Letter Word
 Randomize Time
 Wordcounter = Rnd * UBound(Mixed)

 Level = 1
 ShowHighscore
 Me.Show
 PrintText "SCYTHE"
End Sub
'Show the Highscores on top of the game
Private Sub ShowHighscore()
 Dim i As Long
 FrmHigh.Visible = True
 LblHighTit = Lng(14)
 LblHighName = ""
 LblHighScore = ""
 For i = 9 To 0 Step -1
  LblHighName = LblHighName & HSC(i).Name & vbCrLf
  LblHighScore = LblHighScore & HSC(i).Score & vbCrLf
 Next i
End Sub

'Load Words and Captions
Private Sub LoadLanguage()
 Dim Tmp As String
 Dim i As Long
 Dim f As Long
 'Load Captions
 On Local Error GoTo NoLng
 Open Language & ".lng" For Input As #1
  Do Until EOF(1)
   Line Input #1, Tmp
   If Left$(Tmp, 1) <> "/" Then
    Lng(i) = Trim(Tmp) & " "
    i = i + 1
   End If
  Loop
 Close

 'Load Wordlist
 ReDim Words(0)
 ReDim Mixed(0)
 i = 0
 On Local Error GoTo NoWords
 Open Language & ".txt" For Input As #1
  Do Until EOF(1)
   ReDim Preserve Words(i)
   Line Input #1, Words(i)
   If Len(Words(i)) > 5 Then
    ReDim Preserve Mixed(f)
    Mixed(f) = Scramble(Words(i))
    f = f + 1
   End If
   i = i + 1
  Loop
 Close

 'Load Highscores
 'Every language has its own Highscore
 On Local Error GoTo NoScore
 Open Language & ".hsc" For Input As #1
  For i = 0 To 9
   Line Input #1, Tmp
   HSC(i).Name = Tmp
   Line Input #1, Tmp
   HSC(i).Score = Val(Tmp)
  Next i
 Close
 Exit Sub
NoScore:
 For i = 0 To 9
  HSC(i).Name = "Scythe"
  HSC(i).Score = i * 50 + 25
 Next i
 Exit Sub
NoLng:
 MsgBox "Cant find " & Language & ".lng", vbCritical
 FrmLoadLng.Show 1
 Resume
NoWords:
 i = MsgBox("Cant find " & Language & ".txt" & vbCrLf & "Select an other Language", vbCritical + vbOKCancel, "Missing Wordlist")
 If i = vbCancel Then End
 FrmLoadLng.Show 1
 LoadLanguage
End Sub

'Quit the game so save Data
Private Sub Form_Unload(Cancel As Integer)
 Dim i As Long
 StopMIDI MidiSnd
 'Save Settings
 Open "WordPuz.ini" For Output As #1
  Print #1, Language
  Print #1, Sound
  Print #1, Music
 Close
 'Save Highscore
 Open Language & ".hsc" For Output As #1
  For i = 0 To 9
   Print #1, HSC(i).Name
   Print #1, Str(HSC(i).Score)
  Next i
 Close
End Sub

'Click a letter
Private Sub PicBlock_Click(Index As Integer)
 If Timer1.Enabled = False Then Exit Sub
 'Ad a letter to our Resultword
 ResultWord = ResultWord & PicBlock(Index).Tag
 '3 and more letters are needed to press OK
 If Len(ResultWord) > 2 Then CmdOK.Enabled = True
 'Set the new position
 'Not the best way Bitblt... does better but this also works :-)
 PicBlock(Index).Top = 100
 PicBlock(Index).Left = (Len(ResultWord) * 96 + 24 - 96)
End Sub

'Search possible words
'show letters on board
'show list of possible words on "Game Over" screen
Private Sub InitWord()
 Dim i As Long
 Dim f As Long
 Dim a As String
 Dim j As Long
 Dim Y As Long
 Dim Ctr As Long
 Timer1.Enabled = True
 'Reached last 6 letter word so restart
 If Wordcounter > UBound(Mixed) Then Wordcounter = 0
 'Set typed word to ""
 ResultWord = ""
 CmdOK.Enabled = False

 'Search possible Words
 If Wordcounter <> OldWCTR Then
  LblPos1 = ""
  LblPos2 = ""
  LblPos3 = ""
  Possible = 0
  ReDim ResultWords(0)
  For i = 0 To UBound(Words)
   a = Words(i)
   j = Len(a)
   For f = 1 To 6
    Y = InStr(1, a, Mid$(Mixed(Wordcounter), f, 1))
    If Y <> 0 Then
     a = Left$(a, Y - 1) & " " & Right$(a, j - Y)
    End If
   Next f

   If Trim(a) = "" Then
    Possible = Possible + 1
    ReDim Preserve ResultWords(Ctr)
    ResultWords(Ctr) = Words(i)
    Ctr = Ctr + 1
   End If
  Next i

  'Sort for a better look
  BubbleSort ResultWords()
  'Show list of possible words on screen
  For i = 0 To Ctr - 1
   If i > 19 Then
    If i > 39 Then
     LblPos3 = LblPos3 & ResultWords(i) & vbCrLf
    Else
     LblPos2 = LblPos2 & ResultWords(i) & vbCrLf
    End If
   Else
    LblPos1 = LblPos1 & ResultWords(i) & vbCrLf
   End If
  Next i
  'Remember Position
  OldWCTR = Wordcounter
 End If

 'Refresh Data on screen
 LblDone = Lng(1) & Done & " " & Lng(2) & Int(Level * 0.5) + 1
 LblPT = Lng(3) & GamePoints
 LblMax = Lng(0) & Possible
 PrintText Mixed(Wordcounter)
End Sub

'Write a word on our board
Private Sub PrintText(ShownText As String)
 Dim i As Long
 Dim a As String
 For i = 0 To 5
With PicBlock(i)
 a = Mid$(ShownText, 1 + i, 1)
 .Cls
 .CurrentY = (.Height - .TextHeight(a)) / 2
 .CurrentX = (.Width - .TextWidth(a)) / 2
 .Top = 260
 .Left = 96 * i + 24
 .Tag = a
 PicBlock(i).Print a
End With
Next i

End Sub

'Game Over ?
Private Sub Timer1_Timer()
 Dim X As Long
 Dim Tmp As String
 Dim i As Long
 Dim f As Long
 'Get the actual Time we have for this level
 X = DateDiff("s", Time, StartTime)
 'Less than 30 sec so make it red on the screen
 If X < 30 Then
  LblTime.ForeColor = &HFF
 Else
  LblTime.ForeColor = &HFFFFFF
 End If
 If X < 0 Then X = 0

 'Show time
 LblTime = Lng(4) & Int(X / 60) & ":" & Format(Int(X - (Int(X / 60) * 60)), "00")

 'Game Over
 If X < 1 Then
  'Clean up
  Timer1.Enabled = False
  LblTime = ""
  LblGO = Lng(6)
  LblNL = ""
  LblDone = ""
  LblMax = ""
  'Play looser sound
  If Sound Then PlayWaveRes 101, soundASYNC
  'Made a new Highscore ?
  For i = 9 To 0 Step -1
   If HSC(i).Score < GamePoints Then
    For f = 0 To i - 1
     HSC(f) = HSC(f + 1)
    Next f
    HSC(i).Score = GamePoints
    HSC(i).Name = "³"
    FrmHSC.Show 1
    Exit For
   End If
  Next i
  'Show Game over on screen
  FrmDone.Visible = True
  'Show Highscorelist on screen (behind Game over)
  ShowHighscore
 End If
End Sub

'Midiroutine for random midisound
Private Sub TimerSound_Timer()
 'no midis ind directory
 If File1.ListCount = 0 Then
  TimerSound.Enabled = False
  Exit Sub
 End If

 'Is the song still playing
 If StillMidi(MidiSnd) = False Then
  'Remove the song
  StopMIDI (MidiSnd)
  'Select a random song
  Randomize Time
  MidiSnd = GetShortName(File1.Path & "\" & File1.List(Rnd * File1.ListCount))
  'Play ne song
  PlayMIDI MidiSnd
 End If
End Sub
'Jump to www.scythe-tools.de for a newer Version
Private Sub LblLink_Click()
 ShellExecute 0&, vbNullString, LblLink.Caption, vbNullString, vbNullString, vbNormalFocus
End Sub
