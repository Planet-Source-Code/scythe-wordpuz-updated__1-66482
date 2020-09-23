VERSION 5.00
Begin VB.Form FrmHSC 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   BorderStyle     =   4  'Festes Werkzeugfenster
   Caption         =   "Highscore"
   ClientHeight    =   4215
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   4140
   ControlBox      =   0   'False
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
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   281
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   276
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows-Standard
   Begin VB.CommandButton CmdOK 
      Caption         =   "OK"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1560
      TabIndex        =   1
      Top             =   3720
      Width           =   1095
   End
   Begin VB.TextBox TxtHSC 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1920
      TabIndex        =   0
      Top             =   1200
      Visible         =   0   'False
      Width           =   2055
   End
End
Attribute VB_Name = "FrmHSC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Enter Highscore for WordPuz
'© 2006 ScytheVB

Option Explicit

Dim NewName As Long         'Holds the Pos in the list

'Hide form
Private Sub CmdOK_Click()
 Unload Me
End Sub

'Show list and InputTextbox
Private Sub Form_Load()
 Dim i As Long
 Dim X As Long
 Dim Y As Long
 Dim TmpB As Boolean

 Me.Caption = Lng(14)
 'Center Form
 Me.Left = FrmGame.Left + (FrmGame.Width - Me.Width) / 2
 Me.Top = FrmGame.Top + (FrmGame.Height - Me.Height) / 2
 'Hide Textbox
 TxtHSC.Visible = False

 'Go thru the highscores and pos our textbox if needed
 For i = 0 To 9
  Me.CurrentY = i * 22 + 10
  If HSC(9 - i).Name = "³" Then
   TxtHSC.Left = 10
   TxtHSC.Top = i * 22 + 10
   NewName = 9 - i
   TmpB = True
  Else
   Me.CurrentX = 10
   Me.Print HSC(9 - i).Name
  End If
  Me.CurrentY = i * 22 + 10
  Me.CurrentX = 200
  Me.Print HSC(9 - i).Score
 Next i
 'Show button or Textbox
 TxtHSC.Visible = TmpB
 CmdOK.Visible = Not TmpB
End Sub

'Only chars and Max 14 letters
Private Sub TxtHSC_Change()
 Dim Tmp As String
 Dim i As Long
 For i = 1 To Len(TxtHSC)
  If Mid$(TxtHSC, i, 1) Like "[a-zA-Z0-9 ]" Then Tmp = Tmp & Mid$(TxtHSC, i, 1)
 Next
 If Len(Tmp) > 14 Then Tmp = Left$(Tmp, 6)
 TxtHSC = Tmp
End Sub

'Enter and the Name is set
Private Sub TxtHSC_KeyUp(KeyCode As Integer, Shift As Integer)
 If KeyCode = 13 Then
  HSC(NewName).Name = TxtHSC
  Form_Load
 End If
End Sub
