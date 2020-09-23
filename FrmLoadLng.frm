VERSION 5.00
Begin VB.Form FrmLoadLng 
   BackColor       =   &H00000000&
   BorderStyle     =   4  'Festes Werkzeugfenster
   Caption         =   "WordPuz"
   ClientHeight    =   5190
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   3330
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   346
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   222
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows-Standard
   Begin VB.CommandButton CmdOK 
      Caption         =   "OK"
      Height          =   375
      Left            =   1080
      TabIndex        =   1
      Top             =   4560
      Width           =   975
   End
   Begin VB.OptionButton OptLng 
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   0
      Left            =   480
      TabIndex        =   0
      Top             =   360
      Width           =   3015
   End
End
Attribute VB_Name = "FrmLoadLng"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Language Selector for WordPuz
'© 2006 ScytheVB

Option Explicit

'Search for possible Languages
Private Sub Form_Load()
 Dim Tmp As String
 Dim i As Long

 Tmp = Dir("*.lng")
 If Tmp = "" Then
  MsgBox "No Languagefiles found" & vbCrLf & "Get some on www.scythe-tools.de" & vbCrLf & "Game will quit now", vbCritical, "Select language"
  End
 End If

 'Show all possible Languages
 Do Until Tmp = ""
  If i > 0 Then
   Load OptLng(i)
  End If
  'Load new Option
  OptLng(i).Top = i * 20 + 20
  OptLng(i).Left = 20
  OptLng(i).Caption = Left$(Tmp, Len(Tmp) - 4)
  OptLng(i).Visible = True
  i = i + 1
  Tmp = Dir
 Loop
 'Resize and pos the form
 Me.Height = (OptLng(i - 1).Top + 100) * Screen.TwipsPerPixelY
 CmdOK.Top = Me.ScaleHeight - 30
 Me.Left = FrmGame.Left + (FrmGame.Width - Me.Width) / 2
 Me.Top = FrmGame.Top + (FrmGame.Height - Me.Height) / 2
End Sub

'Set new language
Private Sub CmdOK_Click()
 Dim i As Long
 For i = 0 To OptLng.Count - 1
  If OptLng(i).Value Then Language = OptLng(i).Caption
 Next i
 Unload Me
End Sub


