VERSION 5.00
Begin VB.Form FrmCreateLng 
   BorderStyle     =   1  'Fest Einfach
   Caption         =   "Create Language Pack"
   ClientHeight    =   1080
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7530
   Icon            =   "FrmCreateLng.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   72
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   502
   StartUpPosition =   2  'Bildschirmmitte
   Begin VB.CommandButton CmdSave 
      Caption         =   "Save"
      Height          =   375
      Left            =   3000
      TabIndex        =   2
      Top             =   600
      Width           =   1575
   End
   Begin VB.TextBox TxtLng 
      Height          =   285
      Index           =   0
      Left            =   120
      TabIndex        =   1
      Top             =   240
      Width           =   7335
   End
   Begin VB.Label LblDesc 
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   7215
   End
End
Attribute VB_Name = "FrmCreateLng"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Create Language Pack for WordPuz
'© 2006 ScytheVB

Option Explicit

'For Hyperjump
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Private Sub CmdSave_Click()
 Dim FName As String
 Dim i As Long
 FName = GetSaveName("Save in Game Directory !!!", App.Path, FName, "Lng Files|*.lng", "*.lng", OFN_EXPLORER, Me.hWnd)
 If FName <> "" Then
  Open FName For Output As #1
   For i = 0 To TxtLng.Count - 1

    Print #1, TxtLng(i)
   Next i
  Close
  If MsgBox("The Game is free" & vbCrLf & "Please Support it" & vbCrLf & "and send me the Languagepack", vbExclamation + vbOKCancel, "Created new Language Pack") = vbOK Then
   ShellExecute 0&, vbNullString, "mailto:scythe@cablenet.de?subject=New%20Language%20Pack%20for%20WordPuz", vbNullString, vbNullString, vbNormalFocus
   MsgBox "Thanks for supporting the Game"
  End If
  If LenB(FName) <> 0 Then
   MsgBox "Done" & vbCrLf & "Dont forget to create a Wordlist File" & vbCrLf & "Named " & Left$(FName, Len(FName) - 4), vbInformation
  End If
 End If
End Sub

Private Sub Form_Load()
 Dim Tmp As String
 Dim i As Long

 Tmp = Dir("English.lng")
 If Tmp = "" Then
  MsgBox "English.lng is Missing" & vbCrLf & "Get it from www.scythe-tools.de" & vbCrLf & "Tool will quit now", vbCritical, "Create Language Pack"
  End
 End If

 Open "english.lng" For Input As #1

  Do Until EOF(1)
   If i > 0 Then
    Load LblDesc(i)
    LblDesc(i).Top = 50 * i
    LblDesc(i).Visible = True
    Load TxtLng(i)
    TxtLng(i).Top = 50 * i + 16
    TxtLng(i).Visible = True
   End If
   Line Input #1, Tmp
   Tmp = Trim(Right$(Tmp, Len(Tmp) - 1))
   LblDesc(i) = Tmp
   Line Input #1, Tmp
   TxtLng(i) = Tmp
   i = i + 1
  Loop
  Close
  Me.Height = (i + 1) * 50 * Screen.TwipsPerPixelX
  CmdSave.Top = Me.ScaleHeight - 28
End Sub

