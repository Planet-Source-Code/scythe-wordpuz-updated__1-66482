Attribute VB_Name = "CommanDialog"
' ---------------------------------------------
' Standardmodul FileOpenSave.bas
' Copyright © 2003 by Mathias Schiffer
' ---------------------------------------------
' Ruft die Standarddialoge "Datei öffnen" und
' "Datei speichern unter" auf und ermöglicht
' das Umschalten der verwendeten Dateiansicht.
' ---------------------------------------------
Option Explicit
' Diverse API-Konstante
Private Const MAX_PATH As Long = 260&
'Private Const WM_NOTIFY As Long = &H4E&
Private Const WM_COMMAND As Long = &H111&
Private Const OFS_MAXPATHNAME As Long = 128&
' WM_COMMAND-Nachrichten für den Common Dialog
Private Const CDN_FIRST As Long = &HFFFFFDA7
Private Const CDN_INITDONE As Long = (CDN_FIRST - &H0&)
Private Const CDN_SELCHANGE As Long = (CDN_FIRST - &H1&)
Private Const CDN_FOLDERCHANGE As Long = (CDN_FIRST - &H2&)
Private Const CDN_SHAREVIOLATION As Long = (CDN_FIRST - &H3&)
Private Const CDN_HELP As Long = (CDN_FIRST - &H4&)
Private Const CDN_FILEOK As Long = (CDN_FIRST - &H5&)
Private Const CDN_TYPECHANGE As Long = (CDN_FIRST - &H6&)
Private Const CDN_INCLUDEITEM As Long = (CDN_FIRST - &H7&)
Private Const WM_DESTROY = &H2
' Aufzählung für die Dateiansicht
Public Enum VIEWENUM
LargeIcon = &H7029&
List = &H702B&
Report = &H702C&
SmallIcon = &H702A&
Thumbnails = &H702D&
End Enum
' Konstante für Flags-Parameter der OPENFILENAME-Struktur
Public Enum OFN_FLAGS
OFN_SHAREWARN = 0&
OFN_SHARENOWARN = 1&
OFN_READONLY = &H1&
OFN_SHAREFALLTHROUGH = 2&
OFN_OVERWRITEPROMPT = &H2&
OFN_HIDEREADONLY = &H4&
OFN_NOCHANGEDIR = &H8&
OFN_SHOWHELP = &H10&
OFN_ENABLEHOOK = &H20&
OFN_ENABLETEMPLATE = &H40&
OFN_ENABLETEMPLATEHANDLE = &H80&
OFN_NOVALIDATE = &H100&
OFN_ALLOWMULTISELECT = &H200&
OFN_EXTENSIONDIFFERENT = &H400&
OFN_PATHMUSTEXIST = &H800&
OFN_FILEMUSTEXIST = &H1000&
OFN_CREATEPROMPT = &H2000&
OFN_SHAREAWARE = &H4000&
OFN_NOREADONLYRETURN = &H8000&
OFN_NOTESTFILECREATE = &H10000
OFN_NONETWORKBUTTON = &H20000
OFN_NOLONGNAMES = &H40000
OFN_EXPLORER = &H80000
OFN_NODEREFERENCELINKS = &H100000
OFN_LONGNAMES = &H200000
OFN_ENABLEINCLUDENOTIFY = &H400000
OFN_ENABLESIZING = &H800000
OFN_USESHELLITEM = &H1000000
OFN_DONTADDTORECENT = &H2000000
OFN_FORCESHOWHIDDEN = &H10000000
End Enum
' Details zu einer "Notification Message" (WM_NOTIFY)
Private Type NMHDR
 hwndFrom As Long
 idfrom As Long
 code As Long
End Type
' OPENFILENAME-Struktur für GetOpenFilename / GetSaveFileName
Private Type OPENFILENAME
 lStructSize As Long
 hWndOwner As Long
 hInstance As Long
 lpstrFilter As String
 lpstrCustomFilter As String
 nMaxCustFilter As Long
 nFilterIndex As Long
 lpstrFile As String
 nMaxFile As Long
 lpstrFileTitle As String
 nMaxFileTitle As Long
 lpstrInitialDir As String
 lpstrTitle As String
 Flags As OFN_FLAGS
 nFileOffset As Integer
 nFileExtension As Integer
 lpstrDefExt As String
 lCustData As Long
 lpfnHook As Long
 lpTemplateName As String
' Ab Windows 2000:
 pvReserved As Long
 dwReserved As Long
 FlagsEx As Long
End Type
' API-Funktionsprototypen
Private Declare Function GetOpenFileName Lib "comdlg32" Alias "GetOpenFileNameA" (ByRef pOpenfilename As OPENFILENAME) As Long
Private Declare Function GetSaveFileName Lib "comdlg32" Alias "GetSaveFileNameA" (ByRef lpOpenfilename As OPENFILENAME) As Long
Private Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWndParent As Long, ByVal hWndChildAfter As Long, ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare Function GetParent Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByRef wParam As Any, ByRef lParam As Any) As Long
Private Declare Function GetFileVersionInfoSize Lib "Version.dll" Alias "GetFileVersionInfoSizeA" (ByVal lptstrFilename As String, ByRef lpdwHandle As Long) As Long
Private Declare Function GetFileVersionInfo Lib "Version.dll" Alias "GetFileVersionInfoA" (ByVal lptstrFilename As String, ByVal dwHandle As Long, ByVal dwLen As Long, ByRef lpData As Any) As Long
Private Declare Function VerQueryValue Lib "Version.dll" Alias "VerQueryValueA" (ByRef pBlock As Any, ByVal lpSubBlock As String, ByRef lplpBuffer As Any, ByRef puLen As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByRef lpDest As Any, ByRef lpSource As Any, ByVal cBytes As Long)
'Get Dialogs Position and Size

Private Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long
Private Type RECT
 Left As Long
 Top As Long
 Right As Long
 Bottom As Long
End Type
Private Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long


'Move/Resize the Dialog and our own form
Private Declare Function MoveWindow Lib "user32" (ByVal hWnd As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long
Private Declare Sub SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long)
Private Const HWND_TOP = 0


Private CdlgHwnd As Long 'This will hold the Hwnd for Common Dialog
Private Const WM_NOTIFY = &H4E
'Private Const WM_DESTROY = &H2
Private Const WM_INITDIALOG = &H110

Public CMPrev As Long
'Common Dialog Messages
Private Const CDM_GETFILEPATH = &H465
Private Const CDM_GETFOLDERPATH = &H466

' Modulweit gültige Variable
Private ShowStateValue As VIEWENUM
Public Function GetSaveName(Optional ByVal DialogTitle As String = vbNullString, Optional ByVal InitialDir As String = vbNullString, Optional ByVal DefaultFilename As String = vbNullString, Optional ByVal Filter As String = vbNullString, Optional ByVal DefaultExtension As String = vbNullString, Optional ByVal Flags As OFN_FLAGS = 0&, Optional ByVal hWndOwner As Long = 0&, Optional ByVal ShowState As VIEWENUM = List) As String

 On Error GoTo GetSaveName_Error
 ' Zeigt den Dialog "Datei speichern" an und liefert
 ' den gewählten Dateipfad zurück (bzw. einen leeren
 ' String bei Abbruch durch den Benutzer).
 Dim OFN As OPENFILENAME
 Dim lNullPos As Long
With OFN
 ' Angabe der Größe der OPENFILENAME-Struktur
 .lStructSize = Len(OFN)
 ' Falls OS < Windows 2000, die letzten drei
' Mitglieder der Struktur ausnehmen:
 If Not Win2000Shell() Then
  .lStructSize = .lStructSize - 12
 End If
 ' OPENFILENAME-Parameter angeben
 .hWndOwner = hWndOwner
 .lpstrFile = DefaultFilename & String$(4096, 0)
 .nMaxFile = Len(DefaultFilename) + 4096
 .lpstrFileTitle = String$(MAX_PATH, 0)
 .nMaxFileTitle = MAX_PATH
 ' Filter-Trennzeichen "|" in vbNullChar wandeln
 .lpstrFilter = Replace(Filter, "|", vbNullChar)
 ' Vorgabe einer Standard-Erweiterung
 If Len(DefaultExtension) Then
  .lpstrDefExt = DefaultExtension
 End If
 ' Startverzeichnis des Dialogs
 If Len(InitialDir) Then
  .lpstrInitialDir = InitialDir
 End If
 ' Titel des Dialogs
 If Len(DialogTitle) Then
  .lpstrTitle = DialogTitle
 End If
 ' Eigenschaften des Dialogs
 If Flags = 0 Then
  .Flags = OFN_EXTENSIONDIFFERENT Or OFN_NOCHANGEDIR Or OFN_OVERWRITEPROMPT Or OFN_HIDEREADONLY Or OFN_EXPLORER Or OFN_ENABLESIZING Or OFN_FORCESHOWHIDDEN
 Else
  .Flags = Flags
 End If
' Bei Bedarf den Dialog-Hook aktivieren:
 If ShowState <> List Then
  ShowStateValue = ShowState
  .Flags = .Flags Or OFN_ENABLEHOOK
  .lpfnHook = ProcAddress(AddressOf OFNHookProc)
 End If
End With
If GetSaveFileName(OFN) Then
 lNullPos = InStr(OFN.lpstrFile, vbNullChar)
 If lNullPos > 1 Then
  GetSaveName = Left$(OFN.lpstrFile, lNullPos - 1)
 End If
End If

Exit Function

GetSaveName_Error:

MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure" & Chr(13) & Chr(10) & "(GetSaveName) of (Cmdlg2.bas).", vbCritical

End Function
Public Function GetOpenName(Optional ByVal DialogTitle As String = vbNullString, Optional ByVal InitialDir As String = vbNullString, Optional ByVal Filter As String = vbNullString, Optional ByVal DefaultExtension As String = vbNullString, Optional ByVal Flags As OFN_FLAGS = 0&, Optional ByVal hWndOwner As Long = 0&, Optional ByVal ShowState As VIEWENUM = List) As String

 On Error GoTo GetOpenName_Error
 ' Zeigt den Dialog "Datei öffnen" an und liefert
 ' den angegebenen Dateipfad zurück (bzw. einen leeren
 ' String bei Abbruch durch den Benutzer).
 Dim OFN As OPENFILENAME
 Dim lNullPos As Long
With OFN
 ' Angabe der Größe der OPENFILENAME-Struktur
 .lStructSize = Len(OFN)
 ' Falls OS < Windows 2000, die letzten drei
' Mitglieder der Struktur ausnehmen:
 If Not Win2000Shell() Then
  .lStructSize = .lStructSize - 12
 End If
 ' OPENFILENAME-Parameter angeben
 .hWndOwner = hWndOwner
 .lpstrFile = String$(4096, 0)
 .nMaxFile = 4096
 .lpstrFileTitle = String$(MAX_PATH, 0)
 .nMaxFileTitle = MAX_PATH
 ' Filter-Trennzeichen "|" in vbNullChar wandeln
 .lpstrFilter = Replace(Filter, "|", vbNullChar)
 ' Vorgabe einer Standard-Erweiterung
 If Len(DefaultExtension) Then
  .lpstrDefExt = DefaultExtension
 End If
 ' Startverzeichnis des Dialogs
 If Len(InitialDir) Then
  .lpstrInitialDir = InitialDir
 End If
 ' Titel des Dialogs
 If Len(DialogTitle) Then
  .lpstrTitle = DialogTitle
 End If
 ' Eigenschaften des Dialogs
 If Flags = 0 Then
  .Flags = OFN_EXTENSIONDIFFERENT Or OFN_NOCHANGEDIR Or OFN_HIDEREADONLY Or OFN_EXPLORER Or OFN_ENABLESIZING Or OFN_FORCESHOWHIDDEN
 Else
  .Flags = Flags
  .lpfnHook = ProcAddress(AddressOf CmdlgHook)
 End If
' Bei Bedarf den Dialog-Hook aktivieren:
 If ShowState <> List Then
  ShowStateValue = ShowState
  .Flags = .Flags Or OFN_ENABLEHOOK
  .lpfnHook = ProcAddress(AddressOf CmdlgHook)
 End If
End With
' Aufruf und Auswertung von GetOpenFileName
If GetOpenFileName(OFN) Then
 lNullPos = InStr(OFN.lpstrFile, vbNullChar)
 If lNullPos > 1 Then
  GetOpenName = Left$(OFN.lpstrFile, lNullPos - 1)
 End If
End If

Exit Function

GetOpenName_Error:

MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure" & Chr(13) & Chr(10) & "(GetOpenName) of (Cmdlg2.bas).", vbCritical

End Function
Private Function OFNHookProc(ByVal hDialog As Long, ByVal Message As Long, ByVal wParam As Long, ByRef lParam As NMHDR) As Long

 On Error GoTo OFNHookProc_Error
 ' Diese Hook-Funktion für den CommonDialog
 ' ermöglicht die Interaktion mit dem modalen Dialog.
 Dim hWndLVParent As Long ' Handle des ListView-Elternfensters
 Dim lpNMHDR As NMHDR  ' WM_NOTIFY-Nachrichtendetails
 If Message = WM_NOTIFY Then ' Nachricht vom Dialog
  Select Case lParam.code
  Case CDN_INITDONE
   ' Die Initialisierung des Dialogs wurde abgeschlossen.
   ' Alle Fenster des Dialogs sind an ihren Positionen.
   ' Das ListView-Elternfenster ermitteln
   hWndLVParent = FindWindowEx(GetParent(hDialog), 0, "SHELLDLL_DefView", vbNullString)
   If hWndLVParent <> 0 Then
    ' Mittels WM_COMMAND die Dateiansicht des ListView-
    ' Fensters ändern
    Call SendMessage(hWndLVParent, WM_COMMAND, ByVal ShowStateValue, ByVal 0&)
   End If
  Case CDN_FOLDERCHANGE
   ' Der Anwender hat einen Ordner ausgewählt.
  Case CDN_SELCHANGE
   ' Der Anwender hat seine Auswahl geändert.
  Case CDN_SHAREVIOLATION
   ' Eine Zugriffsverletzung ist aufgetreten.
  Case CDN_HELP
   ' Der Anwender hat den Hilfe-Button betätigt.
  Case CDN_FILEOK
   ' Der Anwender hat den OK-Button betätigt.
  Case CDN_TYPECHANGE
   ' Die Dateityp-Vorauswahl wurde geändert.
  End Select
 End If

 Exit Function

OFNHookProc_Error:

 MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure" & Chr(13) & Chr(10) & "(OFNHookProc) of (Cmdlg2.bas).", vbCritical

End Function
Private Function ProcAddress(ByVal AddressOfProc As Long) As Long

 On Error GoTo ProcAddress_Error
 ' Hilfsfunktion, um die Adresse einer Prozedur
 ' mithilfe des AddressOf-Operators ermitteln zu köennen.
 ProcAddress = AddressOfProc

 Exit Function

ProcAddress_Error:

 MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure" & Chr(13) & Chr(10) & "(ProcAddress) of (Cmdlg2.bas).", vbCritical

End Function
Private Function Win2000Shell() As Boolean

 On Error GoTo Win2000Shell_Error
 ' Ermittelt, ob die Hauptversionsnummer der Bibliothek
 ' Comdlg32.dll größer oder gleich 5 ist, um die Größe
 ' der OPENFILENAME-Struktur passend angeben zu können.
 Dim ByteBuffer() As Byte
 Dim lBufferSize As Long
 Dim lpBuffer As Long
 Dim lMajorVersion As Long
 Dim lDummy As Long
 ' Größe der Versionsinformationen ermitteln
 lBufferSize = GetFileVersionInfoSize("Comdlg32.dll", lDummy)
 If lBufferSize > 0 Then
  ' Versionsinformationen abholen
  ReDim ByteBuffer(0 To lBufferSize - 1)
  Call GetFileVersionInfo("Comdlg32.dll", 0&, lBufferSize, ByteBuffer(0))
  ' Versionsinformationen auswerten
  If VerQueryValue(ByteBuffer(0), "\", lpBuffer, lDummy) Then
   ' Hauptversionsnummer in lMajorVersion kopieren
   CopyMemory lMajorVersion, ByVal lpBuffer + 10&, 2&
   Win2000Shell = (lMajorVersion >= 5)
  End If
 End If

 Exit Function

Win2000Shell_Error:

 MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure" & Chr(13) & Chr(10) & "(Win2000Shell) of (Cmdlg2.bas).", vbCritical

End Function


'This is the Mainroutine
'Every time cmdlg does anything it calls this routine
Private Function CmdlgHook(ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

 On Error GoTo CmdlgHook_Error
 Dim X As Long
 Dim Y As Long
 Dim nWidth As Long
 Dim nHeight As Long
 Dim Re As RECT
 Dim FName As String
 Dim Buffer As String
 Dim NullPos As Long

 'Get the Messages cmdlg sends
 Select Case uMsg

  'We startet Common Dialog
 Case WM_INITDIALOG
  'So lets resize Comondialog to put our Frame on it
  CdlgHwnd = GetParent(hWnd) 'Get Adress
  GetWindowRect CdlgHwnd, Re 'Get Position as Rect
  'Calculate new size to hold Our preview Form
With Re
 nWidth = .Right - .Left
 nHeight = .Bottom - .Top + 210
End With
X = ((Screen.Width \ Screen.TwipsPerPixelX) - nWidth) \ 2
Y = ((Screen.Height \ Screen.TwipsPerPixelY) - nHeight) \ 2
'Stretch Common Dialog
MoveWindow CdlgHwnd, X, Y, nWidth, nHeight, True
'Now Place our Window over Common Dialog
'10 Pixels Border from the new place we createt on Common Dialog
'FrmCmdlg.Show
'FrmCmdlg.Enabled = False
'CdlgHook = 1
'Set the position for the preview
'FrmCmdlg.ChkPreview = CMPrev
SetWindow


'We got Something
 Case WM_NOTIFY
'  'Get the Filename
'  Buffer = String$(260, 0)
'  NullPos = SendMessage(CdlgHwnd, CDM_GETFILEPATH, 260, ByVal Buffer)
'  If NullPos = -1 Then Exit Function  'So we havnt any directory selectet
'  FName = Left$(Buffer, NullPos - 1)
'  'Get the Path
'  Buffer = String$(260, 0)
'  NullPos = SendMessage(CdlgHwnd, CDM_GETFOLDERPATH, 260, ByVal Buffer)
'  Buffer = Left$(Buffer, NullPos - 1)
'  'Test if path not the same as filename
'  'This routine is not the best but it works
'  If Buffer <> FName And LenB(Dir$(FName)) <> 0 Then
'   'Show the Picture
'   FrmTileCreator.ShowPreview FName
'  End If
TestForPic

 Case WM_DESTROY
'Unload FrmCmdlg
 End Select


Exit Function

CmdlgHook_Error:

MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure" & Chr(13) & Chr(10) & "(CmdlgHook) of (Cmdlg2.bas).", vbCritical

End Function
Public Sub TestForPic()

 On Error GoTo TestForPic_Error
 Dim FName As String
 Dim Buffer As String
 Dim NullPos As Long


 Buffer = String$(260, 0)
 NullPos = SendMessage(CdlgHwnd, CDM_GETFILEPATH, 260, ByVal Buffer)
 If NullPos = -1 Then Exit Sub  'So we havnt any directory selectet
 FName = Left$(Buffer, NullPos - 1)
 'Get the Path
 Buffer = String$(260, 0)
 NullPos = SendMessage(CdlgHwnd, CDM_GETFOLDERPATH, 260, ByVal Buffer)
 Buffer = Left$(Buffer, NullPos - 1)
 'Test if path not the same as filename
 'This routine is not the best but it works
 If Buffer <> FName And LenB(Dir$(FName)) <> 0 Then
  'Show the Picture
  'FrmTileCreator.ShowPreview FName
 End If

 Exit Sub

TestForPic_Error:

 MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure" & Chr(13) & Chr(10) & "(TestForPic) of (Cmdlg2.bas).", vbCritical

End Sub
'Move the preview window to its position
'Put it over common dialog
Public Sub SetWindow()

 On Error GoTo SetWindow_Error

 Dim Re As RECT
 'Get the Size of Cmdlg
 GetWindowRect CdlgHwnd, Re
 'Resize our PreviewWindow
 'MoveWindow FrmCmdlg.hwnd, Re.Left + 10, Re.Bottom - 110, Re.Right - Re.Left - 20, 100, True
 'MoveWindow FrmCmdlg.hwnd, (Screen.Width * Screen.TwipsPerPixelX - Re.Right + Re.Left) / 2, (Screen.Height * Screen.TwipsPerPixelY - Re.Bottom + Re.Top) / 2, Re.Right - Re.Left - 20, 100, True
 'Put it over Cmdlg
 'SetWindowPos CdlgHwnd, FrmCmdlg.hwnd, Re.Left, Re.Top, Re.Right - Re.Left, Re.Bottom - Re.Top, 0 ' HWND_TOP
 'SetParent FrmCmdlg.hWnd, CdlgHwnd
 'MoveWindow FrmCmdlg.hWnd, 10, Re.Bottom - Re.Top - 250, Re.Right - Re.Left - 20, 200, True
 'FrmCmdlg.PicPreview.Width = FrmCmdlg.ScaleWidth - FrmCmdlg.PicPreview.Left

 Exit Sub

SetWindow_Error:

 MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure" & Chr(13) & Chr(10) & "(SetWindow) of (Cmdlg2.bas).", vbCritical

End Sub

Public Function Replace(sIn As String, sFind As String, sReplace As String, Optional nStart As Long = 1, Optional nCount As Long = -1, Optional bCompare As VbCompareMethod = vbBinaryCompare) As String

 On Error GoTo Replace_Error

 Dim nC As Long
 Dim nPos As Long
 Dim sOut As String
 sOut = sIn
 nPos = InStr(nStart, sOut, sFind, bCompare)
 nStart = nPos + Len(sReplace) + 1
If nPos = 0 Then GoTo EndFn:
Do
nC = nC + 1
sOut = Left(sOut, nPos - 1) & sReplace & Mid(sOut, nPos + Len(sFind))
If nCount <> -1 And nC >= nCount Then Exit Do
nPos = InStr(nStart, sOut, sFind, bCompare)
nStart = nPos + Len(sReplace) + 1
Loop While nPos > 0
EndFn:
 Replace = sOut

 Exit Function

Replace_Error:

 MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure" & Chr(13) & Chr(10) & "(Replace) of (Cmdlg2.bas).", vbCritical

End Function

