Attribute VB_Name = "ModSubs"
'Some Subs for WordPuz
'Not all mine

Option Explicit

Public Type HighScore           'Holds the HighScore
 Score As Long
 Name As String
End Type

Public Lng(14) As String        'Language
Public Language As String       'Language Name
Public GamePoints As Long       'Points
Public HSC(9) As HighScore      'Highsore list


'Cant use Long names for midi
Private Declare Function GetShortPathName Lib "kernel32" Alias "GetShortPathNameA" (ByVal lpszLongPath As String, ByVal lpszShortPath As String, ByVal cchBuffer As Long) As Long

'Sound
Public Declare Function mciSendString Lib "winmm.dll" Alias "mciSendStringA" (ByVal lpstrCommand As String, ByVal lpstrReturnString As String, ByVal uReturnLength As Long, ByVal hwndCallback As Long) As Long
Private Declare Function sndPlaySound Lib "winmm" Alias "sndPlaySoundA" (lpszSoundName As Any, ByVal uFlags As Long) As Long
Private Const SND_SYNC = &H0        ' Play synchronously (default).
Private Const SND_NODEFAULT = &H2    ' Do not use default sound.
Private Const SND_MEMORY = &H4
Private Const SND_LOOP = &H8         ' Loop the sound until next
Private Const SND_NOSTOP = &H10      ' Do not stop any currently
Private Const SND_ASYNC = &H1          '  play asynchronously
Private bytSound() As Byte ' Always store binary data in byte arrays!
Public Enum SoundFlags
soundSYNC = SND_SYNC
soundNO_DEFAULT = SND_NODEFAULT
soundMEMORY = SND_MEMORY
soundLOOP = SND_LOOP
soundNO_STOP = SND_NOSTOP
soundASYNC = SND_ASYNC
End Enum

'Play Wave from resourcefile
Public Sub PlayWaveRes(vntResourceID As Long, Optional vntFlags As SoundFlags = soundASYNC)
 bytSound = LoadResData(vntResourceID, "WAVE")
 If IsMissing(vntFlags) Then
  vntFlags = SND_NODEFAULT Or SND_SYNC Or SND_MEMORY
 End If
 If (vntFlags And SND_MEMORY) = 0 Then
  vntFlags = vntFlags Or SND_MEMORY
 End If
 sndPlaySound bytSound(0), vntFlags
End Sub
'Play a midisong
Public Sub PlayMIDI(MIDIFile As String)
 Dim SafeFile As String
 SafeFile$ = Dir(MIDIFile$)
 If SafeFile$ <> "" Then
  Call mciSendString("play " & MIDIFile$, 0&, 0, 0)
 End If
End Sub
'Stop and unload Midisong
Public Sub StopMIDI(MIDIFile As String)
 Dim SafeFile As String
 SafeFile$ = Dir(MIDIFile$)
 If SafeFile$ <> "" Then
  Call mciSendString("stop " & MIDIFile$, 0&, 0, 0)
  mciSendString "close " & MIDIFile$, 0, 0, 0
 End If
End Sub
'Test if song is still playing
Public Function StillMidi(MIDIFile As String) As Boolean
 Dim Y As String * 255
 mciSendString "Status " & MIDIFile & " position", Y, 255, 0
 If Val(Y) <> 0 Then StillMidi = True
End Function
'Convert long to short filenames
Public Function GetShortName(ByVal sLongFileName As String) As String
 Dim lRetVal As Long, sShortPathName As String, iLen As Integer
 'Set up buffer area for API function call return
 sShortPathName = Space(255)
 iLen = Len(sShortPathName)

 'Call the function
 lRetVal = GetShortPathName(sLongFileName, sShortPathName, iLen)
 'Strip away unwanted characters.
 GetShortName = Left(sShortPathName, lRetVal)
End Function

'Scramble the text for our Board
Public Function Scramble(Text As String) As String
 'This will scramble text up, example:  oCol
 On Error GoTo error
 Dim RndNum As Long
 Dim i As Long
 Dim endstr As String
 Dim ListN(10000) As Long
 Dim CurPos As Long
 Randomize Time
 CurPos = 0
 Text$ = Mid$(Text$, 1, 10000)
Start:
 RndNum = Int((Len(Text$) - 1 + 1) * Rnd + 1)
 For i = 0 To CurPos
  If RndNum = ListN(i) Then
   GoTo Start
  End If
 Next i
 ListN(CurPos) = RndNum
 CurPos = CurPos + 1
 If Not CurPos = Len(Text$) Then
  GoTo Start
 End If
 For i = 0 To CurPos - 1
  endstr$ = endstr$ & Mid$(Text$, ListN(i), 1)
 Next i
 Scramble = endstr$
 Exit Function
error:                MsgBox Err.Description, vbExclamation, "Error"
End Function

'Sort our Strings
Public Sub BubbleSort(SortString() As String)
 Dim LB As Integer
 Dim UB As Integer
 Dim X As Integer
 Dim Y As Integer
 Dim Tmp As Variant

 ' Get the bounds of the array
 LB = LBound(SortString)
 UB = UBound(SortString)

 'Sort Alpha
 ' For each element in the array
 For X = LB To UB - 1
  ' for each element in the array
  For Y = X + 1 To UB
   If SortString(X) > SortString(Y) Then
    Tmp = SortString(X)
    SortString(X) = SortString(Y)
    SortString(Y) = Tmp
   End If
  Next Y
 Next X

 'Sort Size
 For X = LB To UB - 1
  For Y = X + 1 To UB
   If LenB(SortString(X)) > LenB(SortString(Y)) Then
    Tmp = SortString(X)
    SortString(X) = SortString(Y)
    SortString(Y) = Tmp
   End If
  Next Y
 Next X

 For X = 0 To UB
  SortString(X) = SortString(X)
 Next X
End Sub

