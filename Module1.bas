Attribute VB_Name = "Module1"
Declare Function sndPlaySound Lib "winmm" Alias _
"sndPlaySoundA" (ByVal lpszSoundName As String, _
ByVal uFlags As Long) As Long

Public Const SND_SYNC = &H0 ' play synchronously (default)
Public Const SND_ASYNC = &H1 ' play asynchronously
Public Const SND_LOOP = &H8 ' loop the sound until next



