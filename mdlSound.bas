Attribute VB_Name = "mdlSound"
Public Const SND_ASYNC = &H1
Public Declare Function sndPlaySound Lib "winmm.dll" Alias _
    "sndPlaySoundA" (ByVal lpszSoundName As String, _
    ByVal uFlags As Long) As Long

Public Declare Function mciSendString Lib "winmm.dll" Alias _
    "mciSendStringA" (ByVal lpstrCommand As String, ByVal _
    lpstrReturnString As String, ByVal uReturnLength As Long, _
    ByVal hwndCallback As Long) As Long

Sub playIntro()
    Call sndPlaySound(AddASlash(App.Path) + "intro.wav", SND_ASYNC)
End Sub

Sub stopAll()
    Call sndPlaySound(0, SND_ASYNC)
End Sub
