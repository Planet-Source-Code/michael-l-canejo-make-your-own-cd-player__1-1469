VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CDAudio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private Declare Function mciGetErrorString Lib "winmm.dll" Alias "mciGetErrorStringA" (ByVal dwError As Long, ByVal lpstrBuffer As String, ByVal uLength As Long) As Long
Private Declare Function mciSendString Lib "winmm.dll" Alias "mciSendStringA" (ByVal lpstrCommand As String, ByVal lpstrReturnString As String, ByVal uReturnLength As Long, ByVal hwndCallback As Long) As Long

Function StartPlay()
    mciSendString "play cd", 0, 0, 0
End Function

Function SetTrack(Track%)
    mciSendString "seek cd to " & Str(Track), 0, 0, 0
End Function

Function StopPlay()
    mciSendString "stop cd wait", 0, 0, 0
End Function

Function PausePlay()
    mciSendString "pause cd", 0, 0, 0
End Function

Function EjectCD()
    mciSendString "set cd door open", 0, 0, 0
End Function

Function CloseCD()
    mciSendString "set cd door closed", 0, 0, 0
End Function

Function UnloadAll()
    mciSendString "close all", 0, 0, 0
End Function

Function SetCDPlayerReady()
    mciSendString "open cdaudio alias cd wait shareable", 0, 0, 0
End Function
Function SetFormat_tmsf()
    mciSendString "set cd time format tmsf wait", 0, 0, 0
End Function

Function SetFormat_milliseconds()
    mciSendString "set cd time format milliseconds", 0, 0, 0
End Function

Function CheckCD%()
    Dim s As String * 30
    mciSendString "status cd media present", s, Len(s), 0
    CheckCD = s
End Function

Function GetNumTracks%()
    Dim s As String * 30
    mciSendString "status cd number of tracks wait", s, Len(s), 0
    GetNumTracks = CInt(Mid$(s, 1, 2))
End Function

Function GetCDLength$()
    Dim s As String * 30
    mciSendString "status cd length wait", s, Len(s), 0
    GetCDLength = s
End Function

Function GetTrackLength$(TrackNum%)
    Dim s As String * 30
    mciSendString "status cd length track " & TrackNum, s, Len(s), 0
    GetTrackLength = s
End Function

Sub GetCDPosition(Track%, Min%, Sec%)
    Dim s As String * 30
    mciSendString "status cd position", s, Len(s), 0
    Track = CInt(Mid$(s, 1, 2))
    Min = CInt(Mid$(s, 4, 2))
    Sec = CInt(Mid$(s, 7, 2))
End Sub

Function CheckIfPlaying%()
    CheckIfPlaying = 0
    Dim s As String * 30
    mciSendString "status cd mode", s, Len(s), 0
    If Mid$(s, 1, 7) = "playing" Then CheckIfPlaying = 1
End Function

'|---------------------|
'|---Automated Tasks---|
'V---------------------V

Function SeekCDtoX(Track%)
    StopPlay
    SetTrack Track
    StartPlay
End Function

Function ReadyDevice()
    UnloadAll
    SetCDPlayerReady
    SetFormat_tmsf
End Function

Function FastForward(Spd%)
    Dim s As String * 40
    SetFormat_milliseconds
    mciSendString "status cd position wait", s, Len(s), 0
    CheckIfPlaying%
    If CheckIfPlaying = 1 Then
        mciSendString "play cd from " & CStr(CLng(s) + Spd), 0, 0, 0
    Else
        mciSendString "seek cd to " & CStr(CLng(s) + Spd), 0, 0, 0
    End If
    SetFormat_tmsf
End Function

Function ReWind(Spd%)
    Dim s As String * 40
    SetFormat_milliseconds
    mciSendString "status cd position wait", s, Len(s), 0
    CheckIfPlaying%
    If CheckIfPlaying = 1 Then
        mciSendString "play cd from " & CStr(CLng(s) - Spd), 0, 0, 0
    Else
        mciSendString "seek cd to " & CStr(CLng(s) - Spd), 0, 0, 0
    End If
    SetFormat_tmsf
End Function


