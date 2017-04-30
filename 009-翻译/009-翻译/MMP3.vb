Module MMP3
    Private Declare Function mciSendStringA Lib "winmm.dll" _
        (ByVal lpstrCommand As String, ByVal lpstrReturnString As String,
        ByVal uReturnLength As Integer, ByVal hwndCallback As Integer) As Integer

    Function PlayMidiFile(ByVal MusicFile As String) As Boolean
        mciSendStringA("stop music", "", 0, 0)
        mciSendStringA("close music", "", 0, 0)
        mciSendStringA("open " & MusicFile & " alias music", "", 0, 0)
        PlayMidiFile = mciSendStringA("play music", "", 0, 0) = 0
    End Function

    Private Function StopMidi() As Boolean
        StopMidi = mciSendStringA("stop music", "", 0, 0) = 0
        mciSendStringA("close music", "", 0, 0)
    End Function

    Private Function PauseMidi() As Boolean
        Return mciSendStringA("pause music", "", 0, 0) = 0
    End Function

    Private Function ContinueMidi() As Boolean
        Return mciSendStringA("play music", "", 0, 0) = 0
    End Function
End Module
