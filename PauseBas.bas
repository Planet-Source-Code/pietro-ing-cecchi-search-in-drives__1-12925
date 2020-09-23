Attribute VB_Name = "PauseBas"
'++++++++++++++++++++++ PAUSE MILLISECONDS ++++++++++++++++++++++
Public Function Pause(ByVal milliseconds As Double) As Integer
'pauses milliseconds
'pause may be interrupted making the global ExitPause true
'returns: vbOK=1 vbAbort=3
Dim Start, Finish, TotalTime
    'interruption from now on only
    ExitPause = False
    Start = Timer * 1000 ' Set start time.
    Do While Timer * 1000 < Start + milliseconds
        DoEvents    ' Yield to other processes.
        If Timer * 1000 < Start Then 'trepassing midnight
          Start = Start - 24 * 60 * 60 * 1000
        End If
        'interruption
        If ExitPause Then ExitPause = False: Pause = vbAbort: Exit Function
    Loop
    Pause = vbOK
End Function

