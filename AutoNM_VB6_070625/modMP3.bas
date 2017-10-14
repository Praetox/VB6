Attribute VB_Name = "modMP3"
Function mLoad(ByVal file As String)
    frmMain.wmp.URL = file
    frmMain.wmp.Controls.Stop
    DoEvents
End Function

Function mPlay()
    frmMain.wmp.Controls.play
End Function
Function mPause()
    frmMain.wmp.Controls.pause
End Function
Function mStop()
    frmMain.wmp.Controls.Stop
End Function

Function mPlaying() As Boolean
    If frmMain.wmp.playState = wmppsPlaying Then mPlaying = True
End Function
