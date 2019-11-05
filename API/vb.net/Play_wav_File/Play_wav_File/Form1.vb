Public Class Form1
    Private Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Integer) As Integer

    Private Const SND_ASYNC = &H1         '  play asynchronously
    Private Const SND_LOOP = &H8         '  loop the sound until next sndPlaySound
    Private Const SND_MEMORY = &H4         '  lpszSoundName points to a memory file
    Private Const SND_NODEFAULT = &H2         '  silence not default, if sound not found
    Private Const SND_NOSTOP = &H10        '  don't stop any currently playing sound
    Private Const SND_SYNC = &H0         '  play synchronously (default)


    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim tmp As Integer
        tmp = OutOfTimeSound()
    End Sub

    Function OutOfTimeSound()
        Call sndPlaySound("C:\Users\itplayerCR.ITPLAYER\Music\1.wav", SND_ASYNC)
    End Function
End Class
