Public Class Form1
    Private Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Integer) As Integer

    Private Const SND_ASYNC = &H1         '異步播放
    Private Const SND_LOOP = &H8          '循環播放聲音，直到下一個sndPlaySound
    Private Const SND_MEMORY = &H4        'lpszSoundName指向內存文件
    Private Const SND_NODEFAULT = &H2     '無默認值,如果未找到聲音
    Private Const SND_NOSTOP = &H10       '不停止任何當前正在播放的聲音
    Private Const SND_SYNC = &H0          '同步播放（默認）


    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim tmp As Integer
        tmp = OutOfTimeSound()
    End Sub

    Function OutOfTimeSound()
        Call sndPlaySound("C:\Users\itplayerCR.ITPLAYER\Music\1.wav", SND_ASYNC)
    End Function
End Class
