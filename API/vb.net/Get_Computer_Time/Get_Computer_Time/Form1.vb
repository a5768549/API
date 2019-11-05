Option Explicit On
Public Class Form1

    '最高可計算52天開機時間
    Private Declare Function GetTickCount Lib "kernel32" () As Integer

    Private Sub Timer1_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick
        Dim lngMS As Long
        lngMS = GetTickCount
        TextBox1.Text = lngMS & " milliseconds" & vbNewLine
        TextBox1.Text = TextBox1.Text & lngMS / 1000 & " Seconds" & vbNewLine
        TextBox1.Text = TextBox1.Text & lngMS / 1000 / 60 & " Minute" & vbNewLine
        TextBox1.Text = TextBox1.Text & lngMS / 1000 / 60 / 60 & " Hour" & vbNewLine
        TextBox1.Text = TextBox1.Text & lngMS / 1000 / 60 / 60 / 24 & " Day" & vbNewLine
    End Sub
End Class
