
Public Class Form1
    Private Declare Function GetVersion Lib "kernel32" () As Long
    Public Function GetWinVer() As String
        Dim lngVer As Long, lngWinVer As Long
        lngVer = GetVersion()
        lngWinVer = lngVer And &HFFFF&
        ' 取得 Windows 版本
        GetWinVer = Format((lngWinVer Mod 256) + ((lngWinVer \ 256) / 100), "Fixed")
    End Function

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        TextBox1.Text = GetWinVer()
        TextBox2.Text = GetVersion()
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If GetVersion() And &H80000000 Then
            MsgBox("Running under Windows 95/98")
        Else
            MsgBox("Running under Windows NT")
        End If
    End Sub
End Class