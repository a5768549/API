Public Class Form1
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        file_Write(TextBox1.Text, TextBox2.Text, CheckBox1.Checked)
        Console.WriteLine(file_Read(TextBox1.Text))
    End Sub

    Function file_Read(Path As String)
        Dim strTemp As String
        Dim hfile As Integer = FreeFile()
        FileOpen(hfile, Path, OpenMode.Input)
        Do Until EOF(hfile)
            strTemp &= LineInput(hfile) & vbNewLine
        Loop
        FileClose(hfile)
        Return strTemp
    End Function

    Function file_Write(Path As String, text As String, mode As Boolean)
        Dim strTemp As String
        Dim hfile As Integer = FreeFile()
        Dim open_mode As OpenMode
        If mode Then
            open_mode = OpenMode.Output
        Else
            open_mode = OpenMode.Append
        End If
        FileOpen(hfile, Path, open_mode)
        PrintLine(hfile, text)
        FileClose(hfile)
    End Function
End Class
