Imports System.Runtime.InteropServices

Public Class Form1

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub
    Function open_image(Path As String)
        Return Image.FromFile(Path)
    End Function
    Sub delete_image()
        PictureBox1.Image.Dispose()
        PictureBox1.Image = Nothing
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        PictureBox1.Image = open_image("C:\Users\a5768\Pictures\02 (1).png")
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        delete_image()
    End Sub
End Class
