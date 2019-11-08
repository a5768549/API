Imports System.Runtime.InteropServices
Imports System.Text

Public Class Form1
    Declare Function GetTempPath Lib "kernel32.dll" Alias "GetTempPathA" (ByVal nBufferLength As Integer, ByVal lpBuffer As String) As Integer

    Dim temppath As String
    'Dim tempfile As String  ' receives name of temporary file
    Dim slength As Integer  ' receives length of string returned for the path
    'Dim lastfour As Integer  ' receives hex value of the randomly assigned ????


    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        temppath = Space(255)  ' initialize the buffer to receive the path
        slength = GetTempPath(255, temppath)  ' read the path name
        temppath = Microsoft.VisualBasic.Left(temppath, slength)
        TextBox1.Text = temppath
    End Sub
End Class
