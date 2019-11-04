Imports System.Management

Public Class Form1
    Dim tmp() As String
    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Try
            Dim searcher As New ManagementObjectSearcher("root\cimv2", "SELECT * FROM Win32_ComputerSystem")
            For Each queryObj As ManagementObject In searcher.Get()
                Dim UserName As String = queryObj("UserName")
                tmp = UserName.Split("\")
            Next
        Catch err As ManagementException
            TextBox1.Text = ""
            TextBox1.Text = "An error occurred while querying for WMI data: " & err.Message
        End Try
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        TextBox1.Text = tmp(0)
        TextBox2.Text = tmp(1)
    End Sub
End Class
