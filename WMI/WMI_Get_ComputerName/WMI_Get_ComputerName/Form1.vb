Imports System.Management

Public Class Form1
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Try
            Dim searcher As New ManagementObjectSearcher("root\cimv2", "SELECT * FROM Win32_ComputerSystem")
            For Each queryObj As ManagementObject In searcher.Get()
                Dim Name As String = queryObj("Name")
                TextBox1.Text = ""
                TextBox1.Text = Name
            Next
        Catch err As ManagementException
            TextBox1.Text = ""
            TextBox1.Text = "An error occurred while querying for WMI data: " & err.Message
        End Try
    End Sub
End Class
