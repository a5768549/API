Imports System.Management

Public Class Form1
    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Try
            Dim searcher As New ManagementObjectSearcher("root\cimv2", "SELECT * FROM Win32_OperatingSystem")
            For Each queryObj As ManagementObject In searcher.Get()
                Dim WindowsDirectory As String = queryObj("WindowsDirectory")
                TextBox1.Text = ""
                TextBox1.Text = "WindowsDirectory:" & WindowsDirectory
            Next
        Catch err As ManagementException
            TextBox1.Text = ""
            TextBox1.Text = "An error occurred while querying for WMI data: " & err.Message
        End Try
    End Sub
End Class
