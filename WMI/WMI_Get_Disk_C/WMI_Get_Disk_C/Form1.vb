Imports System.Management

Public Class Form1
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Try
            Dim searcher As New ManagementObjectSearcher("root\cimv2", "SELECT * FROM Win32_LogicalDisk where DeviceID='C:'")
            For Each queryObj As ManagementObject In searcher.Get()
                Dim DeviceID As String = queryObj("DeviceID")
                Dim FileSystem As String = queryObj("FileSystem")
                Dim FreeSpace As String = queryObj("FreeSpace")
                Dim Size As String = queryObj("Size")
                Dim VolumeSerialNumber As String = queryObj("VolumeSerialNumber")
                TextBox1.Text = DeviceID
                TextBox2.Text = FileSystem
                TextBox3.Text = Size & " Byte"
                TextBox4.Text = FreeSpace & " Byte"
                TextBox5.Text = VolumeSerialNumber
                TextBox6.Text = Size / 1024 / 1024 / 1024 & " GB"
                TextBox7.Text = FreeSpace / 1024 / 1024 / 1024 & " GB"
            Next
        Catch err As ManagementException
            TextBox1.Text = ""
            TextBox1.Text = "An error occurred while querying for WMI data: " & err.Message
        End Try
    End Sub
End Class
