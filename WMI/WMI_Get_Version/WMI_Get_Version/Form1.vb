Imports System.Management

Public Class Form1
    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Try
            Dim searcher As New ManagementObjectSearcher("root\cimv2", "SELECT * FROM Win32_OperatingSystem")
            For Each queryObj As ManagementObject In searcher.Get()
                Dim BuildNumber As String = CDbl(queryObj("BuildNumber")) '建構碼
                Dim Caption As String = queryObj("Caption") '版本名稱
                Dim Version As String = queryObj("Version") '版本號
                TextBox1.Text = ""
                TextBox1.Text &= "BuildNumber：" & BuildNumber & vbNewLine
                TextBox1.Text &= "Caption：" & Caption & vbNewLine
                TextBox1.Text &= "Version：" & Version
                Dim tmp As String = ""
                Select Case BuildNumber
                    Case "7601" : tmp = "Windows 7"
                    Case "9200" : tmp = "Windows 8"
                    Case "9600" : tmp = "Windows 8.1"
                    Case "10240" : tmp = "Windows 10"
                    Case "10586" : tmp = "Windows 10 1511"
                    Case "14393" : tmp = "Windows 10 1607"
                    Case "15063" : tmp = "Windows 10 1703"
                    Case "16299" : tmp = "Windows 10 1709"
                    Case "17134" : tmp = "Windows 10 1083"
                    Case "17763" : tmp = "Windows 10 1809"
                    Case "18362" : tmp = "Windows 10 1903"
                End Select
                TextBox2.Text = tmp
            Next
        Catch err As ManagementException
            TextBox1.Text = ""
            TextBox1.Text = "An error occurred while querying for WMI data: " & err.Message
        End Try


    End Sub
End Class
