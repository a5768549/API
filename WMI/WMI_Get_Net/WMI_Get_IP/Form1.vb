Imports System.Management

Public Class Form1
    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        TextBox1.Text = ""
        Try
            Dim searcher As New ManagementObjectSearcher("root\cimv2", "SELECT * FROM Win32_NetworkAdapterConfiguration WHERE IPEnabled = TRUE")
            For Each queryObj As ManagementObject In searcher.Get()
                Dim IPAddress As String() = queryObj("IPAddress")
                TextBox1.Text &= "IPAddress:" & IPAddress(0) & vbNewLine
                TextBox2.Text &= "IPAddress:" & IPAddress(1) & vbNewLine
                Dim MACAddress As String = queryObj("MACAddress")
                TextBox3.Text &= "MACAddress:" & MACAddress & vbNewLine
                Dim Description As String = queryObj("Description")
                TextBox4.Text &= "Description:" & Description & vbNewLine
            Next
        Catch err As ManagementException
            TextBox1.Text = ""
            TextBox1.Text = "An error occurred while querying for WMI data: " & err.Message
        End Try
    End Sub
End Class
