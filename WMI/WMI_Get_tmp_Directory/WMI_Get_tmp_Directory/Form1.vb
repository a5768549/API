Imports System.Management

Public Class Form1
    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        TextBox1.Text = ""
        If CheckBox1.Checked = True Then
            TextBox1.Text = Environment.UserDomainName & "\\" & Environment.UserName
        Else
            TextBox1.Text = Environment.UserName
        End If

        Try
            Dim WQL_Query As String = "SELECT * FROM Win32_Environment where Name='TEMP' and UserName='" & TextBox1.Text & "'"
            Dim searcher As New ManagementObjectSearcher("root\cimv2", WQL_Query)
            For Each queryObj As ManagementObject In searcher.Get()
                Dim Caption As String = queryObj("Caption")
                TextBox2.Text = ""
                TextBox2.Text = "Caption:" & Caption
            Next

        Catch err As ManagementException
            TextBox2.Text = ""
            TextBox2.Text = "An error occurred while querying for WMI data: " & err.Message
        End Try
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Dim System_Directory As String
        Dim tmp_Directory As String
        Try
            Dim searcher As New ManagementObjectSearcher("root\cimv2", "SELECT * FROM Win32_OperatingSystem")
            For Each queryObj As ManagementObject In searcher.Get()
                Dim WindowsDirectory As String = queryObj("WindowsDirectory")
                System_Directory = WindowsDirectory
            Next
        Catch err As ManagementException
            TextBox1.Text = ""
            TextBox1.Text = "An error occurred while querying for WMI data: " & err.Message
        End Try

        Try
            Dim WQL_Query As String = "SELECT * FROM Win32_Environment where Name='TEMP' and UserName='<SYSTEM>'"
            Dim searcher As New ManagementObjectSearcher("root\cimv2", WQL_Query)
            For Each queryObj As ManagementObject In searcher.Get()
                Dim Caption As String = queryObj("Caption")
                TextBox3.Text = ""
                TextBox3.Text = "Caption:" & Caption
                tmp_Directory = Caption
            Next

        Catch err As ManagementException
            TextBox3.Text = ""
            TextBox3.Text = "An error occurred while querying for WMI data: " & err.Message
        End Try

        TextBox5.Text = tmp_Directory.Replace("<SYSTEM>", System_Directory)
    End Sub
End Class
