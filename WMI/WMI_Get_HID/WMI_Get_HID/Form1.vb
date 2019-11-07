Imports System
Imports System.Management

Public Class Form1
    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub
    Dim InstanceName As String
    Dim HID_Device
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        ListBox1.Items.Clear()

        Try
            Dim searcher As New ManagementObjectSearcher("root\WMI", "SELECT * FROM MSWmi_PnPInstanceNames")
            For Each queryObj As ManagementObject In searcher.Get()
                InstanceName = queryObj("InstanceName")
                Dim tmp = InstanceName.Split("\")
                HID_Device = tmp
                If tmp(0) = "HID" Then
                    ListBox1.Items.Add(InstanceName)
                End If

            Next
        Catch err As ManagementException
        End Try
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Dim selected = ListBox1.Items(ListBox1.SelectedIndex)
        Dim tmp As String() = selected.Split("\")

        Dim ID As String() = tmp(1).Split("&")
        Dim VID As String() = ID(0).Split("_")
        Dim PID As String() = ID(1).Split("_")
        TextBox1.Text = VID(1)
        TextBox2.Text = PID(1)
    End Sub
End Class
