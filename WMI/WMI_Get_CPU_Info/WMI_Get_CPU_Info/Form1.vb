Imports System.Management

Public Class Form1
    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Try
            Dim searcher As New ManagementObjectSearcher("root\cimv2", "SELECT * FROM Win32_Processor")
            For Each queryObj As ManagementObject In searcher.Get()
                Dim AddressWidth As String = queryObj("AddressWidth")
                Dim Description As String = queryObj("Description")
                Dim L2CacheSize As String = queryObj("L2CacheSize")
                Dim L3CacheSize As String = queryObj("L3CacheSize")
                Dim Name As String = queryObj("Name")
                Dim NumberOfCores As String = queryObj("NumberOfCores")
                Dim ProcessorId As String = queryObj("ProcessorId")
                Dim SocketDesignation As String = queryObj("SocketDesignation")
                Dim CurrentClockSpeed As String = queryObj("CurrentClockSpeed")
                Dim ExtClock As String = queryObj("ExtClock")
                TextBox1.Text = "" : TextBox1.Text = Name
                TextBox2.Text = "" : TextBox2.Text = Description
                TextBox3.Text = "" : TextBox3.Text = SocketDesignation
                TextBox4.Text = "" : TextBox4.Text = AddressWidth
                TextBox6.Text = "" : TextBox6.Text = L2CacheSize
                TextBox7.Text = "" : TextBox7.Text = L3CacheSize
                TextBox8.Text = "" : TextBox8.Text = ProcessorId
                TextBox9.Text = "" : TextBox9.Text = NumberOfCores
                TextBox10.Text = "" : TextBox10.Text = CurrentClockSpeed
                TextBox11.Text = "" : TextBox11.Text = ExtClock

            Next
        Catch err As ManagementException
            TextBox1.Text = ""
            TextBox1.Text = "An error occurred while querying for WMI data: " & err.Message
        End Try
    End Sub

    Private Sub Timer1_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick
        Try
            Dim searcher As New ManagementObjectSearcher("root\cimv2", "SELECT * FROM Win32_Processor")
            For Each queryObj As ManagementObject In searcher.Get()

                Dim CurrentVoltage As String = queryObj("CurrentVoltage")
                Dim LoadPercentage As String = queryObj("LoadPercentage")
                TextBox5.Text = LoadPercentage
                TextBox12.Text = CurrentVoltage / 10

            Next
        Catch err As ManagementException
            TextBox1.Text = ""
            TextBox1.Text = "An error occurred while querying for WMI data: " & err.Message
        End Try
    End Sub
End Class
