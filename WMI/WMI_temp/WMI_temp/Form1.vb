Imports System.Management

Public Class Form1

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Try
            Dim searcher As New ManagementObjectSearcher("root\WMI", "SELECT * FROM MSAcpi_ThermalZoneTemperature")
            For Each queryObj As ManagementObject In searcher.Get()
                Dim temp As Double = CDbl(queryObj("CurrentTemperature"))
                temp = (temp - 2732) / 10.0
                TextBox1.Text = temp.ToString
            Next
        Catch err As ManagementException
            TextBox1.Text = "An error occurred while querying for WMI data: " & err.Message
        End Try
    End Sub
End Class