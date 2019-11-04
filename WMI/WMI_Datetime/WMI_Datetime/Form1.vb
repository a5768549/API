Imports System.Management

Public Class Form1

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If Button1.Text = "開啟timer" Then
            Timer1.Enabled = True
            Button1.Text = "關閉timer"
        ElseIf Button1.Text = "關閉timer" Then
            Timer1.Enabled = False
            Button1.Text = "開啟timer"
        End If
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        If Button2.Text = "開啟timer" Then
            Timer2.Enabled = True
            Button2.Text = "關閉timer"
        ElseIf Button2.Text = "關閉timer" Then
            Timer2.Enabled = False
            Button2.Text = "開啟timer"
        End If
    End Sub

    Private Sub Timer1_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick
        Try
            Dim searcher As New ManagementObjectSearcher("root\cimv2", "SELECT * FROM Win32_CurrentTime")
            For Each queryObj As ManagementObject In searcher.Get()
                Dim Year As Double = CDbl(queryObj("Year"))
                Dim Month As Double = CDbl(queryObj("Month"))
                Dim Day As Double = CDbl(queryObj("Day"))
                Dim Hour As Double = CDbl(queryObj("Hour"))
                Dim Minute As Double = CDbl(queryObj("Minute"))
                Dim Second As Double = CDbl(queryObj("Second"))
                TextBox1.Text = ""
                TextBox1.Text &= "Year：" & Year & vbNewLine
                TextBox1.Text &= "Month：" & Month & vbNewLine
                TextBox1.Text &= "Day：" & Day & vbNewLine
                TextBox1.Text &= "Hour：" & Hour & vbNewLine
                TextBox1.Text &= "Minute：" & Minute & vbNewLine
                TextBox1.Text &= "Second：" & Second & vbNewLine

            Next
        Catch err As ManagementException
            TextBox1.Text = ""
            TextBox1.Text = "An error occurred while querying for WMI data: " & err.Message
        End Try
    End Sub

    Private Sub Timer2_Tick(sender As Object, e As EventArgs) Handles Timer2.Tick
        Try
            Dim searcher As New ManagementObjectSearcher("root\cimv2", "SELECT * FROM Win32_OperatingSystem")
            For Each queryObj As ManagementObject In searcher.Get()
                Dim LocalDateTime As String = queryObj("LocalDateTime")
                TextBox2.Text = ""
                TextBox2.Text &= "Year：" & LocalDateTime & vbNewLine


            Next
        Catch err As ManagementException
            TextBox2.Text = ""
            TextBox2.Text = "An error occurred while querying for WMI data: " & err.Message
        End Try
    End Sub
End Class