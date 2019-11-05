Option Explicit On


Public Class Form1
    Private Structure SYSTEM_INFO
        Dim dwOemID As Integer
        Dim dwPageSize As Integer
        Dim lpMinimumApplicationAddress As Integer
        Dim lpMaximumApplicationAddress As Integer
        Dim dwActiveProcessorMask As Integer
        Dim dwNumberOrfProcessors As Integer
        Dim dwProcessorType As Integer
        Dim dwAllocationGranularity As Integer
        Dim dwReserved As Integer
    End Structure

    Private Declare Sub GetSystemInfo Lib "kernel32" Alias "GetSystemInfo" (ByRef lpSystemInfo As SYSTEM_INFO)

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim cpuInfo As New SYSTEM_INFO()
        GetSystemInfo(cpuInfo)
        TextBox1.Text = "CPU Type = " & cpuInfo.dwProcessorType.ToString()
        TextBox2.Text = "CPU Core = " & cpuInfo.dwNumberOrfProcessors.ToString()
    End Sub
End Class
