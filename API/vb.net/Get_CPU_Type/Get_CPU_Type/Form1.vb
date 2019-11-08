Option Explicit On


Public Class Form1
    Private Structure SYSTEM_INFO
        Dim dwOemID As Integer                     '為了兼容性而保留的過時成員
        Dim dwPageSize As Integer                  '頁面大小以及頁面保護和承諾的粒度
        Dim lpMinimumApplicationAddress As Integer '指向應用程序和動態鏈接庫(DLL)可訪問的最低內存位址的指針
        Dim lpMaximumApplicationAddress As Integer '指向應用程序和DLL可訪問的最高內存地址的指針
        Dim dwActiveProcessorMask As Integer       '掩碼，表示配置到系統中的一組處理器。0=處理器0,31=處理器31
        Dim dwNumberOrfProcessors As Integer       '當前邏輯處理器的數量
        Dim dwProcessorType As Integer             '為了兼容性而保留的成員，(386 486 586 2200 8664)
        Dim dwAllocationGranularity As Integer     '可以分配虛擬內存的起始地址的粒度
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
