Imports System.Management
Imports System.Runtime.InteropServices

Public Class Form1

    ' Private memoryInfo As MEMORYSTATUSEX = New MEMORYSTATUSEX
    Private Declare Auto Sub GlobalMemoryStatusEx Lib "kernel32" (<[In](), Out()> lpBuffer As MEMORYSTATUSEX)

    ' Private mullTotalRAM As ULong


    Private Sub Form1_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load

    End Sub


    Private Function GetMemoryInfo(Type As String, Format As Boolean)
        Dim memoryInfo As MEMORYSTATUSEX = New MEMORYSTATUSEX
        GlobalMemoryStatusEx(memoryInfo)
        Dim tmp As String
        Select Case Type
            Case "dwMemoryLoad" : tmp = memoryInfo.dwMemoryLoad
            Case "ullTotalPhys" : tmp = memoryInfo.ullTotalPhys
            Case "ullAvailPhys" : tmp = memoryInfo.ullAvailPhys
            Case "ullTotalPageFile" : tmp = memoryInfo.ullTotalPageFile
            Case "ullAvailPageFile" : tmp = memoryInfo.ullAvailPageFile
            Case "ullTotalVirtual" : tmp = memoryInfo.ullTotalVirtual
            Case "ullAvailVirtual" : tmp = memoryInfo.ullAvailVirtual
            Case "ullAvailExtendedVirtual" : tmp = memoryInfo.ullAvailExtendedVirtual
        End Select
        If Format Then
            If Type <> "dwMemoryLoad" Then
                tmp = FormatBytes(tmp)
            End If
        End If
        Return tmp
    End Function

    Private Function FormatBytes(ByVal ullBytes As ULong) As String '自動轉換單位
        Dim dblTemp As Double
        Try
            Select Case ullBytes
                Case Is >= 1073741824 'GB
                    dblTemp = CDbl(ullBytes / 1073741824)
                    Return FormatNumber(dblTemp, 2) & " GB"
                Case 1048576 To 1073741823
                    dblTemp = CDbl(ullBytes / 1048576) 'MB
                    Return FormatNumber(dblTemp, 0) & " MB"
                Case 1024 To 1048575
                    dblTemp = CDbl(ullBytes / 1024) 'KB
                    Return FormatNumber(dblTemp, 0) & " KB"
                Case 0 To 1023
                    dblTemp = ullBytes ' bytes
                    Return FormatNumber(dblTemp, 0) & " Bytes"
                Case Else
                    Return ""
            End Select
        Catch
            Return ""
        End Try
    End Function

    Private Sub Timer1_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick
        TextBox1.Text = GetMemoryInfo("dwMemoryLoad", False) & " %"
        TextBox2.Text = GetMemoryInfo("ullTotalPhys", True)
        TextBox3.Text = GetMemoryInfo("ullAvailPhys", True)
        TextBox4.Text = GetMemoryInfo("ullTotalPageFile", True)
        TextBox5.Text = GetMemoryInfo("ullAvailPageFile", True)
        TextBox6.Text = GetMemoryInfo("ullTotalVirtual", True)
        TextBox7.Text = GetMemoryInfo("ullAvailVirtual", True)
        TextBox8.Text = GetMemoryInfo("ullAvailExtendedVirtual", True)
    End Sub
End Class

<StructLayout(LayoutKind.Sequential, CharSet:=CharSet.Auto)>
Public Class MEMORYSTATUSEX

    ' Initializes a new instance of the <see cref="T:MEMORYSTATUSEX" /> class. 
    Public Sub New()
        Me.dwLength = CType(Marshal.SizeOf(GetType(MEMORYSTATUSEX)), UInt32)
    End Sub

    Public dwLength As UInt32 ' 結構的大小，以Byte為單位。您必須在調用GlobalMemoryStatusEx之前設置此成員。
    Public dwMemoryLoad As UInt32 ' 0到100之間的數字，指定正在使用的物理內存的大約百分比（0表示未使用內存，而100表示​​已使用全部內存）。
    Public ullTotalPhys As UInt64 ' 物理內存的總大小，以Byte為單位。
    Public ullAvailPhys As UInt64 ' 可用物理內存的大小，以Byte為單位。
    Public ullTotalPageFile As UInt64 ' 提交的內存限制的大小，以Byte為單位。這是物理內存加上分頁檔(虛擬記憶體)的大小，再減去少量開銷。
    Public ullAvailPageFile As UInt64 ' 提交的可用內存大小，以Byte為單位。限制為ullTotalPageFile。
    Public ullTotalVirtual As UInt64 ' 調用進程的虛擬地址空間的用戶模式部分的總大小，以Byte為單位。
    Public ullAvailVirtual As UInt64 ' 調用進程的虛擬地址空間的用戶模式部分中未保留和未提交的內存大小，以Byte為單位。
    Public ullAvailExtendedVirtual As UInt64 ' 調用進程的虛擬地址空間擴展部分中未保留和未提交的內存大小，以Byte為單位。 
End Class