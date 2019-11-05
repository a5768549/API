Imports System.Runtime.InteropServices
Imports System.Text

Public Class Form1
    <DllImport("kernel32", SetLastError:=True)>
    Private Shared Function WritePrivateProfileString(
                        ByVal lpAppName As String,
                        ByVal lpKeyName As String,
                        ByVal lpString As String,
                        ByVal lpFileName As String) As Boolean
    End Function

    <DllImport("kernel32", SetLastError:=True, EntryPoint:="GetPrivateProfileIntA")>
    Public Shared Function GetPrivateProfileInt(
                        ByVal lpApplicationName As String,
                        ByVal lpKeyName As String,
                        ByVal nDefault As Integer,
                        ByVal lpFileName As String) As Integer
    End Function

    <DllImport("kernel32.dll", SetLastError:=True)>
    Private Shared Function GetPrivateProfileString(ByVal lpAppName As String,
                            ByVal lpKeyName As String,
                            ByVal lpDefault As String,
                            ByVal lpReturnedString As StringBuilder,
                            ByVal nSize As Integer,
                            ByVal lpFileName As String) As Integer
    End Function


    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim bres As Boolean
        bres = WritePrivateProfileString("test1", "value1", CStr(TextBox1.Text), Application.StartupPath & "\config.ini")
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        TextBox2.Text = GetPrivateProfileInt("test1", "value1", 0, Application.StartupPath & "\config.ini")
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Dim res As Integer
        Dim sb As StringBuilder = New StringBuilder(500)
        res = GetPrivateProfileString("test1", "value1", "", sb, sb.Capacity, Application.StartupPath & "\config.ini")
        TextBox3.Text = sb.ToString()
    End Sub
End Class
