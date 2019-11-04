Imports System.Runtime.InteropServices
Imports System.Text

Public Class Form1
    <DllImport("kernel32.dll", SetLastError:=True, EntryPoint:="GetSystemDirectoryW", CharSet:=CharSet.Unicode)>
    Shared Function GetSystemDirectory(<MarshalAs(UnmanagedType.LPTStr)> lpBuffer As System.Text.StringBuilder,
                   uSize As UInteger) As UInteger
    End Function

    Shared Function wins() As String
        Dim sb As StringBuilder = New StringBuilder(100)
        Dim f As UInteger = 100
        GetSystemDirectory(sb, f)
        Return sb.ToString()
    End Function
    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        TextBox1.Text = wins()
    End Sub
End Class
