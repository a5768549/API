Imports System.ComponentModel
Imports System.Runtime.InteropServices
Imports System.Text

Public Class Form1
    <DllImport("kernel32.dll", CharSet:=CharSet.Auto, SetLastError:=True)>
    Private Shared Function GetComputerNameW(ByVal lpBuffer As StringBuilder, ByRef lpnSize As Integer) As Integer
    End Function

    Dim buffer As StringBuilder = New StringBuilder(512)
    Dim length As Long = 512

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If GetComputerNameW(buffer, length) Then
            ' Returns nonzero if successful, and modifies the length argument
            TextBox1.Text = buffer.ToString
        End If
    End Sub
End Class
