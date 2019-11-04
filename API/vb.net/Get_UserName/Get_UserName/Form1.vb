Imports System.Runtime.InteropServices
Imports System.Text

Public Class Form1
    <DllImport("advapi32.dll", EntryPoint:="GetUserNameW", SetLastError:=True)>
    Shared Function GetUserName(<MarshalAs(UnmanagedType.LPTStr)> sb As System.Text.StringBuilder,
         ByRef length As Integer) As <MarshalAs(UnmanagedType.Bool)> Boolean
    End Function

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim Buffer = New StringBuilder(64)
        Dim nSize As Integer = 64
        GetUserName(Buffer, nSize)
        TextBox1.Text = Buffer.ToString()

    End Sub
End Class
