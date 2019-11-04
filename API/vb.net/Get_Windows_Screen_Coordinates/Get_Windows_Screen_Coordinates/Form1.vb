Imports System.Runtime.InteropServices

Public Class Form1
    Structure POINTAPI ' This holds the logical cursor information
        Dim x As Integer
        Dim y As Integer
    End Structure
    <DllImport("user32.dll", SetLastError:=True)>
    Private Shared Function ScreenToClient(ByVal hWnd As IntPtr, ByRef lpPoint As POINTAPI) As Boolean
    End Function

    ' Display Windows screen coordinates relative to current form.
    Private Sub Timer1_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer1.Tick
        Dim rect As POINTAPI
        ScreenToClient(Handle, rect)
        TextBox1.Text = "x:" & rect.x * -1
        TextBox2.Text = "y:" & rect.y * -1
    End Sub

End Class