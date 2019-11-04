Imports System.Runtime.InteropServices

Public Class Form1
    Structure POINTAPI ' This holds the logical cursor information
        Dim x As Integer
        Dim y As Integer
    End Structure

    <DllImport("user32.dll", SetLastError:=True)>
    Private Shared Function SetCursorPos(ByVal X As Integer, ByVal Y As Integer) As Long
    End Function
    <DllImport("user32.dll", SetLastError:=True)>
    Private Shared Function ClientToScreen(ByVal hWnd As IntPtr, ByRef lpPoint As POINTAPI) As Boolean
    End Function

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim lpPoint As POINTAPI
        lpPoint.x = TextBox1.Text 'ScaleX(Command1.Width / 2, vbTwips, vbPixels)
        lpPoint.y = TextBox2.Text 'ScaleY(Command1.Height / 2, vbTwips, vbPixels)
        ' Convert to screen coordinates.
        ClientToScreen(Me.Handle, lpPoint)
        ' Move the mouse cursor to that point.
        SetCursorPos(lpPoint.x, lpPoint.y)
    End Sub
End Class
