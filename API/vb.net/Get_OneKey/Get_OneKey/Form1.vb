Imports System.Runtime.InteropServices

Public Class Form1
    Const VK_LSHIFT = &HA0
    Const VK_LCONTROL = &HA2
    Const VK_LMENU = &HA4
    Const VK_RSHIFT = &HA1
    Const VK_RCONTROL = &HA3
    Const VK_RMENU = &HA5

    Private Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As IntPtr) As Short

    Private Sub Timer1_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick
        Select Case True
            Case GetAsyncKeyState(Keys.LControlKey)
                TextBox1.Text = "Left Control"
            Case GetAsyncKeyState(Keys.RControlKey)
                TextBox1.Text = "Right Control"
            Case GetAsyncKeyState(Keys.LShiftKey)
                TextBox1.Text = "Left Shift"
            Case GetAsyncKeyState(Keys.RShiftKey)
                TextBox1.Text = "Right Shift"
            Case GetAsyncKeyState(Keys.LWin)
                TextBox1.Text = "Left Win"
            Case GetAsyncKeyState(Keys.RWin)
                TextBox1.Text = "Right Win"
            Case GetAsyncKeyState(Keys.LMenu)
                TextBox1.Text = "Left Menu"
            Case GetAsyncKeyState(Keys.RMenu)
                TextBox1.Text = "Right Menu"
        End Select
    End Sub

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub
End Class
