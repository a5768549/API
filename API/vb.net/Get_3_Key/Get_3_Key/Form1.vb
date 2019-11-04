Public Class Form1
    Private Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As IntPtr) As Short

    Private Sub Timer1_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick
        If GetAsyncKeyState(Keys.LControlKey) Then '需限定左右，不可直接使用ControlKey,Menu,Shift等
            If GetAsyncKeyState(Keys.LMenu) Then
                If GetAsyncKeyState(Keys.A) Then
                    TextBox1.Text = "Ctrl+Alt+A"
                End If
            End If
        End If


    End Sub
End Class
