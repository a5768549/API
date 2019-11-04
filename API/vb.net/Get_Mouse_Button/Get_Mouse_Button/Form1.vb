Public Class Form1
    Private Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As IntPtr) As Short
    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub

    Function MouseButton() As String
        Select Case True
            Case GetAsyncKeyState(Keys.LButton) < 0 : MouseButton = "左鍵"
            Case GetAsyncKeyState(Keys.RButton) < 0 : MouseButton = "右鍵"
            Case GetAsyncKeyState(Keys.MButton) < 0 : MouseButton = "中鍵"
            Case Else : MouseButton = "無"
        End Select
    End Function

    Private Sub Timer1_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick
        TextBox1.Text = MouseButton()
    End Sub
End Class
