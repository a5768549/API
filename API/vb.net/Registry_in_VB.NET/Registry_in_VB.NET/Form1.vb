Public Class Form1
    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        My.Computer.Registry.CurrentUser.CreateSubKey("MyTestKey")
        My.Computer.Registry.SetValue("HKEY_CURRENT_USER\MyTestKey", "MyTestKeyValue", "This is a test value.")
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Dim readValue = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\MyTestKey", "MyTestKeyValue", Nothing)
        TextBox1.Text = "The value is:" & vbNewLine & readValue
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        My.Computer.Registry.CurrentUser.DeleteSubKey("MyTestKey")
    End Sub
End Class
