Public Class Form1
    Structure POINTAPI ' This holds the logical cursor information
        Dim x As UInteger
        Dim y As UInteger
    End Structure
    'Private Declare Function SetCursorPos Lib "user32" (ByVal x As Long, ByVal y As Long) As Long
    Private Declare Function GetCursorPos Lib "user32" (ByRef point As POINTAPI) As Long

    Private Sub Timer1_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer1.Tick
        Dim rect As POINTAPI
        ' Get the current mouse cursor coordinates:
        GetCursorPos(rect)
        ' Print out current position on the form:
        TextBox1.Text = "x:" & rect.x
        TextBox2.Text = "y:" & rect.y
    End Sub


End Class
