Option Explicit On
Public Class Form1
    'GetDriveType return values
    Const DRIVE_REMOVABLE = 2
    Const DRIVE_FIXED = 3
    Const DRIVE_REMOTE = 4
    Const DRIVE_CDROM = 5
    Const DRIVE_RAMDISK = 6

    Private Declare Function GetDriveType Lib "kernel32" Alias "GetDriveTypeA" (ByVal nDrive As String) As Integer

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim i As Integer
        Dim lngDrive As Long
        Dim strD As String
        TextBox1.Text = ""
        For i = 0 To 25  'All possible drives A to Z
            strD = Chr(i + 65) & ":\"
            lngDrive = GetDriveType(strD)
            Select Case lngDrive
                Case DRIVE_REMOVABLE
                    TextBox1.Text = TextBox1.Text & "Drive " & strD & " is removable." & vbNewLine
                Case DRIVE_FIXED
                    TextBox1.Text = TextBox1.Text & "Drive " & strD & " is fixed." & vbNewLine
                Case DRIVE_REMOTE
                    TextBox1.Text = TextBox1.Text & "Drive " & strD & " is remote." & vbNewLine
                Case DRIVE_CDROM
                    TextBox1.Text = TextBox1.Text & "Drive " & strD & " is CD-ROM." & vbNewLine
                Case DRIVE_RAMDISK
                    TextBox1.Text = TextBox1.Text & "Drive " & strD & " is RAM disk." & vbNewLine
                Case Else
            End Select
        Next i
    End Sub
End Class
