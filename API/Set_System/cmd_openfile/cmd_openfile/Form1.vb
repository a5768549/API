Public Class Form1


    Private Function RunCmd(ByVal command As String) As String
        '實例一個Process類，啟動一個獨立進程
        Dim p As Process = New Process()

        'Process類有一個StartInfo屬性，這個是ProcessStartInfo類，包括了一些屬性和方法，下面我們用到了他的幾個屬性：

        p.StartInfo.FileName = "cmd.exe"           '設定程序名
        p.StartInfo.Arguments = "/c " + command    '設定程式執行參數
        p.StartInfo.UseShellExecute = False        '關閉Shell的使用
        p.StartInfo.RedirectStandardInput = True   '重定向標準輸入
        p.StartInfo.RedirectStandardOutput = True  '重定向標準輸出
        p.StartInfo.RedirectStandardError = True   '重定向錯誤輸出
        p.StartInfo.CreateNoWindow = True          '設置不顯示窗口

        p.Start() '啟動

        'p.StandardInput.WriteLine(command);       '也可以用這種方式輸入要執行的命令
        'p.StandardInput.WriteLine("exit");        '不過要記得加上Exit要不然下一行程式執行的時候會當機

        Return p.StandardOutput.ReadToEnd()        '從輸出流取得命令執行結果

    End Function

    Function Get_file(Path As String, file_Name As String)
        Console.WriteLine(RunCmd("cd " & Path & " && type " & file_Name))
    End Function

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Get_file(TextBox1.Text, TextBox2.Text)
    End Sub
End Class
