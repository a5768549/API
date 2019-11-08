Imports System.Runtime.InteropServices

Public Class form1
    <StructLayout(LayoutKind.Sequential)>
    Private Structure SYSTEMTIME
        <MarshalAs(UnmanagedType.U2)> Public Year As Short
        <MarshalAs(UnmanagedType.U2)> Public Month As Short
        <MarshalAs(UnmanagedType.U2)> Public DayOfWeek As Short
        <MarshalAs(UnmanagedType.U2)> Public Day As Short
        <MarshalAs(UnmanagedType.U2)> Public Hour As Short
        <MarshalAs(UnmanagedType.U2)> Public Minute As Short
        <MarshalAs(UnmanagedType.U2)> Public Second As Short
        <MarshalAs(UnmanagedType.U2)> Public Milliseconds As Short
    End Structure

    <DllImport("kernel32.dll")>
    Private Shared Sub GetLocalTime(ByRef time As SYSTEMTIME)
    End Sub


    Private Sub Timer1_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick
        Dim Year, Month, Day, Hour, Minute, Second As String
        Dim Time As String
        Dim st As SYSTEMTIME = New SYSTEMTIME()
        GetLocalTime(st)
        Year = st.Year
        Month = st.Month
        Day = st.Day
        Hour = st.Hour
        Minute = st.Minute
        Second = st.Second
        Time = Year & "-" & Month & "-" & Day & " " & Hour & ":" & Minute & ":" & Second
        TextBox1.Text = Time
    End Sub
End Class
