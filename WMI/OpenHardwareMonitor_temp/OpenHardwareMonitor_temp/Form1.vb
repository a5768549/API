Imports OpenHardwareMonitor
Imports OpenHardwareMonitor.Hardware

Public Class Form1

    Private Sub Button1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles Button1.Click
        Dim cp As New Computer()
        cp.Open()
        cp.HDDEnabled = True
        cp.FanControllerEnabled = True
        cp.RAMEnabled = True
        cp.GPUEnabled = True
        cp.MainboardEnabled = True
        cp.CPUEnabled = True

        Dim Info As String = ""
        TextBox3.Text = ""
        For i As Integer = 0 To cp.Hardware.Count() - 1
            Dim hw = cp.Hardware(i)

            Select Case hw.HardwareType
                Case HardwareType.Mainboard
                    TextBox3.AppendText("Motherboard" & vbCrLf)
                    For k = 0 To hw.SubHardware.Count - 1
                        Dim subhardware = hw.SubHardware(k)
                        TextBox3.AppendText(subhardware.Name & vbCrLf)
                        For j = 0 To subhardware.Sensors.Count - 1
                            Dim sensor = subhardware.Sensors(j)
                            TextBox3.AppendText(sensor.SensorType & " - " & sensor.Name & " - " & sensor.Value & vbCrLf)
                        Next
                    Next
                Case HardwareType.CPU
                    TextBox3.AppendText("CPU" & vbCrLf)
                    For j = 0 To hw.Sensors.Count - 1
                        Dim sensor = hw.Sensors(j)
                        TextBox3.AppendText(sensor.SensorType & " - " & sensor.Name & " - " & sensor.Value & vbCrLf)
                    Next
                Case HardwareType.RAM
                    TextBox3.AppendText("RAM" & vbCrLf)
                    For j = 0 To hw.Sensors.Count - 1
                        Dim sensor = hw.Sensors(j)
                        TextBox3.AppendText(sensor.SensorType & " - " & sensor.Name & " - " & sensor.Value & vbCrLf)
                    Next
            End Select
        Next
    End Sub
End Class