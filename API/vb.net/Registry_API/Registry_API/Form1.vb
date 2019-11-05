Imports Microsoft.VisualBasic
Imports System.Runtime.InteropServices
Imports System.Runtime.ConstrainedExecution
Imports System.ComponentModel
Imports System.Text
Imports System.Security
Imports Registry_API
Imports Microsoft.Win32

Public Class Form1
    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        TextBox2.Text = "Face Height = " & readReg_number()
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If openReg() = True Then
            TextBox1.Text = "True"
            MsgBox("存取成功")
        Else
            TextBox1.Text = "False"
            MsgBox("存取失敗")
        End If
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        CreateRegistryKey(PredefinedKeys.HKEY_CURRENT_USER, "Software\YourCompany")
        CreateRegistryKey(PredefinedKeys.HKEY_CURRENT_USER, "Software\YourCompany\YourApplication")
    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        TextBox3.Text = "Face Name = " & readReg_string()
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        deleteReg()
    End Sub

    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click
        SetValue("LastLogin", FormatDateTime(Now))
    End Sub
End Class


Module Module1

    ' wraps the native handle to the registry key object - HKEY
    Private Class SafeKeyHandle
        Inherits SafeHandle

        <ReliabilityContract(Consistency.WillNotCorruptState, Cer.MayFail)>
        Sub New()
            MyBase.New(IntPtr.Zero, True)
        End Sub

        Public Overrides ReadOnly Property IsInvalid As Boolean
            <ReliabilityContract(Consistency.WillNotCorruptState, Cer.Success)>
            Get
                Return Me.handle = IntPtr.Zero
            End Get
        End Property

        <ReliabilityContract(Consistency.WillNotCorruptState, Cer.Success)>
        Protected Overrides Function ReleaseHandle() As Boolean
            Return NativeMethods.RegCloseKey(Me.handle) = 0
        End Function

    End Class

    Private Enum RegSam
        KEY_QUERY_VALUE = 1
        KEY_SET_VALUE = 2
        KEY_CREATE_SUB_KEY = 4
        KEY_ENUMERATE_SUB_KEYS = 8
        KEY_NOTIFY = &H10
        KEY_CREATE_LINK = &H20
        KEY_WOW64_64KEY = &H100
        KEY_WOW64_32KEY = &H200
        KEY_ALL_ACCESS = &HF003F
        KEY_WRITE = &H20006
        KEY_EXECUTE = &H20019
        KEY_READ = &H20019
        REG_OPENED_EXISTING_KEY = &H2
    End Enum

    Private Enum RegType
        REG_NONE = 0
        REG_SZ = 1
        REG_EXPAND_SZ = 2
        REG_BINARY = 3
        REG_DWORD = 4
        REG_DWORD_LITTLE_ENDIAN = 4
        REG_DWORD_BIG_ENDIAN = 5
        REG_LINK = 6
        REG_MULTI_SZ = 7
        REG_RESOURCE_LIST = 8
        REG_FULL_RESOURCE_DESCRIPTOR = 9
        REG_RESOURCE_REQUIREMENTS_LIST = 10
        REG_QWORD = 11
        REG_QWORD_LITTLE_ENDIAN = 11
    End Enum

    Public Structure PredefinedKeys
        Public Shared HKEY_CLASSES_ROOT As New IntPtr(&H80000000)
        Public Shared HKEY_CURRENT_USER As New IntPtr(&H80000001)
        Public Shared HKEY_LOCAL_MACHINE As New IntPtr(&H80000002)
        Public Shared HKEY_USERS As New IntPtr(&H80000003)
        Public Shared HKEY_PERFORMANCE_DATA As New IntPtr(&H80000004)
        Public Shared HKEY_PERFORMANCE_TEXT As New IntPtr(&H80000050)
        Public Shared HKEY_PERFORMANCE_NLSTEXT As New IntPtr(&H80000060)
        Public Shared HKEY_CURRENT_CONFIG As New IntPtr(&H80000005)
        Public Shared HKEY_DYN_DATA As New IntPtr(&H80000006)
        Public Shared HKEY_CURRENT_USER_LOCAL_SETTINGS As New IntPtr(&H80000007)
    End Structure

    <SuppressUnmanagedCodeSecurity()>
    Private Class NativeMethods
        <ReliabilityContract(Consistency.WillNotCorruptState, Cer.MayFail)>
        <DllImport("Advapi32", SetLastError:=True, CharSet:=CharSet.Unicode)>
        Public Shared Function RegOpenKeyEx(
      <[In]()> ByVal hKey As IntPtr,
      <[In]()> ByVal subKey As String,
      ByVal options As Integer,
      <[In]()> ByVal samDesired As RegSam,
      <Out()> ByRef phkResult As SafeKeyHandle) As Integer
        End Function

        <ReliabilityContract(Consistency.WillNotCorruptState, Cer.Success)>
        <DllImport("Advapi32", SetLastError:=True)>
        Public Shared Function RegCloseKey(ByVal hKey As IntPtr) As Integer
        End Function

        ' Use different overloads for different types. - Integer:
        <DllImport("Advapi32", SetLastError:=True, CharSet:=CharSet.Unicode)>
        Public Shared Function RegQueryValueEx(
      <[In]()> ByVal hKey As SafeKeyHandle,
      <[In]()> ByVal lpValueName As String,
      ByVal reserved As UInteger,
      <Out()> ByRef type As RegType,
      <Out()> ByRef data As Integer,
      <[In](), Out()> ByRef dataSize As Integer) As Integer
        End Function

        ' Use different overloads for different types. - null terminated String:
        <DllImport("Advapi32", SetLastError:=True, CharSet:=CharSet.Unicode)>
        Public Shared Function RegQueryValueEx(
      <[In]()> ByVal hKey As SafeKeyHandle,
      <[In]()> ByVal lpValueName As String,
      ByVal reserved As Integer,
      <Out()> ByRef type As RegType,
      <Out()> ByVal data As StringBuilder,
      <[In](), Out()> ByRef dataSize As Integer) As Integer
        End Function

        <DllImport("advapi32", SetLastError:=False)>
        Public Shared Function RegCreateKeyEx(
        ByVal hKey As Integer,
        ByVal lpSubKey As String,
        ByVal Reserved As Integer,
        ByVal lpClass As Integer,
        ByVal dwOptions As Integer,
        ByVal samDesired As RegSam,
        ByRef lpSecurityAttributes As Integer,
        ByRef phkResult As Integer,
        ByRef lpdwDisposition As Integer) As Integer
        End Function

        <DllImport("advapi32", SetLastError:=False)>
        Public Shared Function RegDeleteKey(
        ByVal hKey As Integer,
        ByVal lpSubKey As String) As Integer
        End Function

        <DllImport("advapi32.dll", SetLastError:=True)>
        Friend Shared Function RegSetValueEx(ByVal hKey As SafeKeyHandle,
            ByVal lpValueName As String,
            ByVal lpReserved As Integer,
            ByVal lpType As RegType,
            ByVal lpData As String,
            ByVal lpcbData As Integer) As Integer
        End Function
    End Class

    Public Sub SetValue(ByVal Name As String, ByVal Value As Object)
        Dim gch As SafeKeyHandle
        Dim ptr As IntPtr
        Dim ret, Size As Integer
        'Dim temp() As Byte                                     //Binary
        'temp = DirectCast(Value, System.Byte())                
        'Size = temp.Length
        'gch = GCHandle.Alloc(temp, GCHandleType.Pinned)
        'ptr = Marshal.UnsafeAddrOfPinnedArrayElement(temp, 0)

        'Dim temp As Integer = CInt(Value)                      //DWORD
        'Size = 4
        'ptr = Marshal.AllocHGlobal(Size)
        'Marshal.WriteInt32(ptr, 0, temp)

        'Dim temp As String = CStr(Value)                       //ExpandString
        'Size = (temp.Length + 1) * Marshal.SystemDefaultCharSize
        'ptr = Marshal.StringToHGlobalAuto(temp)

        'Dim temp, lines() As String                            //MultiString
        'Dim index As Integer
        'lines = DirectCast(Value, System.String())

        ' Calculate the total size, including the terminating null
        'Size = 0
        'For Each temp In lines
        'Size += (temp.Length + 1) * Marshal.SystemDefaultCharSize
        'Next
        'Size += Marshal.SystemDefaultCharSize
        'ptr = Marshal.AllocHGlobal(Size)

        'Index = 0
        'For Each temp In lines
        'Dim tempPtr As IntPtr
        'Dim tempArray() As Char

        'tempArray = temp.ToCharArray
        'tempPtr = New IntPtr(ptr.ToInt64 + index)
        'Marshal.Copy(tempArray, 0, tempPtr, tempArray.Length)
        'Index += (tempArray.Length + 1) * Marshal.SystemDefaultCharSize
        'Next

        'Dim temp As Long = CLng(Value)                             //QWORD
        'Size = 8
        'ptr = Marshal.AllocHGlobal(Size)
        'Marshal.WriteInt64(ptr, 0, temp)

        'Dim temp As String = CStr(Value)                           //String
        'Size = (temp.Length + 1) * Marshal.SystemDefaultCharSize
        'ptr = Marshal.StringToHGlobalAuto(temp)

        Dim temp As String = CStr(Value)
        Size = (temp.Length + 1) * Marshal.SystemDefaultCharSize
        'ptr = Marshal.StringToHGlobalAuto(temp)

        If NativeMethods.RegOpenKeyEx(PredefinedKeys.HKEY_CURRENT_USER, "Software\YourCompany\YourApplication", 0, RegSam.KEY_WRITE, gch) = 0 Then
            ret = NativeMethods.RegSetValueEx(gch, Name, 0, RegType.REG_SZ, temp, Size)
        End If
        gch.Dispose()
    End Sub
    Sub CreateRegistryKey(ByVal hKey As Integer, ByVal KeyName As String)
        Dim keyHandle As Long
        If NativeMethods.RegCreateKeyEx(hKey, KeyName, 0, 0, 0, 0, 0, CInt(keyHandle), 0) Then
            Err.Raise(1001, , "Unable to create the Registry key")
        End If
        NativeMethods.RegCloseKey(CType(keyHandle, IntPtr))
    End Sub
    Function openReg() As Boolean
        Dim keyHandle As SafeKeyHandle = Nothing
        Dim result = NativeMethods.RegOpenKeyEx(PredefinedKeys.HKEY_LOCAL_MACHINE, "HARDWARE\DESCRIPTION\System\FloatingPointProcessor", 0, RegSam.KEY_READ, keyHandle)
        If result = 0 Then
            openReg = True
        Else
            openReg = False
        End If
        keyHandle.Dispose()
    End Function

    Function readReg_number() As String
        Dim keyHandle As SafeKeyHandle = Nothing
        Dim result = NativeMethods.RegOpenKeyEx(PredefinedKeys.HKEY_CURRENT_USER, "Software\Microsoft\VBA\Microsoft Visual Basic", 0, RegSam.KEY_READ, keyHandle)
        Dim number As Integer
        Dim dataSize As Integer = Marshal.SizeOf(number)
        Dim dataType As RegType = RegType.REG_NONE ' This is returned to us.
        result = NativeMethods.RegQueryValueEx(keyHandle, "FontHeight", 0, dataType, number, dataSize)
        readReg_number = number.ToString
        keyHandle.Dispose()
    End Function

    Function readReg_string() As String
        Dim keyHandle As SafeKeyHandle = Nothing
        Dim result = NativeMethods.RegOpenKeyEx(PredefinedKeys.HKEY_CURRENT_USER, "Software\Microsoft\VBA\Microsoft Visual Basic", 0, RegSam.KEY_READ, keyHandle)
        Dim string_data As New StringBuilder(128)
        Dim dataSize As Integer = string_data.Capacity
        Dim dataType As RegType = RegType.REG_NONE ' This is returned to us.
        result = NativeMethods.RegQueryValueEx(keyHandle, "FontFace", 0, dataType, string_data, dataSize)
        readReg_string = string_data.ToString
        keyHandle.Dispose()
    End Function

    Sub deleteReg()
        NativeMethods.RegDeleteKey(PredefinedKeys.HKEY_CURRENT_USER, "Software\YourCompany\YourApplication")
        NativeMethods.RegDeleteKey(PredefinedKeys.HKEY_CURRENT_USER, "Software\YourCompany")
    End Sub

End Module