Option Strict On
Option Explicit On
Option Infer On
Imports Microsoft.VisualBasic
Imports System.Runtime.InteropServices
Imports System.Runtime.ConstrainedExecution
Imports System.ComponentModel
Imports System.Text
Imports System.Security

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

    Private Structure PredefinedKeys
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

    End Class

    Sub Main()
        ' HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows NT\CurrentVersion
        Dim keyHandle As SafeKeyHandle = Nothing
        Try
            ' Get a handle to the key: HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows NT\CurrentVersion
            Dim result = NativeMethods.RegOpenKeyEx(PredefinedKeys.HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows NT\CurrentVersion", 0, RegSam.KEY_READ, keyHandle)
            If result <> 0 Then Throw New Win32Exception
            Dim sbProductName As New StringBuilder(128)
            Dim dataSize As Integer = sbProductName.Capacity
            Dim dataType As RegType = RegType.REG_NONE ' This is returned to us.
            result = NativeMethods.RegQueryValueEx(keyHandle, "ProductName", 0, dataType, sbProductName, dataSize)
            If result <> 0 Then Throw New Win32Exception
            Console.WriteLine("Got a : " & dataType.ToString)
            Console.WriteLine(sbProductName.ToString)
            Dim installDate As Integer ' dword, could use uInteger, if it made sense to.
            dataSize = Marshal.SizeOf(installDate)
            dataType = RegType.REG_NONE
            result = NativeMethods.RegQueryValueEx(keyHandle, "InstallDate", 0, dataType, installDate, dataSize) ' 4 bytes in an integer.
            If result <> 0 Then Throw New Win32Exception
            Console.WriteLine("Got a : " & dataType.ToString)
            Console.WriteLine(installDate) ' Can't be bothered to decode it - if it is 0, make sure you aren't running an x86 build on an x64 platform - use AnyCPU.
        Finally
            If keyHandle IsNot Nothing And keyHandle.IsInvalid = False And keyHandle.IsClosed = False Then keyHandle.Dispose()
        End Try

    End Sub

End Module
