Imports Microsoft.VisualBasic
Imports System.Runtime.InteropServices
Imports System.Runtime.ConstrainedExecution
Imports System.ComponentModel
Imports System.Text
Imports System.Security
Imports Microsoft.Win32
Imports System.Collections.Specialized

Public Class Form1
    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim ans As String() = GetSubKeyNames(PredefinedKeys.HKEY_CURRENT_USER, "SoftWare")
        For i = LBound(ans) To UBound(ans)
            ListBox1.Items.Add(ans(i))
        Next

    End Sub
End Class


Module Module1

    ' wraps the native handle to the registry key object - HKEY
    Public Class SafeKeyHandle
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

    Public Enum RegSam
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
        MAX_REG_KEYNAME_SIZE = 255
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


        <ReliabilityContract(Consistency.WillNotCorruptState, Cer.Success)>
        <DllImport("Advapi32", SetLastError:=True)>
        Public Shared Function RegEnumKeyEx(
           ByVal hKey As SafeKeyHandle,
           ByVal dwIndex As Integer,
           ByVal lpName As StringBuilder,
           ByRef lpcName As Integer,
           ByVal lpReserved As IntPtr,
           ByVal lpClass As IntPtr,
           ByVal lpcClass As IntPtr,
           ByVal lpftLastWriteTime As IntPtr) As Integer
        End Function
    End Class

    Public Function GetSubKeyNames(ByVal H As IntPtr, ByVal KeyName As String) As String()
        Dim i, ret, NameSize As Integer
        Dim sc As New StringCollection
        Dim sb As New StringBuilder(RegType.MAX_REG_KEYNAME_SIZE + 1)
        Dim ans(-1) As String
        ' quick sanity check
        Dim hkey As SafeHandle
        Dim result = NativeMethods.RegOpenKeyEx(H, KeyName, 0, RegSam.KEY_READ, hkey)
        If hKey.Equals(IntPtr.Zero) Then
            Throw New ApplicationException("Cannot access a closed registry key")
        End If

        Do
            NameSize = RegType.MAX_REG_KEYNAME_SIZE + 1
            ret = NativeMethods.RegEnumKeyEx(hKey, i, sb, NameSize, IntPtr.Zero, IntPtr.Zero, IntPtr.Zero, IntPtr.Zero)
            If ret <> 0 Then
                Exit Do
            End If
            sc.Add(sb.ToString)
            i += 1
        Loop

        If sc.Count > 0 Then
            ReDim ans(sc.Count - 1)
            sc.CopyTo(ans, 0)
        End If
        Return ans
    End Function

End Module