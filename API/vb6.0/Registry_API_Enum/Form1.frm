VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   7845
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   13650
   LinkTopic       =   "Form1"
   ScaleHeight     =   7845
   ScaleWidth      =   13650
   StartUpPosition =   3  '系統預設值
   Begin VB.ListBox List1 
      Height          =   6720
      Left            =   2400
      TabIndex        =   0
      Top             =   480
      Width           =   8295
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function RegEnumKey Lib "advapi32.dll" _
    Alias "RegEnumKeyA" (ByVal hKey As Long, ByVal dwIndex As Long, _
    ByVal lpName As String, ByVal cbName As Long) As Long
    
Private Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" _
    (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As _
Long, ByVal samDesired As Long, phkResult As Long) As Long

Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
    
Private Const HKEY_CURRENT_USER = &H80000001
Private Const HKEY_LOCAL_MACHINE = &H80000002
Private Const HKEY_USERS = &H80000003
Private Const HKEY_CLASSES_ROOT = &H80000000
Private Const KEY_WRITE = &H6
Private Const KEY_QUERY_VALUE = &H1
Private Const KEY_SET_VALUE = &H2
Private Const KEY_CREATE_SUB_KEY = &H4
Private Const KEY_ENUMERATE_SUB_KEYS = &H8
Private Const KEY_ALL_ACCESS = &H3F
Private Const KEY_READ = &H19
    
Private Sub Form_Load()
    Dim keys() As String, i As Long
    keys() = EnumRegistryKeys(HKEY_CURRENT_USER, "SoftWare")
    List1.Clear
    For i = LBound(keys) To UBound(keys)
        List1.AddItem keys(i)
    Next
End Sub


Function EnumRegistryKeys(ByVal hKey As Long, ByVal KeyName As String) As String()
    Dim handle As Long, index As Long, length As Long
    ReDim result(0 To 100) As String
    ' Open the key, exit if not found.
    If Len(KeyName) Then
        If RegOpenKeyEx(hKey, KeyName, 0, KEY_READ, handle) Then
            Exit Function
        End If
        ' Subsequent functions use hKey.
        hKey = handle
    End If
    For index = 0 To 999999
        ' Make room in the array.
        If index > UBound(result) Then
            ReDim Preserve result(index + 99) As String
        End If
        length = 260                   ' Max length for a key name.
        result(index) = Space$(length)
        If RegEnumKey(hKey, index, result(index), length) Then Exit For
        ' Trim excess characters.
        result(index) = Left$(result(index), InStr(result(index), vbNullChar) - 1)
    Next
    ' Close the key if it was actually opened.
    If handle Then RegCloseKey handle
    ' Trim unused items in the array, and return the results to the caller.
    ReDim Preserve result(index - 1) As String
    EnumRegistryKeys = result()
End Function

