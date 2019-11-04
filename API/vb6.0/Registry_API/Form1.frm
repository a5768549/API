VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   8160
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   13050
   LinkTopic       =   "Form1"
   ScaleHeight     =   8160
   ScaleWidth      =   13050
   StartUpPosition =   3  '系統預設值
   Begin VB.CommandButton Command5 
      Caption         =   "創建註冊檔"
      Height          =   975
      Left            =   1200
      TabIndex        =   6
      Top             =   6480
      Width           =   2655
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   15.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3855
      Left            =   5040
      MultiLine       =   -1  'True
      TabIndex        =   5
      Top             =   1920
      Width           =   6255
   End
   Begin VB.CommandButton Command4 
      Caption         =   "讀取註冊檔"
      Height          =   1095
      Left            =   1080
      TabIndex        =   4
      Top             =   3120
      Width           =   2895
   End
   Begin VB.CommandButton Command3 
      Caption         =   "刪除註冊檔"
      Height          =   975
      Left            =   8400
      TabIndex        =   3
      Top             =   6480
      Width           =   2775
   End
   Begin VB.CommandButton Command2 
      Caption         =   "創建註冊檔目錄"
      Height          =   975
      Left            =   4920
      TabIndex        =   2
      Top             =   6480
      Width           =   2655
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   15.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   840
      TabIndex        =   1
      Top             =   1920
      Width           =   3375
   End
   Begin VB.CommandButton Command1 
      Caption         =   "開啟註冊檔"
      Height          =   975
      Left            =   960
      TabIndex        =   0
      Top             =   840
      Width           =   3015
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" _
    (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As _
Long, ByVal samDesired As Long, phkResult As Long) As Long

Private Declare Function RegCreateKeyEx Lib "advapi32.dll" Alias "RegCreateKeyExA" _
(ByVal hKey As Long, ByVal lpSubKey As String, ByVal Reserved As Long, _
ByVal lpClass As Long, ByVal dwOptions As Long, _
    ByVal samDesired As Long, ByVal lpSecurityAttributes As Long, _
    phkResult As Long, lpdwDisposition As Long) As Long

Private Declare Function RegDeleteKey Lib "advapi32.dll" Alias "RegDeleteKeyA" _
(ByVal hKey As Long, ByVal lpSubKey As String) As Long

Private Declare Function RegQueryValueEx Lib "advapi32.dll" Alias _
    "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, _
    ByVal lpReserved As Long, lpType As Long, lpData As Any, _
lpcbData As Long) As Long

Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long

Private Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" _
    (ByVal hKey As Long, ByVal lpValueName As String, _
    ByVal Reserved As Long, ByVal dwType As Long, lpData As Any, _
    ByVal cbData As Long) As Long



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

Const REG_OPENED_EXISTING_KEY = &H2

Function MathProcessor() As Boolean
    Dim hKey As Long, Key As String
    Dim result As String
    Key = "HARDWARE\DESCRIPTION\System\FloatingPointProcessor"
    result = RegOpenKeyEx(HKEY_LOCAL_MACHINE, "HARDWARE\DESCRIPTION\System\FloatingPointProcessor", 0, KEY_READ, hKey)
    RegCloseKey hKey
    If result = 0 Then
        MathProcessor = True
    Else
        MathProcessor = False
    End If
    
End Function

Private Sub Command1_Click()
    Text1.Text = MathProcessor()
    If Text1.Text = "True" Then
        MsgBox ("存取成功")
    Else
        MsgBox ("存取失敗")
    End If
    
End Sub


Function CreateRegistryKey(ByVal hKey As Long, ByVal KeyName As String) As Boolean
    Dim handle As Long, disp As Long
    If RegCreateKeyEx(hKey, KeyName, 0, 0, 0, 0, 0, handle, disp) Then
        Err.Raise 1001, , "Unable to create the Registry key"
    Else
        ' Return True if the key already existed.
        If disp = REG_OPENED_EXISTING_KEY Then CreateRegistryKey = True
        ' Close the key.
        RegCloseKey handle
    End If
End Function

Private Sub Command2_Click()
    CreateRegistryKey HKEY_CURRENT_USER, "Software\YourCompany"
    CreateRegistryKey HKEY_CURRENT_USER, "Software\YourCompany\YourApplication"
End Sub

Private Sub Command3_Click()
    RegDeleteKey HKEY_CURRENT_USER, "Software\YourCompany\YourApplication"
    RegDeleteKey HKEY_CURRENT_USER, "Software\YourCompany"
End Sub

Private Sub Command4_Click()
    Text2.Text = ""
    Dim KeyName As String, handle As Long
    Dim FontHeight As Long, FontFace As String, FontFaceLen As Long
    KeyName = "Software\Microsoft\VBA\Microsoft Visual Basic"
    If RegOpenKeyEx(HKEY_CURRENT_USER, KeyName, 0, KEY_READ, handle) Then
        Text2.Text = "Unable to open the specified Registry key"
    Else
    ' Read the ihFontHeightlH value.
    If RegQueryValueEx(handle, "FontHeight", 0, REG_DWORD, FontHeight, 4) = 0 Then
        Text2.Text = Text2.Text & "Face Height = " & FontHeight & vbNewLine
    End If
    
    ' Read the "FontFace" value.
    FontFaceLen = 128                   ' Prepare the receiving buffer.
    FontFace = Space$(FontFaceLen)
    ' Notice that FontFace is passed using ByVal.
    If RegQueryValueEx(handle, "FontFace", 0, REG_SZ, ByVal FontFace, FontFaceLen) = 0 Then
        ' Trim excess characters, including the trailing Null char.
        FontFace = Left$(FontFace, FontFaceLen - 1)
        Text2.Text = Text2.Text & "Face Name = " & FontFace
    End If
    ' Close the Registry key.
    RegCloseKey handle
End If
End Sub

Private Sub Command5_Click()
    Dim handle As Long, strValue As String
    If RegOpenKeyEx(HKEY_CURRENT_USER, "Software\YourCompany\YourApplication", 0, KEY_WRITE, handle) Then
        MsgBox "Unable to open the key."
    Else
        strValue = FormatDateTime(Now)
        RegSetValueEx handle, "LastLogin", 0, REG_SZ, ByVal strValue, Len(strValue)
        RegCloseKey handle
    End If
End Sub
