VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5775
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   11775
   LinkTopic       =   "Form1"
   ScaleHeight     =   5775
   ScaleWidth      =   11775
   StartUpPosition =   3  '系統預設值
   Begin VB.CommandButton Command3 
      Caption         =   "String"
      Height          =   735
      Left            =   6960
      TabIndex        =   5
      Top             =   4440
      Width           =   3975
   End
   Begin VB.TextBox Text3 
      Height          =   735
      Left            =   6360
      TabIndex        =   4
      Top             =   3360
      Width           =   5055
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Integer"
      Height          =   735
      Left            =   6960
      TabIndex        =   3
      Top             =   1920
      Width           =   3975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Push"
      Height          =   735
      Left            =   1080
      TabIndex        =   2
      Top             =   1920
      Width           =   3975
   End
   Begin VB.TextBox Text2 
      Height          =   615
      Left            =   6360
      TabIndex        =   1
      Top             =   840
      Width           =   5055
   End
   Begin VB.TextBox Text1 
      Height          =   615
      Left            =   840
      TabIndex        =   0
      Top             =   840
      Width           =   4695
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function WritePrivateProfileString Lib "kernel32" _
Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, _
ByVal lpKeyName As Any, ByVal lpString As Any, _
ByVal lpFileName As String) As Long

Private Declare Function GetPrivateProfileInt Lib "kernel32" Alias _
"GetPrivateProfileIntA" (ByVal lpApplicationName As String, _
ByVal lpKeyName As String, ByVal nDefault As Long, _
ByVal lpFileName As String) As Long

Private Declare Function GetPrivateProfileString _
 Lib "kernel32" Alias "GetPrivateProfileStringA" _
(ByVal lpSectionName As String, _
 ByVal lpKeyName As Any, _
 ByVal lpDefault As String, _
 ByVal lpbuffurnedString As String, _
 ByVal nBuffSize As Long, _
 ByVal lpFileName As String) As Long

Private Sub Command1_Click()
    WritePrivateProfileString "test1", "value1", CStr(Text1.Text), App.Path & "\config.ini" 'text1
End Sub

Private Sub Command2_Click()
    Text2.Text = GetPrivateProfileInt("test1", "value1", 0, App.Path & "\config.ini ")
End Sub

Private Sub Command3_Click()
    Dim buff As String * 50
    Dim ret As Long
    ret = GetPrivateProfileString("test1", "value1", "0", buff, 50, App.Path & "\config.ini")
    Text3.Text = buff
End Sub
