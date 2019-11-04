VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   6090
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   11460
   LinkTopic       =   "Form1"
   ScaleHeight     =   6090
   ScaleWidth      =   11460
   StartUpPosition =   3  '系統預設值
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   18
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   2640
      TabIndex        =   1
      Top             =   840
      Width           =   6495
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   1575
      Left            =   3720
      TabIndex        =   0
      Top             =   2640
      Width           =   3975
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function GetSystemDirectory Lib "kernel32" Alias _
    "GetSystemDirectoryA" (ByVal lpBuffer As String, _
    ByVal nSize As Long) As Long
Dim buffer As String, length As Integer

Private Sub Command1_Click()
    buffer = Space$(512)
    length = GetSystemDirectory(buffer, Len(buffer))
    Text1.Text = "System Directory =" & Left$(buffer, length)
End Sub

