VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   6765
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9855
   LinkTopic       =   "Form1"
   ScaleHeight     =   6765
   ScaleWidth      =   9855
   StartUpPosition =   3  '系統預設值
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
      Height          =   1455
      Left            =   2040
      TabIndex        =   1
      Top             =   1800
      Width           =   5655
   End
   Begin VB.CommandButton Command1 
      Caption         =   "查詢"
      Height          =   1335
      Left            =   3120
      TabIndex        =   0
      Top             =   4080
      Width           =   3255
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function GetComputerName Lib "kernel32" Alias _
    "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Dim buffer As String * 512, length As Long

Private Sub Command1_Click()
    length = Len(buffer)
    If GetComputerName(buffer, length) Then
        ' Returns nonzero if successful, and modifies the length argument
    Text1.Text = "Computer Name = " & Left$(buffer, length)
    End If
End Sub
