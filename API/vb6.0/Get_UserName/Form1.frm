VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5880
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9720
   LinkTopic       =   "Form1"
   ScaleHeight     =   5880
   ScaleWidth      =   9720
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
      Left            =   1200
      TabIndex        =   1
      Top             =   840
      Width           =   7575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "查詢"
      Height          =   975
      Left            =   3000
      TabIndex        =   0
      Top             =   2880
      Width           =   3615
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpBuffer As String, nSize As Long) As Long
    
    
Dim xstr As String
Dim max, rc As Integer

Private Sub Command1_Click()
    xstr = Space$(255)
    max = 255
    rc = GetUserName(xstr, max)
    Text1.Text = (Mid(xstr, 1, max))
End Sub
