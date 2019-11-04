VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   6585
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   11940
   LinkTopic       =   "Form1"
   ScaleHeight     =   6585
   ScaleWidth      =   11940
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
      Height          =   1215
      Left            =   1560
      TabIndex        =   1
      Top             =   1080
      Width           =   8775
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   1215
      Left            =   3360
      TabIndex        =   0
      Top             =   2640
      Width           =   4455
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function GetTempPath Lib "kernel32" Alias "GetTempPathA" _
    (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long
Dim buffer As String, length As Integer


Private Sub Command1_Click()
    buffer = Space$(512)
    length = GetTempPath(Len(buffer), buffer)
    Text1.Text = "Temporary Directory =" & Left$(buffer, length)
End Sub

