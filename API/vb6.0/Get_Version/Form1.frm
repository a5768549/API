VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Get_Version"
   ClientHeight    =   6840
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   10140
   LinkTopic       =   "Form1"
   ScaleHeight     =   6840
   ScaleWidth      =   10140
   StartUpPosition =   3  '系統預設值
   Begin VB.TextBox Text3 
      Height          =   855
      Left            =   1080
      TabIndex        =   6
      Top             =   4680
      Width           =   3135
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   855
      Left            =   4320
      TabIndex        =   5
      Top             =   2040
      Width           =   1695
   End
   Begin VB.TextBox Text2 
      Height          =   735
      Left            =   4320
      TabIndex        =   3
      Top             =   1200
      Width           =   3135
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   855
      Left            =   960
      TabIndex        =   1
      Top             =   2040
      Width           =   1695
   End
   Begin VB.TextBox Text1 
      Height          =   735
      Left            =   960
      TabIndex        =   0
      Top             =   1200
      Width           =   2415
   End
   Begin VB.Label Label3 
      Caption         =   "原始碼"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   24
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1080
      TabIndex        =   7
      Top             =   3960
      Width           =   1455
   End
   Begin VB.Label Label2 
      Caption         =   "判斷NT或95 98"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   24
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4320
      TabIndex        =   4
      Top             =   600
      Width           =   3135
   End
   Begin VB.Label Label1 
      Caption         =   "版本號"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   24
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   960
      TabIndex        =   2
      Top             =   600
      Width           =   1455
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function GetVersion Lib "kernel32" () As Long



Private Sub Command1_Click()
    lngVer = GetVersion()
    lngWinVer = lngVer And &HFFFF&
    ' 取得 Windows 版本
    GetWinVer = Format((lngWinVer Mod 256) + ((lngWinVer \ 256) / 100), "Fixed")
    Text1.Text = GetWinVer
End Sub

Private Sub Command2_Click()
    If GetVersion() And &H80000000 Then
        Text2.Text = "Running under Windows 95/98"
    Else
        Text2.Text = "Running under Windows NT"
    End If
End Sub

Private Sub Form_Load()
    Text3.Text = GetVersion()
End Sub
