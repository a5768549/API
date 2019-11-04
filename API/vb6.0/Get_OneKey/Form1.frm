VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5430
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9570
   LinkTopic       =   "Form1"
   ScaleHeight     =   5430
   ScaleWidth      =   9570
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
      Height          =   1095
      Left            =   1800
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   1920
      Width           =   5535
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   1440
      Top             =   3960
   End
   Begin VB.Label Label1 
      Caption         =   "判斷單鍵"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   15.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1680
      TabIndex        =   1
      Top             =   840
      Width           =   1935
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
    
Private Const VK_LSHIFT = &HA0
Private Const VK_LCONTROL = &HA2
Private Const VK_LMENU = &HA4
Private Const VK_RSHIFT = &HA1
Private Const VK_RCONTROL = &HA3
Private Const VK_RMENU = &HA5
'詳細用API檢視員查 關鍵字:VK_()常數

Private Sub Timer1_Timer()
    Dim msg As String
    If GetAsyncKeyState(VK_LSHIFT) And &H8000 Then msg = msg & "LSHIFT "
    If GetAsyncKeyState(VK_LCONTROL) And &H8000 Then msg = msg & "LCTRL "
    If GetAsyncKeyState(VK_LMENU) And &H8000 Then msg = msg & "LALT "
    
    If GetAsyncKeyState(VK_RSHIFT) And &H8000 Then msg = msg & "RSHIFT "
    If GetAsyncKeyState(VK_RCONTROL) And &H8000 Then msg = msg & "RCTRL "
    If GetAsyncKeyState(VK_RMENU) And &H8000 Then msg = msg & "RALT "
    Text1.Text = msg
End Sub
