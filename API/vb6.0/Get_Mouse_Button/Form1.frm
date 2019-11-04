VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5430
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   12030
   LinkTopic       =   "Form1"
   ScaleHeight     =   5430
   ScaleWidth      =   12030
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
      Height          =   1815
      Left            =   2640
      TabIndex        =   0
      Top             =   1680
      Width           =   6135
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   480
      Top             =   3840
   End
   Begin VB.Label Label1 
      Caption         =   "滑鼠狀態"
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
      Left            =   2640
      TabIndex        =   1
      Top             =   840
      Width           =   2055
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer

' Detect the Ctrl+Alt+A key combination.
Private Sub Timer1_Timer()
    Text1.Text = MouseButton()
End Sub

Function MouseButton() As String
    Select Case True
        Case GetAsyncKeyState(vbKeyLButton) < 0: MouseButton = "左鍵"
        Case GetAsyncKeyState(vbKeyRButton) < 0: MouseButton = "右鍵"
        Case GetAsyncKeyState(vbKeyMButton) < 0: MouseButton = "中鍵"
        Case Else: MouseButton = "無"
    End Select
End Function
