VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   6945
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   13185
   LinkTopic       =   "Form1"
   ScaleHeight     =   6945
   ScaleWidth      =   13185
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
      Height          =   2055
      Left            =   3480
      TabIndex        =   0
      Top             =   2160
      Width           =   6375
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   1920
      Top             =   5160
   End
   Begin VB.Label Label1 
      Caption         =   "視窗位置(像素)"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   18
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3480
      TabIndex        =   1
      Top             =   1440
      Width           =   3015
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function ScreenToClient Lib "user32" (ByVal hWnd As Long, lpPoint As POINTAPI) As Long

Private Type POINTAPI
    X As Long
    Y As Long
End Type
' Display Windows screen coordinates relative to current form.



Private Sub Timer1_Timer()
    Dim lpPoint As POINTAPI
        ScreenToClient Me.hWnd, lpPoint
        lblMouseState = "X = " & lpPoint.X & "   Y = " & lpPoint.Y
        Text1.Text = lblMouseState
End Sub
