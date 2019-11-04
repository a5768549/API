VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   7350
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   11565
   LinkTopic       =   "Form1"
   ScaleHeight     =   7350
   ScaleWidth      =   11565
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
      Left            =   2880
      TabIndex        =   0
      Top             =   2040
      Width           =   4335
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   1320
      Top             =   6120
   End
   Begin VB.Label Label1 
      Caption         =   "滑鼠位置(像素)"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   15.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2760
      TabIndex        =   1
      Top             =   1320
      Width           =   2295
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Type POINTAPI
    X As Long
    Y As Long
End Type

Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long



Private Sub Timer1_Timer()
    ' Display current mouse screen coordinates in pixels using a Label control.
    Dim lpPoint As POINTAPI
    GetCursorPos lpPoint
    lblMouseState = "X = " & lpPoint.X & "   Y = " & lpPoint.Y
    Text1.Text = lblMouseState
End Sub
