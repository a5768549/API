VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   6285
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   13200
   LinkTopic       =   "Form1"
   ScaleHeight     =   6285
   ScaleWidth      =   13200
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
      Left            =   3720
      TabIndex        =   0
      Top             =   1800
      Width           =   4815
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   480
      Top             =   5040
   End
   Begin VB.Label Label1 
      Caption         =   "三鍵按下"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   18
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3720
      TabIndex        =   1
      Top             =   840
      Width           =   3015
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
    If GetAsyncKeyState(vbKeyA) And &H8000 Then
        If GetAsyncKeyState(vbKeyControl) And &H8000 Then
            If GetAsyncKeyState(vbKeyMenu) And &H8000 Then
                ' Process the Ctrl+Alt+A hot key here.
                Text1.Text = "Ctrl+Alt+A"
            End If
        End If
    End If
End Sub
