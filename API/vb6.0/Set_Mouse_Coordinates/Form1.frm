VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   6780
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   13335
   LinkTopic       =   "Form1"
   ScaleHeight     =   6780
   ScaleWidth      =   13335
   StartUpPosition =   3  '系統預設值
   Begin VB.TextBox Text2 
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
      Left            =   4800
      TabIndex        =   2
      Text            =   "0"
      Top             =   2280
      Width           =   3015
   End
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
      Height          =   615
      Left            =   4800
      TabIndex        =   1
      Text            =   "0"
      Top             =   1320
      Width           =   3015
   End
   Begin VB.CommandButton Command1 
      Caption         =   "移動"
      Height          =   855
      Left            =   4920
      TabIndex        =   0
      Top             =   3600
      Width           =   2775
   End
   Begin VB.Label Label3 
      Caption         =   "Y"
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
      Left            =   4080
      TabIndex        =   5
      Top             =   2280
      Width           =   615
   End
   Begin VB.Label Label2 
      Caption         =   "X"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   15.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4080
      TabIndex        =   4
      Top             =   1440
      Width           =   495
   End
   Begin VB.Label Label1 
      Caption         =   "鼠標位置"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   15.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4800
      TabIndex        =   3
      Top             =   720
      Width           =   1695
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

Private Declare Function SetCursorPos Lib "user32" (ByVal X As Long, ByVal Y As Long) As Long

Private Declare Function ClientToScreen Lib "user32" (ByVal hWnd As Long, lpPoint As POINTAPI) As Long
' Get the coordinates (in pixels) of the center of the Command1 button.
' The coordinates are relative to the button's client area.



Private Sub Command1_Click()
    Dim lpPoint As POINTAPI
    lpPoint.X = Text1.Text 'ScaleX(Command1.Width / 2, vbTwips, vbPixels)
    lpPoint.Y = Text2.Text 'ScaleY(Command1.Height / 2, vbTwips, vbPixels)
    ' Convert to screen coordinates.
    ClientToScreen Me.hWnd, lpPoint
    ' Move the mouse cursor to that point.
    SetCursorPos lpPoint.X, lpPoint.Y
End Sub

