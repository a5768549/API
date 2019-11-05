VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   7425
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   13380
   LinkTopic       =   "Form1"
   ScaleHeight     =   7425
   ScaleWidth      =   13380
   StartUpPosition =   3  '系統預設值
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   1440
      Top             =   4920
   End
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
      Height          =   2535
      Left            =   3000
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   1440
      Width           =   7335
   End
   Begin VB.Label Label1 
      Caption         =   "已經過時間"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   18
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2880
      TabIndex        =   1
      Top             =   960
      Width           =   3015
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'最高可計算52天開機時間
Private Declare Function GetTickCount Lib "kernel32" () As Long


Private Sub Timer1_Timer()
         Dim i As Integer
         Dim lngMS As Long
         lngMS = GetTickCount
         Text1.Text = lngMS & " milliseconds" & vbNewLine
         Text1.Text = Text1.Text & lngMS / 1000 & " Seconds" & vbNewLine
         Text1.Text = Text1.Text & lngMS / 1000 / 60 & " Minute" & vbNewLine
         Text1.Text = Text1.Text & lngMS / 1000 / 60 / 60 & " Hour" & vbNewLine
         Text1.Text = Text1.Text & lngMS / 1000 / 60 / 60 / 24 & " Day" & vbNewLine
End Sub
