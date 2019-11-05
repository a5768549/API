VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   6495
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   11835
   LinkTopic       =   "Form1"
   ScaleHeight     =   6495
   ScaleWidth      =   11835
   StartUpPosition =   3  '系統預設值
   Begin VB.CommandButton Command1 
      Caption         =   "播放"
      Height          =   1455
      Left            =   3360
      TabIndex        =   0
      Top             =   2160
      Width           =   4575
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long

Private Const SND_ASYNC = &H1         '  play asynchronously
Private Const SND_LOOP = &H8         '  loop the sound until next sndPlaySound
Private Const SND_MEMORY = &H4         '  lpszSoundName points to a memory file
Private Const SND_NODEFAULT = &H2         '  silence not default, if sound not found
Private Const SND_NOSTOP = &H10        '  don't stop any currently playing sound
Private Const SND_SYNC = &H0         '  play synchronously (default)


Private Sub Command1_Click()
Dim tmp As Integer
 tmp = OutOfTimeSound()
End Sub

Function OutOfTimeSound()
    Call sndPlaySound("C:\Users\itplayerCR.ITPLAYER\Music\2.wav", SND_SYNC)
End Function

