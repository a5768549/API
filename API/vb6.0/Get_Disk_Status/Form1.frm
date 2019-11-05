VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   6885
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   12420
   LinkTopic       =   "Form1"
   ScaleHeight     =   6885
   ScaleWidth      =   12420
   StartUpPosition =   3  '系統預設值
   Begin VB.CommandButton Command1 
      Caption         =   "查詢"
      Height          =   1095
      Left            =   3960
      TabIndex        =   1
      Top             =   4320
      Width           =   3495
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
      Height          =   2775
      Left            =   2280
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   840
      Width           =   6855
   End
   Begin VB.Label Label1 
      Caption         =   "可儲存設備狀態"
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
      Left            =   2280
      TabIndex        =   2
      Top             =   240
      Width           =   3495
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'GetDriveType return values
Const DRIVE_REMOVABLE = 2
Const DRIVE_FIXED = 3
Const DRIVE_REMOTE = 4
Const DRIVE_CDROM = 5
Const DRIVE_RAMDISK = 6

Private Declare Function GetDriveType Lib "kernel32" Alias "GetDriveTypeA" (ByVal nDrive As String) As Long

Private Sub Command1_Click()
    Dim i As Integer
    Dim lngDrive As Long
    Dim strD As String
    Text1.Text = ""
    For i = 0 To 25  'All possible drives A to Z
        strD = Chr(i + 65) & ":\"
        lngDrive = GetDriveType(strD)
        Select Case lngDrive
            Case DRIVE_REMOVABLE
                    Text1.Text = Text1.Text & "Drive " & strD & " is removable." & vbNewLine
            Case DRIVE_FIXED
                    Text1.Text = Text1.Text & "Drive " & strD & " is fixed." & vbNewLine
            Case DRIVE_REMOTE
                    Text1.Text = Text1.Text & "Drive " & strD & " is remote." & vbNewLine
            Case DRIVE_CDROM
                    Text1.Text = Text1.Text & "Drive " & strD & " is CD-ROM." & vbNewLine
            Case DRIVE_RAMDISK
                    Text1.Text = Text1.Text & "Drive " & strD & " is RAM disk." & vbNewLine
            Case Else
        End Select
    Next i
End Sub
