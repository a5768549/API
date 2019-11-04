VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   7260
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   13965
   LinkTopic       =   "Form1"
   ScaleHeight     =   7260
   ScaleWidth      =   13965
   StartUpPosition =   3  '系統預設值
   Begin VB.CheckBox Check5 
      Caption         =   "全刪"
      Height          =   495
      Left            =   6600
      TabIndex        =   16
      Top             =   5520
      Width           =   1935
   End
   Begin VB.CommandButton Command3 
      Caption         =   "刪除"
      Height          =   855
      Left            =   8760
      TabIndex        =   15
      Top             =   4560
      Width           =   2415
   End
   Begin VB.CheckBox Check4 
      Caption         =   "Value4"
      Height          =   495
      Left            =   6600
      TabIndex        =   14
      Top             =   4920
      Width           =   1935
   End
   Begin VB.CheckBox Check3 
      Caption         =   "Value3"
      Height          =   495
      Left            =   6600
      TabIndex        =   13
      Top             =   4320
      Width           =   1935
   End
   Begin VB.CheckBox Check2 
      Caption         =   "Value2"
      Height          =   495
      Left            =   6600
      TabIndex        =   12
      Top             =   3720
      Width           =   1935
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Value1"
      Height          =   495
      Left            =   6600
      TabIndex        =   11
      Top             =   3120
      Width           =   1935
   End
   Begin VB.TextBox Text5 
      Height          =   735
      Left            =   11040
      TabIndex        =   6
      Text            =   "0"
      Top             =   1440
      Width           =   2415
   End
   Begin VB.TextBox Text4 
      Height          =   735
      Left            =   6720
      TabIndex        =   5
      Text            =   "0"
      Top             =   1440
      Width           =   2415
   End
   Begin VB.TextBox Text3 
      Height          =   735
      Left            =   11040
      TabIndex        =   4
      Text            =   "0"
      Top             =   360
      Width           =   2415
   End
   Begin VB.TextBox Text2 
      Height          =   735
      Left            =   6720
      TabIndex        =   3
      Text            =   "0"
      Top             =   360
      Width           =   2415
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
      Height          =   3855
      Left            =   240
      MultiLine       =   -1  'True
      TabIndex        =   2
      Top             =   2880
      Width           =   4935
   End
   Begin VB.CommandButton Command2 
      Caption         =   "讀入"
      Height          =   855
      Left            =   240
      TabIndex        =   1
      Top             =   1680
      Width           =   2775
   End
   Begin VB.CommandButton Command1 
      Caption         =   "寫入"
      Height          =   975
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   2775
   End
   Begin VB.Label Label4 
      Caption         =   "Value4"
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
      Left            =   9720
      TabIndex        =   10
      Top             =   1560
      Width           =   1455
   End
   Begin VB.Label Label3 
      Caption         =   "Value3"
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
      Left            =   5280
      TabIndex        =   9
      Top             =   1680
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "Value2"
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
      Left            =   9720
      TabIndex        =   8
      Top             =   480
      Width           =   1575
   End
   Begin VB.Label Label1 
      Caption         =   "Value1"
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
      Left            =   5280
      TabIndex        =   7
      Top             =   600
      Width           =   1335
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    SaveSetting "MyValue", "Form1", "Value1", Text2.Text
    SaveSetting "MyValue", "Form1", "Value2", Text3.Text
    SaveSetting "MyValue", "Form1", "Value3", Text4.Text
    SaveSetting "MyValue", "Form1", "Value4", Text5.Text
End Sub

Private Sub Command2_Click()
    Dim values As Variant, i As Long
    values = GetAllSettings("MyValue", "Form1")
    ' Each row holds two items, the key name and the key value.
    Text1.Text = ""
    For i = 0 To UBound(values)
        Text1.Text = Text1.Text & "Key =" & values(i, 0) & "  Value = " & values(i, 1) & vbNewLine
    Next
End Sub

Private Sub Command3_Click()
Select Case True
    Case Check1.Value: DeleteSetting "MyValue", "Form1", "Value1"
    Case Check2.Value: DeleteSetting "MyValue", "Form1", "Value2"
    Case Check3.Value: DeleteSetting "MyValue", "Form1", "Value3"
    Case Check4.Value: DeleteSetting "MyValue", "Form1", "Value4"
    Case Check5.Value: DeleteSetting "MyValue", "Form1"
End Select
    
End Sub
