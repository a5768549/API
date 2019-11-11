VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4185
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7575
   LinkTopic       =   "Form1"
   ScaleHeight     =   4185
   ScaleWidth      =   7575
   StartUpPosition =   3  '系統預設值
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   735
      Left            =   1200
      TabIndex        =   0
      Top             =   1080
      Width           =   1335
   End
   Begin VB.Image Image1 
      Height          =   2415
      Left            =   3360
      Top             =   960
      Width           =   3615
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function OleLoadPicturePath Lib "oleaut32.dll" (ByVal szURLorPath As Long, ByVal punkCaller As Long, ByVal dwReserved As Long, ByVal clrReserved As OLE_COLOR, ByRef riid As TGUID, ByRef ppvRet As IPicture) As Long

Private Type TGUID
Data1 As Long
Data2 As Integer
Data3 As Integer
Data4(0 To 7) As Byte
End Type


Public Function LoadPicture(ByVal strFileName As String) As Picture
Dim IID As TGUID
With IID
.Data1 = &H7BF80980
.Data2 = &HBF32
.Data3 = &H101A
.Data4(0) = &H8B
.Data4(1) = &HBB
.Data4(2) = &H0
.Data4(3) = &HAA
.Data4(4) = &H0
.Data4(5) = &H30
.Data4(6) = &HC
.Data4(7) = &HAB
End With

On Error GoTo LocalErr

OleLoadPicturePath StrPtr(strFileName), 0&, 0&, 0&, IID, LoadPicture
Exit Function
LocalErr:
Set LoadPicture = VB.LoadPicture(strFileName)
Err.Clear
End Function

Private Sub Command1_Click()

Dim b() As Byte
Dim strURL As String

Image1.Picture = LoadPicture("c:/csdn.jpg") 'Only JPG
End Sub
