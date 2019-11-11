VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   7395
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   13590
   LinkTopic       =   "Form1"
   ScaleHeight     =   7395
   ScaleWidth      =   13590
   StartUpPosition =   3  '系統預設值
   Begin VB.PictureBox Picture1 
      Height          =   3255
      Left            =   3480
      ScaleHeight     =   3195
      ScaleWidth      =   5235
      TabIndex        =   1
      Top             =   600
      Width           =   5295
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   855
      Left            =   5640
      TabIndex        =   0
      Top             =   6000
      Width           =   3135
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
 
Private Declare Function LoadImage Lib "user32" Alias "LoadImageA" (ByVal hInst As Long, ByVal lpsz As String, ByVal dwImageType As Long, ByVal dwDesiredWidth As Long, ByVal dwDesiredHeight As Long, ByVal dwFlags As Long) As Long
Private Const LR_LOADFROMFILE = &H10
Private Const LR_LOADMAP3DCOLORS = &H1000
Private Const IMAGE_ICON = 1
Private Const DI_NORMAL = 3
 
Private Declare Function DrawIconEx Lib "User32.dll" (ByVal hDC As Long, _
 ByVal xLeft As Long, ByVal yTop As Long, ByVal hIcon As Long, _
 ByVal cxWidth As Long, ByVal cyWidth As Long, ByVal istepIfAniCur As Long, _
 ByVal hbrFlickerFreeDraw As Long, ByVal diFlags As Long) As Long
 
Private Declare Function DestroyIcon Lib "User32.dll" (ByVal hIcon As Long) As Long
 
Private Sub Command1_Click() 'Command1.Style = 1 ' set to Graphical in IDE.
 lngBitmap = LoadImage(App.hInstance, " Filename.bmp", IMAGE_BITMAP, 0, 0, LR_LOADFROMFILE)
 
End Sub

Public Function CopyBitmap(SourceBitmapHandle As Long, objPictureBox As PictureBox) As Long
 
'-- I need to know what to write here, i cant figure it out
 ' First create a DC thats compatible with Destination
    lngSourceDC = CreateCompatibleDC(objPictureBox.hDC)
    
    '-- Select the Bitmap into the New DC
    lngBitmapHandle = SelectObject(lngSourceDC, SourceBitmapHandle)
    
    '-- Copy the image to the picturebox
    BitBlt objPictureBox.hDC, 0, 0, 30, 30, lngSourceDC, 0, 0, vbSrcCopy
    
    '-- Select the old pictureboxs DC and deletes memory DC
    lngRet = SelectObject(lngSourceDC, lngBitmapHandle)
    lngRet = DeleteDC(lngSourceDC)
End Function
