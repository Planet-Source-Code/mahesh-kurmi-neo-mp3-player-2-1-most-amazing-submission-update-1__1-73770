VERSION 5.00
Begin VB.Form frmDummy 
   Caption         =   "Form1"
   ClientHeight    =   8100
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9555
   LinkTopic       =   "Form1"
   ScaleHeight     =   8100
   ScaleWidth      =   9555
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox buf 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFFF&
      Height          =   390
      Left            =   2880
      ScaleHeight     =   26
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   130
      TabIndex        =   1
      Top             =   3000
      Width           =   1950
   End
   Begin VB.PictureBox picBuf3 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   4500
      Left            =   0
      ScaleHeight     =   300
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   420
      TabIndex        =   4
      Top             =   1680
      Width           =   6300
   End
   Begin VB.PictureBox picBuf2 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      Height          =   6015
      Left            =   1200
      ScaleHeight     =   397
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   413
      TabIndex        =   3
      Top             =   720
      Width           =   6255
   End
   Begin VB.PictureBox picBuf 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      Height          =   5175
      Left            =   3120
      ScaleHeight     =   341
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   429
      TabIndex        =   2
      Top             =   2280
      Width           =   6495
   End
   Begin VB.PictureBox picDummy 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      Height          =   4935
      Left            =   2280
      ScaleHeight     =   325
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   413
      TabIndex        =   0
      Top             =   0
      Width           =   6255
   End
End
Attribute VB_Name = "frmDummy"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
