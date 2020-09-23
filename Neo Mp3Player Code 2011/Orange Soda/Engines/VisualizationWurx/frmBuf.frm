VERSION 5.00
Begin VB.Form frmBuf 
   Caption         =   "Form1"
   ClientHeight    =   8835
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10815
   LinkTopic       =   "Form1"
   ScaleHeight     =   589
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   721
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox rbuf1 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      Height          =   4455
      Left            =   600
      ScaleHeight     =   293
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   333
      TabIndex        =   6
      Top             =   480
      Width           =   5055
   End
   Begin VB.PictureBox rbuf2 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      Height          =   4695
      Left            =   4320
      ScaleHeight     =   309
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   349
      TabIndex        =   5
      Top             =   360
      Width           =   5295
   End
   Begin VB.PictureBox rbuf3 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      Height          =   5535
      Left            =   1200
      ScaleHeight     =   365
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   485
      TabIndex        =   4
      Top             =   1080
      Width           =   7335
   End
   Begin VB.PictureBox buf3 
      BackColor       =   &H00000000&
      Height          =   7695
      Left            =   960
      ScaleHeight     =   509
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   621
      TabIndex        =   3
      Top             =   240
      Width           =   9375
   End
   Begin VB.PictureBox buf2 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      Height          =   8115
      Left            =   -2880
      ScaleHeight     =   537
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   659
      TabIndex        =   1
      Top             =   -120
      Width           =   9945
      Begin VB.PictureBox picBuf3 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   3000
         Left            =   600
         Picture         =   "frmBuf.frx":0000
         ScaleHeight     =   200
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   320
         TabIndex        =   2
         Top             =   720
         Width           =   4800
      End
   End
   Begin VB.PictureBox buf 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      Height          =   8175
      Left            =   0
      ScaleHeight     =   541
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   661
      TabIndex        =   0
      Top             =   720
      Width           =   9975
   End
End
Attribute VB_Name = "frmBuf"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
