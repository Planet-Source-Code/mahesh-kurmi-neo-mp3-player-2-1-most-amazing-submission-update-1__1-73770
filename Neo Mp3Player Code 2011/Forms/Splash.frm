VERSION 5.00
Begin VB.Form frmSplash 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   3075
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6750
   Icon            =   "Splash.frx":0000
   LinkTopic       =   "Form1"
   MousePointer    =   99  'Custom
   Picture         =   "Splash.frx":000C
   ScaleHeight     =   3075
   ScaleWidth      =   6750
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Label lblSplash 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   165
      Index           =   2
      Left            =   135
      MousePointer    =   99  'Custom
      TabIndex        =   2
      Top             =   105
      Width           =   45
   End
   Begin VB.Label lblSplash 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   1
      Left            =   150
      MousePointer    =   99  'Custom
      TabIndex        =   1
      Top             =   1560
      Width           =   4740
   End
   Begin VB.Label lblSplash 
      BackStyle       =   0  'Transparent
      Caption         =   "Loading......."
      BeginProperty Font 
         Name            =   "Copperplate Gothic Bold"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404000&
      Height          =   240
      Index           =   0
      Left            =   3480
      MousePointer    =   99  'Custom
      TabIndex        =   0
      Top             =   120
      Width           =   4665
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Click()
    bolSplashScreen = False
    Unload Me
End Sub

Private Sub Form_Load()
'bolSplashScreen = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
    bolSplashScreen = False
End Sub


