VERSION 5.00
Begin VB.Form frmSpectrum 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFC0&
   BorderStyle     =   0  'None
   Caption         =   "Spectrum"
   ClientHeight    =   2010
   ClientLeft      =   105
   ClientTop       =   105
   ClientWidth     =   4440
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   134
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   296
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picSpectrum 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      FillStyle       =   0  'Solid
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   345
      Left            =   90
      ScaleHeight     =   23
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   78
      TabIndex        =   1
      Top             =   210
      Width           =   1170
      Begin VB.Label Label 
         BackColor       =   &H00000000&
         Caption         =   "Visualizacion"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   225
         Left            =   -60
         TabIndex        =   2
         Top             =   30
         Visible         =   0   'False
         Width           =   4245
      End
   End
   Begin VB.PictureBox picFront 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   645
      Left            =   2895
      ScaleHeight     =   43
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   57
      TabIndex        =   0
      Top             =   2835
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Timer Timer 
      Enabled         =   0   'False
      Interval        =   2000
      Left            =   4860
      Top             =   1080
   End
End
Attribute VB_Name = "frmSpectrum"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim bLoadingVis As Boolean
Dim InFormDrag As Boolean
Dim cWindows As New cWindowSkin
Dim cAjustarDesk As New clsDockingHandler

Private Sub Form_Resize()
 On Error Resume Next

' If tConfigVis.Exist = False Then Exit Sub

' picSpectrum.Cls
' picSpectrum.PaintPicture picFront.Picture, 0, 0, picSpectrum.ScaleWidth, picSpectrum.ScaleHeight, 0, 0
' picSpectrum.Picture = picSpectrum.Image

 DoEvents

End Sub


Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
cWindows.Formulario_Down x, Y
            cAjustarDesk.StartDockDrag x * Screen.TwipsPerPixelX, _
                Y * Screen.TwipsPerPixelY
InFormDrag = True
    
End Sub


Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
 cWindows.Formulario_MouseMove Button, x, Y

End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
cWindows.Formulario_MouseUp x, Y
If cWindows.ClickExitButton = True Then

End If
InFormDrag = False
    
End Sub


Private Sub picSpectrum_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
   ' If Button = vbRightButton Then PopupMenu frmPopUp.mnuSpectrum

End Sub



Private Sub Form_Load()
  Me.Left = (Screen.Width - Me.Width) / 2   '// centrar formulario
  Me.Top = (Screen.Height - Me.Height) / 2
  Call cargar_formulario

End Sub
Sub cargar_formulario()
On Error Resume Next
Dim ix As Integer, iy As Integer

  Set cWindows.FormularioPadre = Me
  Set cAjustarDesk.ParentForm = Me
  cWindows.ColorInvisible = Read_INI("NormalMode", "ColorTrans", RGB(255, 0, 255), True)
  ix = Read_INI("Configuration", "ExitButtonX", 1, False, False, tAppConfig.AppConfig & "Skins\" & tAppConfig.Skin & "\VIS_STUDIO\config.ini")
  iy = Read_INI("Configuration", "ExitButtonY", 1, False, False, tAppConfig.AppConfig & "Skins\" & tAppConfig.Skin & "\VIS_STUDIO\config.ini")
  cWindows.ButtonExitXY CLng(ix), CLng(iy)
  cWindows.MinimoAlto = Read_INI("Configuration", "MinHeight", 10, False, False, tAppConfig.AppConfig & "Skins\" & tAppConfig.Skin & "\VIS_STUDIO\config.ini")
  cWindows.MinimoAncho = Read_INI("Configuration", "MinWidth", 10, False, False, tAppConfig.AppConfig & "Skins\" & tAppConfig.Skin & "\VIS_STUDIO\config.ini")
  
  cWindows.CargarSkin tAppConfig.AppConfig & "Skins\" & tAppConfig.Skin & "\VIS_STUDIO\"
  'picSpectrum.Left = cWindows.AreaLeft
  'picSpectrum.Top = cWindows.AreaTop
  'picSpectrum.Width = cWindows.AreaWidth
  'picSpectrum.Height = cWindows.AreaHeight
  'Label.Width = cWindows.AreaWidth

End Sub


Private Sub Form_Unload(Cancel As Integer)
  Me.Hide
  bolVisShow = False
  Cancel = 1
End Sub


