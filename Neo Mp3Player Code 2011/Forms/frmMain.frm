VERSION 5.00
Begin VB.Form frmOptions 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "BrowseForFolder Demo"
   ClientHeight    =   3630
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5925
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3630
   ScaleWidth      =   5925
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdExit 
      Appearance      =   0  'Flat
      Cancel          =   -1  'True
      Caption         =   "SET PARAMETERS "
      Height          =   375
      Left            =   210
      TabIndex        =   12
      Top             =   3060
      Width           =   2685
   End

   Begin VB.Menu mnuMenuPrincipal 
      Caption         =   "MenuPrincipal"
      Begin VB.Menu mnuNuevaBusqueda 
         Caption         =   "Nueva Busqueda ..."
      End
      Begin VB.Menu mnuBrowsers 
         Caption         =   "Browsers"
         Begin VB.Menu mnuExplorar 
            Caption         =   "Explorar ..."
         End
         Begin VB.Menu mnuExpAlbum 
            Caption         =   "Album(s)  Explorar "
         End
         Begin VB.Menu mnuTagEditor 
            Caption         =   " Tag Editor"
         End
         Begin VB.Menu mnuLyrics 
            Caption         =   "#Browser..#L&yrics"
         End
      End
      Begin VB.Menu mnuListSpec 
         Caption         =   "Visualization"
         Begin VB.Menu mnuMaxSpec 
            Caption         =   "Show Visualization"
         End
         Begin VB.Menu Dispaly 
            Caption         =   "Dispaly..."
         End
         Begin VB.Menu mnuShowspec 
            Caption         =   "#Visaulize..#Configure Visualization"
         End
      End
      Begin VB.Menu mnuControles 
         Caption         =   "Player Controls"
         Begin VB.Menu mnuVolumen 
            Caption         =   " Volume"
            Begin VB.Menu mnuSubirVolumen 
               Caption         =   "+  Volume up"
            End
            Begin VB.Menu mnuBajarVolumen 
               Caption         =   "#Vol...#-   Volume down"
            End
         End
         Begin VB.Menu mnuTrackAnterior 
            Caption         =   "Z   Track Anterior"
         End
         Begin VB.Menu mnuReproducir 
            Caption         =   "X   Reproducir"
         End
         Begin VB.Menu mnuPausa 
            Caption         =   "C   Pause"
         End
         Begin VB.Menu mnuDetener 
            Caption         =   "V   Stop"
         End
         Begin VB.Menu mnuSigTrack 
            Caption         =   "B   Siguiente Track"
         End
         Begin VB.Menu mnuSigAlbum 
            Caption         =   ">   Album/Folder"
         End
         Begin VB.Menu mnuAnteriorAlbum 
            Caption         =   "<   hjk Album/Folder"
         End
         Begin VB.Menu mnuIntro 
            Caption         =   "I   Intro 10 Seconds"
         End
         Begin VB.Menu mnuRepetir 
            Caption         =   "R   Repeat Track"
         End
         Begin VB.Menu mnuSilencio 
            Caption         =   "S   Mute"
         End
         Begin VB.Menu mnuOrdenAleatorio 
            Caption         =   " Random"
            Begin VB.Menu mnuAleatorioActAlbum 
               Caption         =   "Q   Album/Folder"
            End
            Begin VB.Menu mnuAleatorioTodaColec 
               Caption         =   "W   Total track"
            End
            Begin VB.Menu repCurTack 
               Caption         =   "#Randomize#current track"
            End
         End
         Begin VB.Menu mnuAtras5Seg 
            Caption         =   "A   seek  5 Seconds"
         End
         Begin VB.Menu mnuAdelante5Seg 
            Caption         =   "#Player controls.....#Back 5 seconds"
         End
      End
      Begin VB.Menu mnuOpciones 
         Caption         =   "Opcion ..."
      End
      Begin VB.Menu mnuSkins 
         Caption         =   "Skins"
         WindowList      =   -1  'True
         Begin VB.Menu mnuExpSkins 
            Caption         =   "<<  Skins Browser >>"
         End
         Begin VB.Menu mnuSkinsAdd 
            Caption         =   "#Skins#Get more Skins..."
            Index           =   25
         End
      End
      Begin VB.Menu mnuWOpacity 
         Caption         =   "Window Opacity"
         Begin VB.Menu mnuAlpha 
            Caption         =   "100%"
            Index           =   0
         End
         Begin VB.Menu mnuAlpha 
            Caption         =   "90%"
            Index           =   1
         End
         Begin VB.Menu mnuAlpha 
            Caption         =   "80%"
            Index           =   2
         End
         Begin VB.Menu mnuAlpha 
            Caption         =   "70%"
            Index           =   3
         End
         Begin VB.Menu mnuAlpha 
            Caption         =   "60%"
            Index           =   4
         End
         Begin VB.Menu mnuAlpha 
            Caption         =   "50%"
            Index           =   5
         End
         Begin VB.Menu mnuAlpha 
            Caption         =   "40%"
            Index           =   6
         End
         Begin VB.Menu mnuAlpha 
            Caption         =   "30%"
            Index           =   7
         End
         Begin VB.Menu mnuAlpha 
            Caption         =   "20%"
            Index           =   8
         End
         Begin VB.Menu mnuAlpha 
            Caption         =   "10%"
            Index           =   9
         End
         Begin VB.Menu mnuAlphaPer 
            Caption         =   "#Opacity#Custom.."
         End
      End
      Begin VB.Menu mnuAcercaDe 
         Caption         =   "About me ..."
      End
      Begin VB.Menu mnuSalir 
         Caption         =   "#Mahesh Mp3 Player#Exit"
      End
   End
   Begin VB.Menu mnuMainSpec 
      Caption         =   "MainSpectrum"
      Begin VB.Menu mnuSpecNone 
         Caption         =   "None Visualisation"
      End
      Begin VB.Menu mnuSpecBars 
         Caption         =   "Spectrum Analyzer"
      End
      Begin VB.Menu mnuSpecOsc 
         Caption         =   "#Spectrum#Oscilloscope"
      End
   End
   Begin VB.Menu mnuMainAlbum 
      Caption         =   "MainAlbum"
      Begin VB.Menu mnuAlbumTags 
         Caption         =   "Edit Album Tags"
      End
      Begin VB.Menu mnuAlbumBrowser 
         Caption         =   "Explore in Album Browser"
      End
      Begin VB.Menu mnuAlbumExp 
         Caption         =   "Explore in Explorer.exe"
      End
      Begin VB.Menu mnuAlbumPlay 
         Caption         =   "#Album#Play"
      End
   End
End



   Begin VB.ListBox List1 
      BackColor       =   &H8000000F&
      Height          =   2595
      Left            =   3090
      TabIndex        =   11
      Top             =   420
      Width           =   2715
   End
   Begin VB.ComboBox cmbSpecial 
      Height          =   315
      ItemData        =   "frmMain.frx":0000
      Left            =   120
      List            =   "frmMain.frx":0049
      TabIndex        =   7
      Top             =   990
      Width           =   2895
   End
   Begin VB.CheckBox chkFlags 
      Caption         =   "New Style"
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   5
      Top             =   1680
      Width           =   2295
   End
   Begin VB.CommandButton cmdBrowse 
      Appearance      =   0  'Flat
      Caption         =   "APPLY"
      Default         =   -1  'True
      Height          =   315
      Left            =   3120
      TabIndex        =   4
      Top             =   3090
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      Caption         =   "Options"
      Height          =   2115
      Left            =   120
      TabIndex        =   3
      Top             =   1410
      Width           =   2895
      Begin VB.CheckBox chkOK 
         Caption         =   "Disable OK Button"
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   1200
         Width           =   2295
      End
      Begin VB.CheckBox chkFlags 
         Caption         =   "Include Files"
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   9
         Top             =   960
         Width           =   2295
      End
      Begin VB.CheckBox chkFlags 
         Caption         =   "Current Selection Label"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   8
         Top             =   720
         Width           =   2295
      End
      Begin VB.CheckBox chkFlags 
         Caption         =   "Edit Box"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   6
         Top             =   480
         Width           =   2295
      End
      Begin VB.Line Line3 
         BorderColor     =   &H80000014&
         X1              =   120
         X2              =   2760
         Y1              =   1560
         Y2              =   1560
      End
      Begin VB.Line Line4 
         BorderColor     =   &H80000010&
         BorderWidth     =   2
         X1              =   120
         X2              =   2760
         Y1              =   1590
         Y2              =   1590
      End
   End
   Begin VB.TextBox txtInstruction 
      Height          =   285
      Left            =   120
      TabIndex        =   0
      Text            =   "Type Your Own Instruction Here."
      Top             =   420
      Width           =   2895
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000014&
      X1              =   3120
      X2              =   5820
      Y1              =   3480
      Y2              =   3480
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000010&
      BorderWidth     =   2
      X1              =   3930
      X2              =   6360
      Y1              =   4200
      Y2              =   4200
   End
   Begin VB.Label Label2 
      Caption         =   "Open At:"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   780
      Width           =   2895
   End
   Begin VB.Label Label1 
      Caption         =   "Instruction Message:"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   150
      Width           =   2895
   End
End
Attribute VB_Name = "frmOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub chkOK_Click()
    OKEnable = Not OKEnable
End Sub

Private Sub cmbSpecial_Change()
    cmbSpecial_Click
End Sub

Private Sub cmbSpecial_Click()
    Dim i As Integer

    If cmbSpecial.ListIndex > 0 Then
        chkFlags(0).Value = 0
        chkFlags(0).Enabled = False
    Else
        chkFlags(0).Enabled = True
    End If

End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdBrowse_Click()
    Dim Flags As Long
    Dim i As Integer

    If optBrowse(0) Then
        Select Case cmbSpecial.ListIndex
            Case 0 To 11: SpecialFolder = cmbSpecial.ListIndex
            Case 12 To 17: SpecialFolder = cmbSpecial.ListIndex + 4
            Case 18 To 19: SpecialFolder = cmbSpecial.ListIndex + 8
            Case 20 To 22: SpecialFolder = cmbSpecial.ListIndex + 12
            Case Else: StartFolder = "F:\songs\03   Who Lamhe"
        End Select

        If chkFlags(0).Value = 1 Then Flags = Flags + BIF_USENEWUI
        If chkFlags(1).Value = 1 Then Flags = Flags + BIF_EDITBOX
        If chkFlags(2).Value = 1 Then Flags = Flags + BIF_STATUSTEXT
        If chkFlags(3).Value = 1 Then Flags = Flags + BIF_BROWSEINCLUDEFILES
    Else
        Flags = BIF_BROWSEFORPRINTER
        SpecialFolder = 4
    End If

  '  txtReturn = FolderBrowse(Me.hWnd, txtInstruction, Flags)
    
    If szDisplay <> "Printers" And szDisplay <> "Add Printer" Then
     '   txtDisplay = szDisplay
    Else
      '  txtDisplay = ""
    End If

    SpecialFolder = 0
    StartFolder = ""
    szDisplay = ""
    
End Sub

Private Sub Form_Load()
    cmbSpecial.ListIndex = 0
    cmbSpecial = CurDir
    OKEnable = True
    Dim i
For i = 0 To Screen.FontCount - 1
List1.AddItem Screen.Fonts(i) ', "1234654" 'Screen.Fonts(i)
Next
End Sub

Private Sub Form_Unload(Cancel As Integer)
    End
End Sub
