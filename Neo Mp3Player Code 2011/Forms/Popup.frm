VERSION 5.00
Begin VB.Form frmPopUp 
   Caption         =   "MaheshMP3 Player"
   ClientHeight    =   4875
   ClientLeft      =   60
   ClientTop       =   150
   ClientWidth     =   13335
   ClipControls    =   0   'False
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   HasDC           =   0   'False
   Icon            =   "Popup.frx":0000
   LinkTopic       =   "form1"
   Moveable        =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   325
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   889
   ShowInTaskbar   =   0   'False
   Visible         =   0   'False
   Begin MMPlayerXProject.vkSysTray vkSysTrayIcon 
      Index           =   5
      Left            =   8880
      Top             =   240
      _extentx        =   794
      _extenty        =   794
      balloontipstring=   "Next track"
      icon            =   "Popup.frx":000C
   End
   Begin MMPlayerXProject.vkSysTray vkSysTrayIcon 
      Index           =   4
      Left            =   8280
      Top             =   240
      _extentx        =   794
      _extenty        =   794
      balloontipstring=   "Stop"
      icon            =   "Popup.frx":05A8
   End
   Begin MMPlayerXProject.vkSysTray vkSysTrayIcon 
      Index           =   3
      Left            =   7680
      Top             =   240
      _extentx        =   794
      _extenty        =   794
      balloontipstring=   "Pause"
      icon            =   "Popup.frx":0B44
   End
   Begin MMPlayerXProject.vkSysTray vkSysTrayIcon 
      Index           =   2
      Left            =   7080
      Top             =   240
      _extentx        =   794
      _extenty        =   794
      balloontipstring=   "Play"
      icon            =   "Popup.frx":10E0
   End
   Begin MMPlayerXProject.vkSysTray vkSysTrayIcon 
      Index           =   1
      Left            =   6480
      Top             =   240
      _extentx        =   794
      _extenty        =   794
      balloontipstring=   "Previous Track"
      icon            =   "Popup.frx":167C
   End
   Begin VB.ListBox lstLanguage 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   645
      Left            =   105
      TabIndex        =   1
      Top             =   3300
      Width           =   7470
   End
   Begin VB.FileListBox fileBmps 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1590
      Hidden          =   -1  'True
      Left            =   1230
      Pattern         =   "*.jpg;*.bmp"
      System          =   -1  'True
      TabIndex        =   0
      Top             =   180
      Visible         =   0   'False
      Width           =   2880
   End
   Begin MMPlayerXProject.vkCommand CMD 
      Height          =   435
      Left            =   1680
      TabIndex        =   2
      Top             =   2580
      Width           =   1245
      _extentx        =   2196
      _extenty        =   767
      caption         =   ""
      font            =   "Popup.frx":1C18
      picture         =   "Popup.frx":1C40
      pictureoffsetx  =   120
      pictureoffsety  =   -20
   End
   Begin MMPlayerXProject.vkSysTray vkSysTrayIcon 
      Index           =   0
      Left            =   6480
      Top             =   960
      _extentx        =   794
      _extenty        =   794
      balloontipstring=   "Neo MP3 Player: By Mahesh Kurmi"
      icon            =   "Popup.frx":2092
   End
   Begin VB.Menu mnuMenuPrincipal 
      Caption         =   "MenuPrincipal"
      Begin VB.Menu mnuAbout 
         Caption         =   "MaheshMp3 Player..."
      End
      Begin VB.Menu bar 
         Caption         =   "-"
         Index           =   0
      End
      Begin VB.Menu mnuPlay 
         Caption         =   "Play"
         Begin VB.Menu mnuPlayfile 
            Caption         =   "File..."
         End
         Begin VB.Menu mnuPlayfolder 
            Caption         =   "Folder..."
         End
         Begin VB.Menu mnuPlay_playlist 
            Caption         =   "Playlist..."
         End
         Begin VB.Menu bar11 
            Caption         =   "-"
         End
         Begin VB.Menu mnuPlayRemovableMedia 
            Caption         =   "Removable Media"
         End
         Begin VB.Menu bar13 
            Caption         =   "-"
         End
      End
      Begin VB.Menu mnuLibrary 
         Caption         =   "Media Library"
      End
      Begin VB.Menu mnuTagEditor 
         Caption         =   "Tag Editor"
      End
      Begin VB.Menu mnuSearchDirectory 
         Caption         =   "Mp3 Searcher"
      End
      Begin VB.Menu bar2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuDSP 
         Caption         =   "DFX Editor"
      End
      Begin VB.Menu mnuShowEqualizer 
         Caption         =   "Equalizers"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuShowPlst 
         Caption         =   "Playlist Editor"
      End
      Begin VB.Menu mnuListSpec 
         Caption         =   "Visualization"
         Begin VB.Menu mnuShowVis 
            Caption         =   "Show Visualization"
         End
         Begin VB.Menu Dispaly 
            Caption         =   "Dispaly..."
         End
         Begin VB.Menu mnuShowspec 
            Caption         =   "Configure Visualization"
         End
      End
      Begin VB.Menu bar17 
         Caption         =   "-"
      End
      Begin VB.Menu mnuControles 
         Caption         =   "Player Controls"
         Begin VB.Menu mnuVolumen 
            Caption         =   "Volume"
            Begin VB.Menu mnuSubirVolumen 
               Caption         =   "+  Volume up"
            End
            Begin VB.Menu mnuBajarVolumen 
               Caption         =   "-   Volume down"
            End
         End
         Begin VB.Menu bar18 
            Caption         =   "-"
         End
         Begin VB.Menu mnuTrackAnterior 
            Caption         =   "Previous Track"
         End
         Begin VB.Menu mnuplayit 
            Caption         =   "Play"
         End
         Begin VB.Menu mnuPausa 
            Caption         =   "Pause"
         End
         Begin VB.Menu mnuDetener 
            Caption         =   "Stop"
         End
         Begin VB.Menu mnuSigTrack 
            Caption         =   "Next Track"
         End
         Begin VB.Menu bar19 
            Caption         =   "-"
         End
         Begin VB.Menu mnuForward5Sec 
            Caption         =   "Seek  5 Seconds forward"
         End
         Begin VB.Menu mnuBack5Sec 
            Caption         =   "Back 5 Seconds Backward"
         End
      End
      Begin VB.Menu bar10 
         Caption         =   "-"
      End
      Begin VB.Menu mnuOpciones 
         Caption         =   "Options ..."
         Begin VB.Menu Atop 
            Caption         =   "Always on Top"
         End
         Begin VB.Menu mnuRegisterfiles 
            Caption         =   "Register MP3 files"
         End
         Begin VB.Menu bar6 
            Caption         =   "-"
         End
         Begin VB.Menu mnuCrossfade 
            Caption         =   "Crossfade"
         End
         Begin VB.Menu mnuIntro 
            Caption         =   "Intro 10 Seconds"
         End
         Begin VB.Menu mnuRepeatTrack 
            Caption         =   "Repeat Track"
         End
         Begin VB.Menu mnuSilencio 
            Caption         =   "Mute"
         End
         Begin VB.Menu mnuOrdenAleatorio 
            Caption         =   "Shuffle"
         End
         Begin VB.Menu bar16 
            Caption         =   "-"
         End
         Begin VB.Menu mnuPreferences 
            Caption         =   "Preferences"
         End
         Begin VB.Menu bar24 
            Caption         =   "-"
         End
         Begin VB.Menu mnuReset 
            Caption         =   "Reset player"
         End
      End
      Begin VB.Menu mnuChangeMask 
         Caption         =   "Change mask"
      End
      Begin VB.Menu mnuSkins 
         Caption         =   "Skins"
         WindowList      =   -1  'True
         Begin VB.Menu mnuExpSkins 
            Caption         =   "<<  Skins Browser >>"
         End
         Begin VB.Menu bar12 
            Caption         =   "-"
         End
         Begin VB.Menu mnuSkinsAdd 
            Caption         =   "Get more Skins..."
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
         Begin VB.Menu bar22 
            Caption         =   "-"
         End
         Begin VB.Menu mnuAlphaPer 
            Caption         =   "Custom.."
         End
      End
      Begin VB.Menu bar8 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHelp 
         Caption         =   "Player Help"
      End
      Begin VB.Menu bar7 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "Exit"
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
         Caption         =   "Oscilloscope"
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
         Caption         =   "Play"
      End
   End
   Begin VB.Menu mnuPlstOptions 
      Caption         =   "mnuPlstOptions"
      Begin VB.Menu mnuNew_List 
         Caption         =   "New List"
         Shortcut        =   ^N
      End
      Begin VB.Menu mnuSave_List 
         Caption         =   "Save List..."
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuLoad_List 
         Caption         =   "Load List..."
         Shortcut        =   ^L
      End
   End
   Begin VB.Menu mnuadd 
      Caption         =   "mnuAdd tracks"
      Begin VB.Menu mnuaddFile 
         Caption         =   "Add Files..."
         Shortcut        =   ^F
      End
      Begin VB.Menu mnuadddir 
         Caption         =   "Add Folder..."
         Shortcut        =   ^D
      End
      Begin VB.Menu bar30 
         Caption         =   "-"
      End
      Begin VB.Menu mnuaddPlst 
         Caption         =   "Add Playlist..."
         Shortcut        =   ^P
      End
   End
   Begin VB.Menu mnuPlst_rightClick 
      Caption         =   "Plst_rightClick"
      Begin VB.Menu mnuPlayTrack 
         Caption         =   "Play Track(s)                           Enter"
      End
      Begin VB.Menu mnuRemoveTracks 
         Caption         =   "Remove Track(s)                     Delete"
      End
      Begin VB.Menu mnuCropTracks 
         Caption         =   "Crop files                                 Ctrl+Delete"
      End
      Begin VB.Menu bar25 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExploreItem 
         Caption         =   "Explore Item's directory"
      End
      Begin VB.Menu mnuFileinfo 
         Caption         =   "View file Info...                       Ctrl+F"
      End
   End
   Begin VB.Menu mnuVis 
      Caption         =   "Visualization"
      Begin VB.Menu mnuGroupName 
         Caption         =   "Visualizations"
         Begin VB.Menu mnuObjectName 
            Caption         =   ""
            Checked         =   -1  'True
            Index           =   0
         End
      End
      Begin VB.Menu mnunull3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuLoadPi 
         Caption         =   "Import Visualization list ..."
      End
      Begin VB.Menu mnuConfig 
         Caption         =   "Configure this visualization ..."
      End
      Begin VB.Menu mnuNull5 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAdjustLevels 
         Caption         =   "Adjust Levels*"
      End
      Begin VB.Menu regPlugins 
         Caption         =   "Register Plugins"
      End
   End
   Begin VB.Menu mnuPlstSel 
      Caption         =   "mnuPLSTSel"
      Begin VB.Menu mnuSelectAll 
         Caption         =   "Select All"
         Shortcut        =   ^A
      End
      Begin VB.Menu mnuSelectNone 
         Caption         =   "Select None"
         Shortcut        =   ^Q
      End
      Begin VB.Menu mnuInverSelection 
         Caption         =   "Invert Selection"
         Shortcut        =   ^I
      End
   End
   Begin VB.Menu mnuPlstMis 
      Caption         =   "mnuMIS"
      Begin VB.Menu mnuSort 
         Caption         =   "Sort List"
         Begin VB.Menu mnuSortByFilename 
            Caption         =   "Sort By Filename"
         End
         Begin VB.Menu mnuSortbyList 
            Caption         =   "Sort By playlist entry"
         End
         Begin VB.Menu mnuSortbyDuration 
            Caption         =   "Sort By Duration"
         End
      End
      Begin VB.Menu bar27 
         Caption         =   "-"
      End
      Begin VB.Menu mnuReverselist 
         Caption         =   "Reverse List"
      End
      Begin VB.Menu mnuRandomizeList 
         Caption         =   "Randomize List"
      End
      Begin VB.Menu bar28 
         Caption         =   "-"
      End
      Begin VB.Menu mnuIDEInfo 
         Caption         =   "View File Information"
      End
   End
   Begin VB.Menu mnuPLSTRemove 
      Caption         =   "mnuPLSTRemove"
      Begin VB.Menu mnuRemoveSelected 
         Caption         =   "Remove Selected"
      End
      Begin VB.Menu bar29 
         Caption         =   "-"
      End
      Begin VB.Menu mnuRemoveMissingFiles 
         Caption         =   "Remove Missing Files"
      End
      Begin VB.Menu mnuRemoveDuplicateFiles 
         Caption         =   "Remove Duplicate Entries"
      End
      Begin VB.Menu mnuRemoveCorruptedFiles 
         Caption         =   "Remove corrupted Files"
      End
      Begin VB.Menu bar31 
         Caption         =   "-"
      End
      Begin VB.Menu mnuRemoveAll 
         Caption         =   "Remove All Tracks"
      End
   End
End
Attribute VB_Name = "frmPopUp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Sub Atop_Click()    'Always on top
    Atop.Checked = Not Atop.Checked
    OpcionesMusic.SiempreTop = Atop.Checked
    Always_on_Top
    If Atop.Checked = False Then
        frmMain.PicaTop.PaintPicture frmMain.PicConfigInfo, 0, 0, frmMain.PicaTop.Width, frmMain.PicaTop.Height, frmMain.PicShuffle.Width + frmMain.PicRepeat.Width + frmMain.PicCrossfade.Width, 0, frmMain.PicaTop.Width, frmMain.PicaTop.Height
        frmMain.PicaTop.Picture = frmMain.PicaTop.Image
    Else
        frmMain.PicaTop.PaintPicture frmMain.PicConfigInfo, 0, 0, frmMain.PicaTop.Width, frmMain.PicaTop.Height, frmMain.PicShuffle.Width + frmMain.PicRepeat.Width + frmMain.PicCrossfade.Width, frmMain.PicaTop.Height, frmMain.PicaTop.Width, frmMain.PicaTop.Height
        frmMain.PicaTop.Picture = frmMain.PicaTop.Image
    End If
End Sub

Private Sub Dispaly_Click()
    frmSearch.Show
End Sub

Private Sub Form_Load()
    Call SetIcon(Me.hwnd, "APPICON", True)
    Me.Top = -Me.Height - 800
    Me.Icon = frmMain.Icon
    Me.Left = frmMain.Left    ' for real animation in minimizing
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim i As Long
    On Error Resume Next
    For i = 0 To 5
        vkSysTrayIcon(i).RemoveFromTray (i)
    Next
End Sub

Private Sub mnuAbout_Click()
    frmMain.ScrollText(1).CaptionText = "--MAHESH MP3 PLAYER--"    '& (index + 1) & " " & Str(Eq_SliderCtrl(index).Value) & " kHz"
    frmAbout.Show
End Sub

Private Sub mnuadddir_Click()
    Call frmPLST.ADD_Dir(True, False)
End Sub

Private Sub mnuaddFile_Click()
    frmPLST.ADD_MULTIPLE_FILES (False)
End Sub

Private Sub mnuaddPlst_Click()
    On Error Resume Next
    Dim openPath As String
    Dim sfilter As String
    sfilter = "All Supported Playlist Files|*.npl;*.m3u;*.pls|" & "Playlist Files (*.npl)|*.npl|" & "Winamp Playlistfile (.m3u)|*.m3u|" & "MediaPlayer Pls file (.PLS)|*.pls|"
    openPath = clsDlg.GetOpenAsName(frmPLST.hwnd, "Open Playlist for NeoMp3 player", , sfilter, ALLOWMULTISELECT Or FILEMUSTEXIST Or LONGNAMES Or EXPLORER)
    If openPath = "" Then Exit Sub
    Call LoadPlaylist(openPath, False)
    frmPLST.Update_Plst_Scrollbar
    frmPLST.Scrollbar.Value = frmPLST.Scrollbar.Max
    frmPLST.ReinitializeList
End Sub

Private Sub mnuBack5Sec_Click()
    frmMain.Five_Seg_Backward
End Sub

Public Sub mnuAlpha_Click(Index As Integer)
    On Error GoTo hell
    Dim tAlpha
    Dim i As Integer
    tAlpha = mnuAlpha(Index).Caption
    tAlpha = Left(tAlpha, Len(tAlpha) - 1)
    Call SetWindowLong(frmMain.hwnd, GWL_EXSTYLE, GetWindowLong(frmMain.hwnd, GWL_EXSTYLE) Or WS_EX_LAYERED)
    Call SetLayeredWindowAttributes(frmMain.hwnd, 0, (255 * tAlpha) / 100, LWA_ALPHA)
    Call SetWindowLong(frmPLST.hwnd, GWL_EXSTYLE, GetWindowLong(frmPLST.hwnd, GWL_EXSTYLE) Or WS_EX_LAYERED)
    Call SetLayeredWindowAttributes(frmPLST.hwnd, 0, (255 * tAlpha) / 100, LWA_ALPHA)

    mnuAlpha(Index).Checked = True
    OpcionesMusic.Alpha = tAlpha
    frmPopUp.mnuAlphaPer.Caption = "Custom..."
    frmPopUp.mnuAlphaPer.Checked = False

    For i = 0 To 9
        If i <> Index Then mnuAlpha(i).Checked = False
    Next i
    Exit Sub
hell:
End Sub

Private Sub mnuAlphaPer_Click()
    frmOptions.Show
End Sub


Private Sub mnuBajarVolumen_Click()
    Call frmMain.Form_KeyDown(&H28, 0)
    Call frmMain.Form_KeyDown(&H28, 0)
End Sub

Private Sub mnuChangeMask_Click()
    If bMiniMask = True Then
        Change_Mask False, True: frmMain.Image_State_Rep
    Else
        Change_Mask True, True: frmMain.Image_State_Rep
    End If
End Sub

Private Sub mnuCropTracks_Click()
    Dim i
    i = 0
    While (i <= frmPLST.clist.ItemCount - 1)
        If frmPLST.clist.bSelected(i) = False Then
            frmPLST.REMOVE_ITEM (i)
        Else
            i = i + 1
        End If
    Wend
    NormalSelection = -1
    frmPLST.Update_Plst_Scrollbar
    frmPLST.ReinitializeList
End Sub

Private Sub mnuCrossfade_Click()
    bCrossFadeEnabled = Not bCrossFadeEnabled
    frmMain.EnableCrossfade (bCrossFadeEnabled)
End Sub

Private Sub mnuDetener_Click()
    frmMain.Stop_Player
End Sub

Public Sub mnuDSP_Click()
    boolDspShow = Not boolDspShow
    mnuDSP.Checked = boolDspShow
    frmMain.Button(10).Selected = boolDspShow
    If boolDspShow Then
        frmDSP.Show
    Else
        frmDSP.Hide
    End If
    Call CheckMenu(9, mnuDSP.Checked)
End Sub


Private Sub mnuExploreItem_Click()
    Dim strPathExplore As String
    If frmPLST.clist.ItemCount = 0 Or NormalSelection = -1 Then Exit Sub
    strPathExplore = frmPLST.clist.exItem(NormalSelection)

    strPathExplore = Left(strPathExplore, InStrRev(strPathExplore, "\"))
    Shell "explorer.exe " & strPathExplore, vbMaximizedFocus
End Sub

Private Sub mnuFileinfo_Click()
    boolTagsShow = False    'next tageditor(toggle menu) event will makeit visible
    Call mnuTagEditor_Click
End Sub

Private Sub mnuForward5Sec_Click()
    frmMain.Five_Seg_Forward
End Sub

Private Sub mnuhelp_Click()
    Const flag As Long = SWP_NOMOVE Or SWP_NOSIZE
    SetWindowPos frmSplash.hwnd, HWND_TOPMOST, 0, 0, 0, 0, flag
    frmSplash.lblSplash(0).Caption = "Click me to exit"
    frmSplash.Show
End Sub

Private Sub mnuIDEInfo_Click()
    Call mnuTagEditor_Click
End Sub

Public Sub mnuIntro_Click()
    frmPopUp.mnuIntro.Checked = Not frmPopUp.mnuIntro.Checked
    frmMain.Intro
End Sub

Private Sub mnuInverSelection_Click()
    Dim i
    i = 0
    For i = 0 To frmPLST.clist.ItemCount - 1
        Call frmPLST.clist.ChangeSelection(i, Not frmPLST.clist.bSelected(i))
    Next
    'NormalSelection = -1
    'frmPLST.Update_Plst_Scrollbar
    frmPLST.ReinitializeList
End Sub

Private Sub mnuLibrary_Click()
    If boolMediaLibraryShow = False Then
        frmLibrary.Show
        boolMediaLibraryShow = True
    Else
        frmLibrary.Visible = False
        boolMediaLibraryShow = False
    End If
    mnuLibrary.Checked = boolMediaLibraryShow
End Sub

Public Sub mnuLoad_List_Click()
    On Error Resume Next
    Dim openPath As String
    Dim sfilter As String
    sfilter = "All Supported Playlist Files|*.npl;*.m3u;*.pls|" & "Playlist Files (*.npl)|*.npl|" & "Winamp Playlistfile (.m3u)|*.m3u|" & "MediaPlayer Pls file (.PLS)|*.pls|"
    openPath = clsDlg.GetOpenAsName(frmPLST.hwnd, "Open Playlist for NeoMp3 player", , sfilter, ALLOWMULTISELECT Or FILEMUSTEXIST Or LONGNAMES Or EXPLORER)
    If openPath = "" Then Exit Sub
    frmPLST.ClearList
    Call LoadPlaylist(openPath, True)
    frmPLST.ReinitializeList

End Sub

Private Sub mnuLoadPi_Click()
'show the file chooser

    Dim sFile As String, sfilter As String

    sfilter = "Orange Soda Visualization List|*.soda"
    sFile = clsDlg.GetOpenAsName(frmMain.hwnd, "Import Visualisation Plugins", , sfilter)
    If sFile = "" Then Exit Sub
    Dim T As TextStream, d As New Dictionary, s As String
    Set T = Fsys.OpenTextFile(App.Path & "\Sodas.list", ForReading, True)
    While Not T.AtEndOfStream
        s = T.readLine
        d.Add s, s
    Wend

    T.Close
    Set T = Fsys.OpenTextFile(App.Path & "\Sodas.list", ForAppending, False)
    If Not d.Exists(sFile) Then
        T.WriteLine sFile
        modPi.importXPIList sFile
    End If
    T.Close
    Set T = Nothing
End Sub

Private Sub mnuConfig_Click()
    On Error Resume Next    ' just in case ...
    oPlugIn.doConfig
End Sub



Public Sub mnuLyrics_Click()
End Sub

Public Sub mnuOrdenAleatorio_Click()
    Call frmMain.Button_Click(8)
End Sub

Private Sub mnuPlayTrack_Click()
    Dim i As Integer
    For i = 0 To frmPLST.clist.ItemCount - 1
        If frmPLST.clist.bSelected(i) = True Then    'play first selected track found
            sFileMainPlaying = frmPLST.clist.exItem(i)
            CurrentTrack_Index = i
            frmMain.PlayerIsPlaying = "true"
            frmPLST.ReinitializeList
            Stream_Stop lCurrentChannel
            frmPLST.Update_Plst_Scrollbar
            frmMain.Play
            Exit For    'exit loop since we have found the track
        End If
    Next
End Sub

Private Sub mnuRandomizeList_Click()
    frmMain.RANDOM_track
    Dim i As Long
    i = 0
    For i = 0 To frmPLST.clist.ItemCount - 1
        Call frmPLST.clist.SwapItems(i, CLng(Random_Order_track(i)))
        Call frmPLST.clist.ChangeSelection(i, False)
    Next
    CurrentTrack_Index = 0
    sFileMainPlaying = frmPLST.clist.exItem(0)
    Stream_Stop lCurrentChannel
    frmMain.Play
    frmPLST.ReinitializeList
End Sub

Private Sub mnuRemoveAll_Click()
    frmPLST.ClearList
End Sub

Private Sub mnuRemoveCorruptedFiles_Click()
    Dim i
    i = 0
    ' Loop should be in reverse ordersince deletion will affect the index of selected
    ' if items are deleted from top to  bottom
    While (i <= frmPLST.clist.ItemCount - 1)
        If Not FileExists(frmPLST.clist.exItem(i)) Or Trim(frmPLST.clist.exTracklength(i)) = "" Or Trim(frmPLST.clist.exTracklength(i)) = "00:00" Then
            frmPLST.REMOVE_ITEM (i)
        Else
            i = i + 1
        End If
    Wend
    'Set normal selection = -1 as the track may be deleted or reordered
    NormalSelection = -1
    'update scroll bar of playlist
    frmPLST.Update_Plst_Scrollbar
    'redraw Playlist
    frmPLST.ReinitializeList
End Sub

Private Sub mnuRemoveDuplicateFiles_Click()
    Dim i, j As Long

    With frmPLST.clist

        For i = 0 To .ItemCount - 1
            For j = .ItemCount To (i + 1) Step -1    'Reverse order deleton ensures that index of itemsto be checked is not changed
                If .exItem(j) = .exItem(i) Then
                    .RemoveItem j
                End If
            Next
        Next

    End With

    'Set normal selection = -1 as the track may be deleted or reordered
    NormalSelection = -1
    'update scroll bar of playlist
    frmPLST.Update_Plst_Scrollbar
    'redraw Playlist
    frmPLST.ReinitializeList

End Sub

Private Sub mnuRemoveMissingFiles_Click()
    Dim i
    i = 0
    ' Loop should be in reverse ordersince deletion will affect the index of selected
    ' if items are deleted from top to  bottom
    While (i <= frmPLST.clist.ItemCount - 1)
        If FileExists(frmPLST.clist.exItem(i)) <> True Then
            frmPLST.REMOVE_ITEM (i)
        Else
            i = i + 1
        End If
    Wend
    'Set normal selection = -1 as the track may be deleted or reordered
    NormalSelection = -1
    'update scroll bar of playlist
    frmPLST.Update_Plst_Scrollbar
    'redraw Playlist
    frmPLST.ReinitializeList
End Sub

Private Sub mnuRemoveSelected_Click()
    mnuRemoveTracks_Click
End Sub

Private Sub mnuRemoveTracks_Click()
    Dim i
    i = 0
    While (i <= frmPLST.clist.ItemCount - 1)
        If frmPLST.clist.bSelected(i) = True Then
            frmPLST.REMOVE_ITEM (i)
        Else
            i = i + 1
        End If
    Wend
    NormalSelection = -1
    frmPLST.Update_Plst_Scrollbar
    frmPLST.ReinitializeList
End Sub



Private Sub mnuReset_Click()
    On Error Resume Next

    Kill (App.Path & "\maheshmp3 player.ini")
    Kill (App.Path & "\mplayerlist.npl")
    bMiniMask = False
    frmMain.Visible = False
    frmPLST.Visible = False
    If boolVisShow Then FrmVisualisation.Visible = False
    If boolOptionsShow Then frmOptions.Visible = False
    If boolDspShow Then frmDSP.Visible = False

    Kill (tAppConfig.AppConfig & "Skins\" & tAppConfig.Skin & "\regions.dat")
    Load_Settings_INI (True)
    frmMain.Visible = True
    frmPLST.Visible = frmPopUp.mnuShowPlst.Checked
    frmPLST.ClearList
    sFileMainPlaying = ""

    frmMain.Stop_Player
    If boolVisShow Then FrmVisualisation.Visible = True: FrmVisualisation.PicSpectrum.cls
    If boolOptionsShow Then frmOptions.Visible = True
    If boolDspShow Then frmDSP.Visible = True
    'Change_Skin bMiniMask, True
End Sub

Private Sub mnuReverselist_Click()
    frmPLST.clist.Reverse
    frmPLST.ReinitializeList
End Sub

Private Sub mnuSelectAll_Click()
    Dim i
    i = 0
    For i = 0 To frmPLST.clist.ItemCount - 1
        Call frmPLST.clist.ChangeSelection(i, True)
    Next
    NormalSelection = 0
    'frmPLST.Update_Plst_Scrollbar
    frmPLST.ReinitializeList
End Sub

Private Sub mnuSelectNone_Click()
    Dim i
    i = 0
    For i = 0 To frmPLST.clist.ItemCount - 1
        Call frmPLST.clist.ChangeSelection(i, False)
    Next
    NormalSelection = -1
    'frmPLST.Update_Plst_Scrollbar
    frmPLST.ReinitializeList
End Sub

Public Sub mnuShowVis_Click()
    boolVisShow = Not boolVisShow
    mnuShowVis.Checked = boolVisShow
    Call CheckMenu(12, mnuShowVis.Checked)
    If boolVisShow = True Then
        FrmVisualisation.Show
    Else
        FrmVisualisation.Hide
    End If
End Sub

Private Sub mnuNew_List_Click()
    frmPLST.ClearList
End Sub

Private Sub mnuObjectName_Click(Index As Integer)
    If boolOpluginLoaded = False Then Exit Sub
    DoStop
    mnuConfig.Enabled = CBool(Split(mnuObjectName(Index).Tag, ",")(1))
    Dim i As Integer
    For i = 0 To mnuObjectName.Count - 1
        mnuObjectName(i).Checked = False
    Next i
    mnuObjectName(Index).Checked = True
    SaveSetting App.EXEName, "Visualization", "Object", Split(mnuObjectName(Index).Tag, ",")(0)
    SaveSetting App.EXEName, "Visualization", "Config", mnuConfig.Enabled
    SaveSetting App.EXEName, "Visualization", "Name", mnuObjectName(Index).Caption
    FrmVisualisation.setupVisualization
End Sub

Public Sub mnuPlay_playlist_Click()
    Call mnuLoad_List_Click
End Sub

Public Sub mnuPlayfile_Click()
    Call frmPLST.ADD_MULTIPLE_FILES(True)    'true for new list
End Sub

Public Sub mnuPlayfolder_Click()
    Call frmPLST.ADD_Dir(False, True)
End Sub

Public Sub mnuplayit_Click()
    frmMain.Play
End Sub

Public Sub mnuPreferences_Click()
    mnuPreferences.Checked = Not mnuPreferences.Checked
    frmOptions.Visible = mnuPreferences.Checked
    boolOptionsShow = frmOptions.Visible
    frmMain.Button(23).Selected = boolOptionsShow

End Sub
Private Sub mnuPausa_Click()
    frmMain.Pause_Play
End Sub

Public Sub mnuRegisterfiles_Click()
    Call RegisterType(".mp3", "MaheshMP3.File", "AUDIO FILES", "Mahesh Media File", 0)
    Call RegisterType(".npl", "MaheshMP3Playlist.File", "PLAYLIST FILES", "Mahesh Media  PlaylistFile", 1)
End Sub

Public Sub mnuRepeatTrack_Click()
    PlayerLoop = Not PlayerLoop
    frmPopUp.mnuRepeatTrack.Checked = PlayerLoop
    frmMain.Player_Repeat (PlayerLoop)
End Sub

Public Sub mnuExit_Click()
    frmMain.Button_Click (15)
End Sub

Private Sub mnuSearchDirectory_Click()
    boolSearchShow = True
    frmSearch.bAddtracktoPlaylist = True    'Default addition to playlist since it is call from frmmain
    frmSearch.Show
End Sub

Private Sub mnuSortByFilename_Click()
    frmPLST.clist.Sort (1)
    frmPLST.ReinitializeList
End Sub

Private Sub mnuSortbyList_Click()
    frmPLST.clist.Sort (2)
    frmPLST.ReinitializeList
End Sub

Private Sub mnuSortbyDuration_Click()
    frmPLST.clist.Sort (3)
    frmPLST.ReinitializeList
End Sub

Private Sub mnuSubirVolumen_Click()
    Call frmMain.Form_KeyDown(&H26, 0)
    Call frmMain.Form_KeyDown(&H26, 0)
End Sub
Private Sub mnuSave_List_Click()
    Dim savePath As String
    Dim sfilter As String
    sfilter = "Playlist Files (*.npl)|*.npl|" & "Winamp Playlistfile (.m3u)|*.m3u|" & "MediaPlayer Pls file (.PLS)|*.pls|"

    savePath = clsDlg.GetSaveAsName(frmPLST.hwnd, "Save Playlist as..", , sfilter, "*.npl")
    'frmPLST.GetSaveAsName(" ", App.Path, sFilter)
    If savePath = "" Then Exit Sub
    frmPLST.SavePlaylist (savePath)
End Sub

Private Sub mnuShowEqualizer_Click()
    Call frmMain.Button_Click(24)
    Call CheckMenu(10, mnuShowEqualizer.Checked)
End Sub

Public Sub mnuShowPlst_Click()
    mnuShowPlst.Checked = Not mnuShowPlst.Checked
    frmMain.Button(26).Selected = Not frmMain.Button(26).Selected
    frmPLST.Visible = frmMain.Button(26).Selected
    Call CheckMenu(11, mnuShowPlst.Checked)
End Sub

Private Sub mnuSigTrack_Click()
    frmMain.Next_Track
End Sub

Public Sub mnuSilencio_Click()
    frmPopUp.mnuSilencio.Checked = Not frmPopUp.mnuSilencio.Checked
    frmMain.Player_Mute
End Sub

Public Sub mnuSkinsAdd_Click(Index As Integer)
    On Error Resume Next
    Dim Skins As String, sRoot As String
    Dim i As Integer
    Skins = Trim(mnuSkinsAdd(Index).Caption)
    '// si es el mismo skin salir
    If Skins = "" Then Exit Sub

    '// chech the existance of skin
    sRoot = tAppConfig.AppConfig & "Skins\"
    If Dir(sRoot & Skins, vbDirectory) = "" Then Exit Sub

    If LCase(Skins) = LCase(tAppConfig.Skin) Then Exit Sub
    '// add skin to menu
    For i = 1 To mnuSkinsAdd.Count
        If i = Index Then
            mnuSkinsAdd(Index).Checked = True
        Else
            mnuSkinsAdd(i).Checked = False
        End If
    Next i
    bLoadingSkin = True
    frmSplash.Show
    frmPLST.Visible = False
    frmMain.Visible = False
    If boolOptionsLoaded And boolOptionsShow Then frmOptions.Visible = False
    If boolVisShow Then FrmVisualisation.Visible = False
    If boolTagsLoaded And boolTagsShow Then frmTags.Visible = False
    If boolDspShow And boolDspShow Then frmDSP.Visible = False
    tAppConfig.Skin = Skins
    Read_Config_Skin
    Form_Mini_Normal

    Change_Mask bMiniMask, False

    Change_Skin Skins
    '// ajust mode

    If boolOptionsLoaded Then
        frmOptions.lblskin.Caption = "Current Skin: " & Skins
        frmOptions.Listaskins.Selected(Index) = True
        frmOptions.Listaskins.ListIndex = Index - 1
        frmOptions.Visible = boolOptionsShow
    End If
    frmMain.picNormalMode.Refresh


    frmMain.Show
    frmSplash.Hide


    If frmMain.Button(26).Selected Then frmPLST.Form_Resize1: frmPLST.Visible = True: frmPLST.ReinitializeList
    If boolOptionsLoaded And boolOptionsShow Then frmOptions.Visible = True
    If boolVisShow Then FrmVisualisation.Visible = True
    If boolTagsLoaded And boolTagsShow Then frmTags.Visible = True
    If boolDspShow And boolDspShow Then frmDSP.Visible = True
    bLoadingSkin = False

    For i = 0 To 9 Step 1
        frmMain.Eq_SliderCtrl(i).Value = frmMain.Eq_SliderCtrl(i).Value
    Next
End Sub

Public Sub mnuSpecBars_Click()
    With frmPopUp
        .mnuSpecBars.Checked = True
        .mnuSpecNone.Checked = False
        .mnuSpecOsc.Checked = False
    End With
End Sub

Public Sub mnuSpecNone_Click()
    With frmPopUp
        .mnuSpecBars.Checked = False
        .mnuSpecNone.Checked = True
        .mnuSpecOsc.Checked = False
    End With
End Sub

Public Sub mnuSpecOsc_Click()
    With frmPopUp
        .mnuSpecBars.Checked = False
        .mnuSpecNone.Checked = False
        .mnuSpecOsc.Checked = True
    End With
End Sub

Public Sub mnuTagEditor_Click()
'On Error Resume Next
    boolTagsShow = Not boolTagsShow
    mnuTagEditor.Checked = boolTagsShow
    If boolTagsShow = True Then
        frmTags.Show
        If cGetInputState() <> 0 Then DoEvents
        'frmTags.fileTags.Clear
        frmTags.vkFiletags.Clear
        frmTags.listRef.ListItems.Clear
        'don't let  vkfiletags draw itself repeatedly which it does on adding an item to it
        frmTags.vkFiletags.UnRefreshControl = True
        Dim i
        For i = 0 To frmPLST.clist.ItemCount - 1
            If frmPLST.clist.bSelected(i) = True Then frmTags.Load_Tags frmPLST.clist.exItem(i)    'CurrentTrack_Index)
        Next

        If frmTags.vkFiletags.ListCount = 0 Then Exit Sub
        'Show tags of first item
        frmTags.Show_tags (1)
        'Show first item as selected
        frmTags.vkFiletags.Selected(1) = True    'index starts from 1 in vkListbox instead of zero
        'Now we can draw vkfiletags listbox
        frmTags.vkFiletags.UnRefreshControl = False
        frmTags.vkFiletags.Refresh

    Else
        frmTags.Visible = False
    End If
End Sub

Private Sub mnuTrackAnterior_Click()
    frmMain.Previous_Track
End Sub

Private Sub vkSysTrayIcon_MouseDblClick(Index As Integer, Button As MouseButtonConstants, ID As Long)
    If ID <> 0 Then Exit Sub
    If frmMain.WindowState = vbMinimized Or frmMain.Visible = False Then
        frmMain.WindowState = vbNormal
        frmMain.Visible = True
    Else
        frmMain.Button_Click (13)    'minimise
    End If
End Sub

Private Sub vkSysTrayIcon_MouseDown(Index As Integer, Button As MouseButtonConstants, ID As Long)
    If Index = 0 Then Exit Sub

    frmMain.Button_Click (Index - 1)

    Select Case Index
    Case 1:

    Case 2:

    Case 3:

    Case 4:

    Case 5:

    End Select
End Sub

Private Sub vkSysTrayIcon_MouseUp(Index As Integer, Button As MouseButtonConstants, ID As Long)
    If Button = vbRightButton And Index = 0 Then Me.PopupMenu frmPopUp.mnuMenuPrincipal

End Sub
