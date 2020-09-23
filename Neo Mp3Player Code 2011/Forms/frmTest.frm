VERSION 5.00
Begin VB.Form frmSearch 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FF0000&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   4950
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8805
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmTest.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   330
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   587
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin MMPlayerXProject.vkFrame frameContainer 
      Height          =   4155
      Left            =   510
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   390
      Width           =   7710
      _extentx        =   13600
      _extenty        =   7329
      backcolor1      =   12632256
      backcolor2      =   8421504
      font            =   "frmTest.frx":000C
      showtitle       =   0
      titlecolor1     =   12632256
      bordercolor     =   0
      roundangle      =   0
      Begin MMPlayerXProject.ucProgressBar pbProgress 
         Height          =   255
         Left            =   240
         TabIndex        =   10
         Top             =   3240
         Visible         =   0   'False
         Width           =   7215
         _extentx        =   12726
         _extenty        =   450
         font            =   "frmTest.frx":0034
         brushstyle      =   0
         color           =   33023
         color2          =   12648384
         value           =   20
      End
      Begin MMPlayerXProject.vkCommand CMD 
         Height          =   375
         Index           =   2
         Left            =   6180
         TabIndex        =   9
         Top             =   3690
         Width           =   1395
         _extentx        =   2461
         _extenty        =   661
         backcolor1      =   12632256
         backcolor2      =   4210752
         backcolorpushed2=   12632256
         caption         =   "Add Files"
         font            =   "frmTest.frx":0060
         forecolor       =   0
         enabled         =   0
         bordercolor     =   12632256
         picture         =   "frmTest.frx":0088
         mousehoverpicture=   "frmTest.frx":0422
         graypicturewhendisabled=   0
         disabledbackcolor=   8421504
         customstyle     =   0
      End
      Begin MMPlayerXProject.vkCommand CMD 
         Height          =   375
         Index           =   1
         Left            =   4380
         TabIndex        =   8
         Top             =   3690
         Width           =   1695
         _extentx        =   2990
         _extenty        =   661
         backcolor1      =   12632256
         backcolor2      =   4210752
         caption         =   "Start Search"
         font            =   "frmTest.frx":07BC
         forecolor       =   0
         bordercolor     =   12632256
         picture         =   "frmTest.frx":07E4
         mousehoverpicture=   "frmTest.frx":11F6
         customstyle     =   0
      End
      Begin VB.TextBox txtDirPath 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   120
         TabIndex        =   5
         Text            =   "C:\"
         Top             =   240
         Width           =   6915
      End
      Begin MMPlayerXProject.vkCommand CMD 
         Height          =   345
         Index           =   0
         Left            =   7080
         TabIndex        =   4
         Top             =   210
         Width           =   495
         _extentx        =   873
         _extenty        =   609
         backcolor2      =   4210752
         backcolorpushed2=   4210752
         caption         =   ""
         font            =   "frmTest.frx":1C08
         picture         =   "frmTest.frx":1C30
         mousehoverpicture=   "frmTest.frx":2642
         pictureoffsetx  =   120
         pictureoffsety  =   -20
         graypicturewhendisabled=   0
         drawfocus       =   0
         drawmouseinrect =   0
         customstyle     =   0
      End
      Begin MMPlayerXProject.vkFrame frameBack 
         Height          =   1815
         Left            =   120
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   630
         Width           =   7440
         _extentx        =   13123
         _extenty        =   3201
         backcolor1      =   12632256
         backcolor2      =   8421504
         font            =   "frmTest.frx":3054
         showtitle       =   0
         titlecolor1     =   12632256
         bordercolor     =   0
         roundangle      =   0
         Begin MMPlayerXProject.vkCheck chkUsefile 
            CausesValidation=   0   'False
            Height          =   225
            Index           =   4
            Left            =   180
            TabIndex        =   18
            Top             =   1380
            Width           =   3030
            _extentx        =   5345
            _extenty        =   397
            backstyle       =   0
            caption         =   "Add tracks to Current Playlist"
            font            =   "frmTest.frx":307C
         End
         Begin VB.TextBox txtSkipsize 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   0  'None
            BeginProperty DataFormat 
               Type            =   0
               Format          =   "12345678"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   0
            EndProperty
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   2790
            MaxLength       =   7
            TabIndex        =   17
            Text            =   "100"
            Top             =   990
            Width           =   765
         End
         Begin MMPlayerXProject.vkCheck chkUsefile 
            CausesValidation=   0   'False
            Height          =   225
            Index           =   3
            Left            =   180
            TabIndex        =   16
            Top             =   990
            Width           =   3945
            _extentx        =   6959
            _extenty        =   397
            backstyle       =   0
            caption         =   "Skip files of size less than               KB"
            font            =   "frmTest.frx":30A4
         End
         Begin MMPlayerXProject.vkCheck chkUsefile 
            CausesValidation=   0   'False
            Height          =   225
            Index           =   1
            Left            =   2460
            TabIndex        =   14
            Top             =   570
            Width           =   2115
            _extentx        =   3731
            _extenty        =   397
            backstyle       =   0
            caption         =   "include system files"
            font            =   "frmTest.frx":30CC
         End
         Begin MMPlayerXProject.vkCheck chkUsefile 
            CausesValidation=   0   'False
            Height          =   225
            Index           =   0
            Left            =   180
            TabIndex        =   13
            Top             =   570
            Width           =   2085
            _extentx        =   3678
            _extenty        =   397
            backstyle       =   0
            caption         =   "include hidden files"
            font            =   "frmTest.frx":30F4
         End
         Begin MMPlayerXProject.vkCheck chkUsefile 
            CausesValidation=   0   'False
            Height          =   225
            Index           =   2
            Left            =   4740
            TabIndex        =   12
            Top             =   570
            Width           =   4845
            _extentx        =   8546
            _extenty        =   397
            backstyle       =   0
            caption         =   "Recursive Search"
            font            =   "frmTest.frx":311C
            value           =   1
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   $"frmTest.frx":3144
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   390
            Index           =   7
            Left            =   180
            TabIndex        =   15
            Top             =   60
            Width           =   7245
            WordWrap        =   -1  'True
         End
      End
      Begin VB.Label lblSearchDir 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Search in Directory......"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   120
         TabIndex        =   6
         Top             =   30
         Width           =   2025
      End
      Begin VB.Label lblDirInfo 
         Appearance      =   0  'Flat
         BackStyle       =   0  'Transparent
         Caption         =   "Searches all supported media files in the specified directory and adds them to the Library"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   675
         Left            =   270
         TabIndex        =   3
         Top             =   2790
         Width           =   7215
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblFileCount 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "FILES FOUND [-]  "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   180
         TabIndex        =   2
         Top             =   3660
         Visible         =   0   'False
         Width           =   1410
      End
      Begin VB.Label lblinfo 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Looking in:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   150
         TabIndex        =   1
         Top             =   2520
         Visible         =   0   'False
         Width           =   945
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H00400000&
         Height          =   855
         Left            =   120
         Top             =   2730
         Width           =   7425
      End
   End
   Begin MMPlayerXProject.Button btnExit 
      Height          =   120
      Left            =   0
      TabIndex        =   7
      Top             =   0
      Width           =   225
      _extentx        =   397
      _extenty        =   212
      style           =   1
      buttoncolor     =   12632256
      mousepointer    =   99
   End
End
Attribute VB_Name = "frmSearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim cWindows As New cWindowSkin   'Initialize Class for Skinning
Dim iCount As Integer             'number of tracks searched
Dim sTitleSearch() As String      'Array to store title name to be searched
Dim sArtistSearch() As String     'Array to store Artist name to be searched
Dim sAlbumSearch() As String      'Array to store album name to be searched
Dim iYear As Integer              'year from IDE
Dim sGenre As String              'genre to be searched
Dim bMatchCaseTitle As Boolean    'whether search should be case sensitive or not when searching for title
Dim bMatchCaseArtist As Boolean   'whether search should be case sensitive or not when searching for artist name
Dim bMatchCaseAlbum As Boolean    'whether search should be case sensitive or not when searching for album
Dim iRating() As Integer          'array that reflects number of matches for track corresponding to its index returned in search result
Dim bHidden As Boolean
Dim bSystem As Boolean
Public bAddtracktoPlaylist As Boolean
Dim lSkipSize As Long

Dim rst As ADODB.Recordset
Dim cnn As ADODB.connection
Dim strSQL As String
'Dim cFile As New cMP3 'declared in addfiletolib function

Dim bAddingTracks As Boolean

Private Enum FileAttributes
    Alias = 1024
    Archive = 32
    Compressed = 2048
    Directory = 16
    Hidden = 2
    Normal = 0
    ReadOnly = 1
    System = 4
    Volume = 8
    All = Alias Or Archive Or Compressed Or Hidden Or ReadOnly Or System
End Enum

Private dummyClist As New clist


'"""""""""""""FUNCTION=cRead_INI"""""""""""""
'"""""""""""Returns color from INI in RGB()"""""""
Private Sub btnExit_Click()
    Unload Me
End Sub

'"""""""""""""""" SUB:=addFilesFromDir"""""""""""
'"""""""""""""""""""adds files with specified pattern located in given directory""""
'"""ARGUMMENTS
'-----------------------
'path=> dir to be searched
'subfolder=> allows recursive search
'--------------------------------------
Public Sub addFilesfromDir(Path As String, Optional SubFolder As Boolean = True)
'Dim li As ListItem
    On Error Resume Next:
    Dim WFD As WIN32_FIND_DATA
    Dim hFile As Long, fPath As String, fName As String
    If bSearching = False Then Exit Sub
    lblDirInfo.Caption = Path
    fPath = AddBackSlash(Path)
    fName = fPath & "*.*"

    If bSearching = False Then Exit Sub
    'if hfile is +ve means file exists
    hFile = FindFirstFile(fName, WFD)
    If (hFile > 0) And ((WFD.dwFileAttributes And FILE_ATTRIBUTE_DIRECTORY) <> FILE_ATTRIBUTE_DIRECTORY) Then
        If ((WFD.dwFileAttributes And FILE_ATTRIBUTE_DIRECTORY) <> FILE_ATTRIBUTE_DIRECTORY) Then
            If ((WFD.dwFileAttributes And Hidden) <> Hidden Or bHidden) And _
               isMediaFile(fPath & StripNulls(WFD.cFileName)) And _
               ((WFD.dwFileAttributes And System) <> System Or bSystem) Then    ' if file is mp3 file
                If (lSkipSize = 0) Or ((lSkipSize > 0) And FileLen(fPath & StripNulls(WFD.cFileName)) > lSkipSize * 1024) Then
                    dummyClist.AddItem fPath & StripNulls(WFD.cFileName), ""
                    iCount = iCount + 1: lblFileCount.Caption = "FILES ADDED [" + str(iCount) + "]"
                End If
            End If
        End If
    End If

    While FindNextFile(hFile, WFD)
        If ((WFD.dwFileAttributes And FILE_ATTRIBUTE_DIRECTORY) <> FILE_ATTRIBUTE_DIRECTORY) Then
            If ((WFD.dwFileAttributes And Hidden) <> Hidden Or bHidden) And _
               isMediaFile(fPath & StripNulls(WFD.cFileName)) And _
               ((WFD.dwFileAttributes And System) <> System Or bSystem) Then    ' if file is mp3 file
                If (lSkipSize = 0) Or ((lSkipSize > 0) And FileLen(fPath & StripNulls(WFD.cFileName)) > lSkipSize * 1024) Then
                    dummyClist.AddItem fPath & StripNulls(WFD.cFileName), ""
                    iCount = iCount + 1: lblFileCount.Caption = "FILES ADDED [" + str(iCount) + "]"
                End If
            End If
        End If
    Wend

    If SubFolder Then
        hFile = FindFirstFile(fName, WFD)
        If (hFile > 0) And ((WFD.dwFileAttributes And FILE_ATTRIBUTE_DIRECTORY) = FILE_ATTRIBUTE_DIRECTORY) And _
           StripNulls(WFD.cFileName) <> "." And StripNulls(WFD.cFileName) <> ".." Then
            'if recursion is true
            addFilesfromDir fPath & StripNulls(WFD.cFileName), True
            DoEvents
        End If

        While FindNextFile(hFile, WFD)
            If ((WFD.dwFileAttributes And FILE_ATTRIBUTE_DIRECTORY) = FILE_ATTRIBUTE_DIRECTORY) And _
               StripNulls(WFD.cFileName) <> "." And StripNulls(WFD.cFileName) <> ".." Then
                'give system some breath
                DoEvents
                addFilesfromDir fPath & StripNulls(WFD.cFileName), True
            End If
        Wend

    End If
SkipSearch:
    FindClose hFile
    'Exit Sub
    'errorhandler:
    'FindClose hFile
    'MsgBox err.Description & "AddfilefromDirectory"
End Sub




Private Sub chkUsefile_Change(Index As Integer, Value As CheckBoxConstants)
    Select Case Index
    Case 0: bHidden = CBool(Value)
    Case 1: bSystem = CBool(Value)
        'Case 2: bsubfolder = CBool(Value)
    Case 3: lSkipSize = 0
    Case 4: bAddtracktoPlaylist = CBool(Value)
    End Select
End Sub

Private Sub CMD_Click(Index As Integer)
    Select Case Index
    Case 0:
        Dim txtReturn As String
        Dim Flags As Long
        Dim txtInstruction As String
        txtInstruction = "Add Directory for NeoMp3 player"
        'SpecialFolder = 0
        'StartFolder = "d:\c backup\manvendra\pankaj udas"
        txtInstruction = "Add Folders to Search  for Media files "
        'Flags = Flags + BIF_USENEWUI
        'Flags = Flags + BIF_EDITBOX
        'Flags = Flags + BIF_STATUSTEXT
        'flags = flags + BIF_NEWDIALOGSTYLE
        'Flags = Flags + BIF_BROWSEINCLUDEFILES
        txtReturn = clsDlg.FolderBrowse(Me.hwnd, txtInstruction, Flags)
        If txtReturn <> "" Then txtDirPath.Text = txtReturn

    Case 1:
        If CMD(1).Caption = "Start Search" And bSearching Then Exit Sub
        If CMD(1).Caption = "Start Search" Then
            If Not FileExists(txtDirPath.Text) Then    'Dir("")does't return zero, it returns 04
                MsgBox "Please select valid directory"
                Exit Sub
            End If
            CMD(1).Caption = "Stop Search"
            frameBack.Enabled = False
            lblinfo.Visible = True
            lblFileCount.Visible = True
            iCount = 0
            Call PrepareSearch
            dummyClist.Clear
            bSearching = True
            Call addFilesfromDir(txtDirPath.Text, CBool(chkUsefile(2).Value))    'CBool(chkUsefile(2).Value))
            Call showSearchResult
            CMD(1).Caption = "Start Search"
            frameBack.Enabled = True
            'Call Sort
            Exit Sub
        ElseIf CMD(1).Caption = "Stop Search" Then
            bSearching = False
            CMD(1).Caption = "Start Search"
            frameBack.Enabled = False
            Call showSearchResult
        End If

    Case 2:
        If bSearching Then Exit Sub
        bSearching = True
        lblDirInfo.Caption = ""
        lblinfo.Caption = "Adding Files"
        lblinfo.Visible = True
        bAddingTracks = True
        AddfilestoLib
        bAddingTracks = False
        lblinfo.Caption = "File addition Finished!"
        pbProgress.Visible = False
        CMD(1).Caption = "Start Search"
        lblDirInfo.Caption = " [ " & iCount & " ] tracks are added to the library. "
        CMD(2).Enabled = False
        bSearching = False
        If boolMediaLibraryShow = True Then frmLibrary.mnurefresh_Click
        frmPLST.ReinitializeList
    End Select

End Sub



'32 bit - Add 4 to counter
'24 bit - Add 3 to counter
Private Sub Form_Load()
    Const flag As Long = SWP_NOMOVE Or SWP_NOSIZE
    boolSearchShow = True
    If OpcionesMusic.SiempreTop = True Then
        SetWindowPos frmSearch.hwnd, HWND_TOPMOST, 0, 0, 0, 0, flag
    Else
        SetWindowPos frmSearch.hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, flag
    End If

    LoadSkin
    LoadConfig

    Set rst = New ADODB.Recordset
    Set cnn = New ADODB.connection
    With cnn
        .Provider = "Microsoft.Jet.OLEDB.4.0"
        .Properties("Data Source") = tAppConfig.AppConfig & "Library\music.mdb"
        '.Properties("Jet OLEDB:Database Password") = "Licenciao159"
        .CursorLocation = adUseClient
        .Open
    End With

    strSQL = "SELECT * FROM Music"

    rst.Open strSQL, cnn, adOpenDynamic, adLockOptimistic

End Sub
Public Sub LoadSkin()
'On Error Resume Next
'Me.Visible = False
    Me.Height = Read_INI("FORM", "formheight", 6020, , , tAppConfig.AppConfig & "Skins\" & tAppConfig.Skin & "\SEARCH\config.ini")
    Me.Width = Read_INI("FORM", "formwidth", 8888, , , tAppConfig.AppConfig & "Skins\" & tAppConfig.Skin & "\SEARCH\config.ini")

    Set cWindows.FormularioPadre = Me
    cWindows.ColorInvisible = Read_INI("NormalMode", "ColorTrans", RGB(255, 0, 255), True)
    cWindows.MinimoAlto = Read_INI("Configuration", "MinHeight", 10, False, False, tAppConfig.AppConfig & "Skins\" & tAppConfig.Skin & "\SEARCH\config.ini")
    cWindows.MinimoAncho = Read_INI("Configuration", "MinWidth", 10, False, False, tAppConfig.AppConfig & "Skins\" & tAppConfig.Skin & "\SEARCH\config.ini")
    'iButtonsLeft = Read_INI("Configuration", "ButtonsLeft", 5, False, False, tAppConfig.AppConfig & "Skins\" & tAppConfig.Skin & "\PLAY_LIST\config.ini")
    'iButtonsTop = Read_INI("Configuration", "ButtonsTop", 10, False, False, tAppConfig.AppConfig & "Skins\" & tAppConfig.Skin & "\PLAY_LIST\config.ini")
    cWindows.CargarSkin tAppConfig.AppConfig & "Skins\" & tAppConfig.Skin & "\SEARCH\", True
    Dim k
    k = Read_Config_Button(btnExit, "Configuration", "exitButton", "0,0,10,10", tAppConfig.AppConfig & "Skins\" & tAppConfig.Skin & "\SEARCH\config.ini")
    Set btnExit.PictureNormal = LoadPicture(tAppConfig.AppConfig & "Skins\" & tAppConfig.Skin & "\SEARCH\Exitnormal.bmp")
    Set btnExit.PictureOver = LoadPicture(tAppConfig.AppConfig & "Skins\" & tAppConfig.Skin & "\SEARCH\Exitover.bmp")
    Set btnExit.PictureDown = LoadPicture(tAppConfig.AppConfig & "Skins\" & tAppConfig.Skin & "\SEARCH\ExitDown.bmp")

End Sub
Private Sub LoadConfig()
    Dim i As Integer

    frameContainer.Top = Read_INI("CONTAINER", "top", 12, , , tAppConfig.AppConfig & "Skins\" & tAppConfig.Skin & "\SEARCH\config.ini")
    frameContainer.Left = Read_INI("CONTAINER", "left", 12, , , tAppConfig.AppConfig & "Skins\" & tAppConfig.Skin & "\SEARCH\config.ini")
    frameContainer.BackColor1 = cRead_INI("CONTAINER", "Backcolor1", RGB(70, 100, 150), tAppConfig.AppConfig & "Skins\" & tAppConfig.Skin & "\SEARCH\config.ini")
    frameContainer.BackColor2 = cRead_INI("CONTAINER", "Backcolor2", RGB(70, 100, 150), tAppConfig.AppConfig & "Skins\" & tAppConfig.Skin & "\SEARCH\config.ini")
    frameBack.BackColor1 = frameContainer.BackColor1
    frameBack.BackColor2 = frameContainer.BackColor1    'GetGradColor(100, 70, frameContainer.BackColor1, frameContainer.BackColor1, frameContainer.BackColor2)

    If bAddtracktoPlaylist Then chkUsefile(4).Value = vbChecked

    For i = 0 To 4
        chkUsefile(i).ForeColor = cRead_INI("CHECKBOX", "forecolor", RGB(70, 100, 150), tAppConfig.AppConfig & "Skins\" & tAppConfig.Skin & "\SEARCH\config.ini")
        chkUsefile(i).BackColor = frameContainer.BackColor1
    Next


    lblFileCount.ForeColor = cRead_INI("LABEL", "forecolor", RGB(70, 100, 150), tAppConfig.AppConfig & "Skins\" & tAppConfig.Skin & "\SEARCH\config.ini")
    lblDirInfo.ForeColor = lblFileCount.ForeColor
    lblinfo.ForeColor = lblFileCount.ForeColor
    lblSearchDir.ForeColor = lblFileCount.ForeColor

    txtDirPath.BackColor = cRead_INI("TEXT", "backcolor", RGB(70, 100, 150), tAppConfig.AppConfig & "Skins\" & tAppConfig.Skin & "\SEARCH\config.ini")
    txtSkipsize.BackColor = frameContainer.BackColor1
    txtSkipsize.ForeColor = chkUsefile(0).ForeColor

    For i = 0 To 2
        CMD(i).UnRefreshControl = True
        CMD(i).BackColor1 = cRead_INI("BUTTON", "Backcolor1", RGB(70, 100, 150), tAppConfig.AppConfig & "Skins\" & tAppConfig.Skin & "\SEARCH\config.ini")
        CMD(i).BackColor2 = cRead_INI("BUTTON", "Backcolor2", RGB(70, 100, 150), tAppConfig.AppConfig & "Skins\" & tAppConfig.Skin & "\SEARCH\config.ini")
        CMD(i).DisabledBackColor = cRead_INI("BUTTON", "Backcolor2", RGB(70, 100, 150), tAppConfig.AppConfig & "Skins\" & tAppConfig.Skin & "\SEARCH\config.ini")
        CMD(i).ForeColor = cRead_INI("BUTTON", "forecolor", RGB(70, 100, 150), tAppConfig.AppConfig & "Skins\" & tAppConfig.Skin & "\SEARCH\config.ini")
        CMD(i).BorderColor = cRead_INI("BUTTON", "bordercolor", RGB(70, 100, 150), tAppConfig.AppConfig & "Skins\" & tAppConfig.Skin & "\SEARCH\config.ini")
        CMD(i).BackColorPushed1 = cRead_INI("BUTTON", "PushedColor1", RGB(70, 100, 150), tAppConfig.AppConfig & "Skins\" & tAppConfig.Skin & "\SEARCH\config.ini")
        CMD(i).BackColorPushed2 = cRead_INI("BUTTON", "PushedColor2", RGB(70, 100, 150), tAppConfig.AppConfig & "Skins\" & tAppConfig.Skin & "\SEARCH\config.ini")
        CMD(i).UnRefreshControl = False
        CMD(i).Refresh
    Next

    pbProgress.Color = CMD(1).BackColor2
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ReleaseCapture
    SendMessage Me.hwnd, 161, 2, 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
    bSearching = False
    bAddingTracks = False
    DoEvents
    boolSearchShow = False
    cnn.Close
    Set cnn = Nothing
End Sub

Public Sub Search_tags(sTrackpath As String)
    Dim sSearchString As String
    Dim i As Integer
    Dim iRate As Integer    'more the rate better is the search
    Dim cfile As New cMP3
    cfile.Read_MPEGInfo = True
    cfile.Read_File_Tags sTrackpath

    ' load tags
    sSearchString = GetFileTitle(sTrackpath)
    'Search in title
    If Not bMatchCaseTitle Then sSearchString = UCase(sSearchString)
    For i = 0 To UBound(sTitleSearch)
        If Trim(sTitleSearch(i)) <> "" Then
            If CBool(InStr(1, sSearchString, sTitleSearch(i), vbBinaryCompare)) Then
                iRate = iRate + 1
            End If
        End If
    Next

    'Search in artist
    If Not bMatchCaseArtist Then sSearchString = UCase(Trim(cfile.Artist))
    For i = 0 To UBound(sArtistSearch)
        If Trim(sArtistSearch(i)) <> "" Then
            If CBool(InStr(1, sSearchString, sArtistSearch(i), vbBinaryCompare)) Then
                iRate = iRate + 1
            End If
        End If
    Next

    'Search for Albums
    If Not bMatchCaseAlbum Then sSearchString = UCase(Trim(cfile.Album))
    For i = 0 To UBound(sAlbumSearch)
        If Trim(sAlbumSearch(i)) <> "" Then
            If CBool(InStr(1, sSearchString, sAlbumSearch(i), vbBinaryCompare)) Then
                iRate = iRate + 1
            End If
        End If
    Next

    'Search Year
    If iYear <> 0 And iYear = Val(cfile.Year) Then
        iRate = iRate + 1
    End If

    'Search Genre
    If UCase(sGenre) <> "" And UCase(sGenre) = UCase(cfile.Genre) Then
        iRate = iRate + 1
    End If


    If iRate > 0 Then
        If iCount > 0 Then ReDim Preserve iRating(LBound(iRating) To UBound(iRating) + 1)
        iRating(UBound(iRating)) = iRate
        iCount = iCount + 1: lblFileCount.Caption = "FILES ADDED [" + str(iCount) + "]"
        frmPLST.Add_track_to_Playlist (sTrackpath)
    End If
End Sub



'---------------------------------------------------------------------------------------
' Procedure : Sort
' Author    : Mahesh Kurmi
' Date      : 3/5/2011
' Purpose   : to sort the data searched according to number of search parameters satisfied
' HOW       :
' Returns   :
' I will be using it later on
'---------------------------------------------------------------------------------------
'
Private Sub Sort()
'The fastets sort algorithm!
    Dim sVal1 As Integer, sVal2 As Integer
    Dim Row As Long
    Dim MaxRow As Long
    Dim MinRow As Long
    Dim Swtch As Long
    Dim Limit As Long
    Dim Offset As Long

    MaxRow = UBound(iRating)
    MinRow = LBound(iRating)
    Offset = MaxRow \ 2

    Do While Offset > 0
        Limit = MaxRow - Offset
        Do
            Swtch = False         ' Assume no switches at this offset.

            ' Compare elements and switch ones out of order:

            For Row = MinRow To Limit
                sVal1 = iRating(Row)
                sVal2 = iRating(Row + Offset)
                ''Debug.Print str(iRating(Row)) + " " + str(iRating(Row + Offset))
                If sVal1 < sVal2 Then
                    Call frmPLST.clist.SwapItems(Row, (Row + Offset))
                    Call intSwap(iRating(Row), iRating(Row + Offset))
                    'Debug.Print str(iRating(Row)) + " " + str(iRating(Row + Offset))
                    Swtch = Row
                End If
            Next Row

            ' Sort on next pass only to where last switch was made:
            Limit = Swtch - Offset
        Loop While Swtch

        ' No switches at last offset, try one half as big:
        Offset = Offset \ 2
    Loop
End Sub
Private Sub PrepareSearch()
    ReDim iRating(0 To 0)
    iCount = 0
    lblFileCount.Caption = "FILES ADDED [" + str(iCount) + "]"
    bHidden = CBool(chkUsefile(0).Value)
    bSystem = CBool(chkUsefile(1).Value)
    lSkipSize = IIf(CBool(chkUsefile(3).Value), Val(txtSkipsize.Text), 0)
End Sub

Private Sub intSwap(ByRef var1 As Integer, ByRef var2 As Integer)
    Dim X As Integer
    X = var1
    var1 = var2
    var2 = X
End Sub


Public Sub AddfilestoLib()
    Dim iProg As Integer

    ' On Error Resume Next
    Dim i As Integer, j As Integer
    Dim lTackCount As Long
    pbProgress.Min = 0
    pbProgress.Max = iCount
    pbProgress.Value = 0
    pbProgress.Visible = True

    bAddtracktoPlaylist = CBool(chkUsefile(4).Value)

    For i = 0 To dummyClist.ItemCount - 1
        If bAddingTracks = False Then Exit For
        If AddTrack(dummyClist.Item(i)) Then lTackCount = lTackCount + 1: lblDirInfo.Caption = " [ " & lTackCount & " ] tracks are added to the library. "
        lblDirInfo.Caption = "Adding File" & dummyClist.Item(i)
        iProg = iProg + 1
        pbProgress.Value = iProg
        DoEvents
    Next i

    lblFileCount.Caption = "Ready: [ " & lTackCount & " ] tracks Added to the Database/playlist"
    If boolMediaLibraryShow = True Then frmLibrary.mnurefresh_Click
    Exit Sub

hell:

End Sub

Public Sub showSearchResult()

'frmPLST.Update_Plst_Scrollbar
    bSearching = False
    If iCount >= 1 Then
        lblDirInfo.Caption = "The Specified directory has beean searched. " + str(iCount) + " Files are found! To add files to library click Add files"
        CMD(2).Enabled = True
        bSearching = False
    Else
        lblDirInfo.Caption = "The Specified directory has beean searched." + " Sorry no File is found"
        CMD(2).Enabled = False
    End If
End Sub

Public Function AddTrack(sPath As String) As Boolean
'PURPOSE: This function has to be defined separately to add one file at a time to library
'because cMP3 class uses pointers to determine mpeg info. So if cMp3 object if not terminated then
'it gives previous values if new data is not overwritten,eg. TRACK1 has duration 4:34 and track 2 gives error
'and hence doesnt get duration value so Track2's value as memory adress at that of prev one and hence
'its duration is shown as 4:34
'This is because(as i think) in VB the same procedure allocates same memory address for a paricular variable declaraion using dim
'so we have to skip the subroutine every time.

    On Error GoTo errorHandler:
    Dim FS As New FileSystemObject
    Dim bCDROM As Boolean
    Dim dDrive As drive
    Dim sTitle As String, sArtist As String, sAlbum As String, sGenre As String, sYear As String, sComment As String
    Dim cfile As New cMP3

    'For playlist
    Dim sTrackName As String, sFileEx As String
    Dim sCleanStr As String, sNewString As String, sFormat As String
    Dim sSplitField() As String
    Dim iSpaces As Integer

    If Not FileExists(sPath) Or frmLibrary.Exist_in_Library(sPath) Then Exit Function

    Set dDrive = FS.Drives(Left(sPath, 1))
    If dDrive.DriveType = CDRom Then
        bCDROM = True
    Else
        bCDROM = False
    End If
    cfile.Read_MPEGInfo = True
    cfile.Read_File_Tags sPath

    sTitle = Replace(cfile.Title, "'", " ", , , vbTextCompare)
    sArtist = Replace(cfile.Artist, "'", " ", , , vbTextCompare)
    sAlbum = Replace(cfile.Album, "'", " ", , , vbTextCompare)
    sYear = Replace(cfile.Year, "'", " ", , , vbTextCompare)
    sGenre = Replace(cfile.Genre, "'", " ", , , vbTextCompare)
    sComment = Replace(cfile.Comment, "'", " ", , , vbTextCompare)
    sTrackName = GetFileTitle(sPath)
    sFileEx = Right(sPath, 3)

    If sTitle = "" Then sTitle = GetFileTitle(sPath)
    If sArtist = "" Then sArtist = "Unknown"
    If sAlbum = "" Then sAlbum = "Unknown"
    'If sYear = "" Then sYear = Year(Now())
    If sGenre = "" Then sGenre = "Other"
    If sComment = "" Then sComment = "Uncomment"

    If bAddtracktoPlaylist Then
        ' load tags
        sFormat = sFormatPlayList
        '// Song Name
        sFormat = Replace(sFormatPlayList, "%S", Trim(sTitle))
        '// Artist
        sFormat = Replace(sFormat, "%A", Trim(sArtist))
        '// Album
        sFormat = Replace(sFormat, "%B", Trim(sAlbum))
        '// Year
        sFormat = Replace(sFormat, "%Y", Trim(sYear))
        '// Genre
        sFormat = Replace(sFormat, "%G", Trim(sGenre))
        '// Time
        ' sFormat = Replace(sFormat, "%T", Trim(cFile.MPEG_DurationTime))
        '// File Name
        sFormat = Replace(sFormat, "%N", sTrackName)
        '// Time
        sFormat = Replace(sFormat, "%P", sPath)
        '// File extencion
        sFormat = Replace(sFormat, "%F", sFileEx)

        If sFormat = sFormatPlayList Then sFormat = sTrackName    'If no fomat is there / tag info fails then display trackname
        '------------------------------------------------------------------------------
        sCleanStr = Trim$(sFormat)

        'Upper case and / or lower case the string correctly.
        sSplitField = Split(sCleanStr, " ", , vbTextCompare)
        sCleanStr = ""
        'Debug.Print sCleanStr
        For iSpaces = 0 To UBound(sSplitField)
            If (Not iSpaces = 0 Or Not IsNumeric(sSplitField(iSpaces))) And sSplitField(iSpaces) <> "" Then
                sNewString = UCase$(Left$(sSplitField(iSpaces), 1))
                sNewString = sNewString & (Right$(sSplitField(iSpaces), Len(sSplitField(iSpaces)) - 1))
                sCleanStr = sCleanStr & sNewString & " "
            End If
        Next iSpaces
        sFormat = Trim$(sCleanStr)
        '------------------------------------------------------------------------------

        frmPLST.clist.AddItem sFormat, sPath, cfile.MPEG_DurationTime
    End If

    rst.AddNew
    rst!File = sPath
    rst!Title = sTitle
    rst!Artist = sArtist
    rst!Album = sAlbum
    rst!Year = IIf(Len(sYear) <= 4, sYear, Left(sYear, 4))
    rst!Genre = sGenre
    rst!Comments = sComment
    rst!Length = cfile.MPEG_DurationTime
    rst!bytes = cfile.FileSize
    rst!Seconds = cfile.DurationInSecs
    rst!LastUpdate = Now    '
    rst!Playcount = 0
    rst!Quality = cfile.Quality
    rst!Situation = cfile.Situation
    '   rst!Mood = cFile.Mood              '"not currently implemented"
    rst!FilePath = getDirFromPath(sPath)
    rst!OnCD = bCDROM
    rst!drive = Left(sPath, 3)
    rst.Update

    AddTrack = True
    Exit Function
errorHandler:
    Debug.Print err.Description
End Function
Private Sub txtSkipsize_Change()
    On Error GoTo error:
    lSkipSize = CInt(txtSkipsize.Text)
    Exit Sub
error:
    txtSkipsize.Text = str(lSkipSize)
    MsgBox "Input numeric values only"
End Sub
