VERSION 5.00
Begin VB.Form frmPLST 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   0  'None
   ClientHeight    =   3225
   ClientLeft      =   5940
   ClientTop       =   4440
   ClientWidth     =   4020
   ControlBox      =   0   'False
   FillColor       =   &H000000FF&
   FillStyle       =   0  'Solid
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "playlist.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   215
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   268
   ShowInTaskbar   =   0   'False
   Visible         =   0   'False
   Begin VB.Timer tmrMove 
      Enabled         =   0   'False
      Interval        =   17
      Left            =   3090
      Top             =   2040
   End
   Begin VB.Timer TmrList 
      Interval        =   50
      Left            =   3390
      Top             =   1140
   End
   Begin VB.VScrollBar Scrollbar 
      Height          =   2175
      Left            =   2280
      TabIndex        =   2
      Top             =   330
      Visible         =   0   'False
      Width           =   225
   End
   Begin VB.PictureBox PicSlidertemp 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00008000&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   150
      Left            =   5070
      ScaleHeight     =   10
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   56
      TabIndex        =   19
      Top             =   1890
      Visible         =   0   'False
      Width           =   840
   End
   Begin VB.PictureBox PicBottomRight 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   405
      Left            =   2310
      MousePointer    =   15  'Size All
      ScaleHeight     =   27
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   49
      TabIndex        =   17
      Top             =   2700
      Width           =   735
      Begin MMPlayerXProject.Button plstButton 
         Height          =   120
         Index           =   4
         Left            =   120
         TabIndex        =   18
         Top             =   150
         Width           =   330
         _extentx        =   582
         _extenty        =   212
         style           =   1
         buttoncolor     =   12632256
         mousepointer    =   99
      End
   End
   Begin VB.PictureBox PicTemp 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00008000&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   150
      Left            =   5370
      ScaleHeight     =   10
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   56
      TabIndex        =   16
      Top             =   1290
      Visible         =   0   'False
      Width           =   840
   End
   Begin VB.PictureBox Picright 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00400000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   630
      Left            =   5010
      ScaleHeight     =   42
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   30
      TabIndex        =   11
      Top             =   330
      Width           =   450
      Begin MMPlayerXProject.Button plstButton 
         Height          =   90
         Index           =   5
         Left            =   0
         TabIndex        =   12
         Top             =   30
         Visible         =   0   'False
         Width           =   225
         _extentx        =   397
         _extenty        =   159
         style           =   1
         buttoncolor     =   12632256
         mousepointer    =   99
      End
      Begin MMPlayerXProject.Button plstButton 
         Height          =   90
         Index           =   7
         Left            =   270
         TabIndex        =   13
         Top             =   30
         Width           =   225
         _extentx        =   397
         _extenty        =   159
         style           =   1
         buttoncolor     =   12632256
         mousepointer    =   99
      End
   End
   Begin VB.PictureBox PicActivate 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00008000&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   150
      Left            =   120
      ScaleHeight     =   10
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   184
      TabIndex        =   14
      Top             =   120
      Width           =   2760
   End
   Begin VB.PictureBox PicBottomLeft 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   405
      Left            =   30
      ScaleHeight     =   27
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   137
      TabIndex        =   6
      Top             =   2700
      Width           =   2055
      Begin MMPlayerXProject.Button plstButton 
         Height          =   120
         Index           =   0
         Left            =   60
         TabIndex        =   7
         Top             =   150
         Width           =   330
         _extentx        =   582
         _extenty        =   212
         style           =   1
         buttoncolor     =   12632256
         mousepointer    =   99
      End
      Begin MMPlayerXProject.Button plstButton 
         Height          =   120
         Index           =   1
         Left            =   540
         TabIndex        =   8
         Top             =   150
         Width           =   330
         _extentx        =   582
         _extenty        =   212
         style           =   1
         buttoncolor     =   12632256
         mousepointer    =   99
      End
      Begin MMPlayerXProject.Button plstButton 
         Height          =   120
         Index           =   2
         Left            =   1020
         TabIndex        =   9
         Top             =   150
         Width           =   330
         _extentx        =   582
         _extenty        =   212
         style           =   1
         buttoncolor     =   12632256
         mousepointer    =   99
      End
      Begin MMPlayerXProject.Button plstButton 
         Height          =   120
         Index           =   3
         Left            =   1500
         TabIndex        =   10
         Top             =   150
         Width           =   330
         _extentx        =   582
         _extenty        =   212
         style           =   1
         buttoncolor     =   12632256
         mousepointer    =   99
      End
   End
   Begin VB.PictureBox picList 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      FillColor       =   &H00F4F5F7&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2475
      Left            =   60
      MousePointer    =   99  'Custom
      OLEDropMode     =   1  'Manual
      ScaleHeight     =   165
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   199
      TabIndex        =   3
      Top             =   150
      Width           =   2985
      Begin VB.PictureBox picBack 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   13350
         Left            =   2880
         MouseIcon       =   "playlist.frx":000C
         MousePointer    =   99  'Custom
         Picture         =   "playlist.frx":016A
         ScaleHeight     =   890
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   8
         TabIndex        =   4
         Top             =   0
         Width           =   120
         Begin VB.PictureBox picBar 
            AutoRedraw      =   -1  'True
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   0
            MousePointer    =   99  'Custom
            Picture         =   "playlist.frx":3098
            ScaleHeight     =   17
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   10
            TabIndex        =   5
            Top             =   390
            Visible         =   0   'False
            Width           =   150
         End
      End
   End
   Begin VB.PictureBox picBarDown 
      AutoRedraw      =   -1  'True
      BackColor       =   &H0000FF00&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3450
      MousePointer    =   99  'Custom
      Picture         =   "playlist.frx":328A
      ScaleHeight     =   15.111
      ScaleMode       =   0  'User
      ScaleWidth      =   8
      TabIndex        =   1
      Top             =   1650
      Visible         =   0   'False
      Width           =   120
   End
   Begin VB.PictureBox picBarOver 
      AutoRedraw      =   -1  'True
      BackColor       =   &H000000FF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   3450
      MousePointer    =   99  'Custom
      Picture         =   "playlist.frx":3464
      ScaleHeight     =   18
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   8
      TabIndex        =   0
      Top             =   450
      Visible         =   0   'False
      Width           =   120
   End
   Begin MMPlayerXProject.Button plstButton 
      Height          =   120
      Index           =   6
      Left            =   30
      TabIndex        =   15
      Top             =   30
      Width           =   225
      _extentx        =   397
      _extenty        =   212
      style           =   1
      buttoncolor     =   12632256
      mousepointer    =   99
   End
End
Attribute VB_Name = "frmPLST"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim iCount As Integer
Private Const MAX_PATH As Long = 360&
Private shiftInit As Integer
Private shiftStop As Integer
Private Type FILETIME
    dwLowDateTime As Long
    dwHighDateTime As Long
End Type

'///Constant for flags in drawtext API call
Private Const DT_PATH_ELLIPSIS As Long = &H4000&
Private Const DT_END_ELLIPSIS As Long = &H8000&
Private Const DT_SINGLELINE As Long = &H20&
Private Const DT_VCENTER             As Long = &H4&
Private Const DT_LEFT                As Long = &H0&

Private Type WIN32_FIND_DATA
    dwFileAttributes As Long
    ftCreationTime As FILETIME
    ftLastAccessTime As FILETIME
    ftLastWriteTime As FILETIME
    nFileSizeHigh As Long
    nFileSizeLow As Long
    dwReserved0 As Long
    dwReserved1 As Long
    cFileName As String * MAX_PATH
    cAlternate As String * 14
End Type


Private Enum eImg
    Normal = 0
    Down = 1
    Over = 2
End Enum

Private bMouseOver As Boolean, bMouseDown As Boolean
Attribute bMouseDown.VB_VarUserMemId = 1073938436
Private bSkinLoading As Boolean
Attribute bSkinLoading.VB_VarUserMemId = 1073938439

Private Declare Function FindFirstFile Lib "kernel32" Alias "FindFirstFileA" (ByVal lpFileName$, lpFindFileData As WIN32_FIND_DATA) As Long
Private Declare Function FindNextFile Lib "kernel32" Alias "FindNextFileA" (ByVal hFindFile&, lpFindFileData As WIN32_FIND_DATA) As Long
Private Declare Function FindClose Lib "kernel32" (ByVal hFindFile&) As Long
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Function ClientToScreen Lib "user32" (ByVal hwnd As Long, lpPoint As POINTAPI) As Long

Dim Value As Long
Attribute Value.VB_VarUserMemId = 1073938440
Dim iY As Long
Attribute iY.VB_VarUserMemId = 1073938443
Dim bDrag As Boolean
Attribute bDrag.VB_VarUserMemId = 1073938444
Dim iMin As Long
Attribute iMin.VB_VarUserMemId = 1073938445
Dim iMax As Long
Attribute iMax.VB_VarUserMemId = 1073938446
Dim iValue As Long
Attribute iValue.VB_VarUserMemId = 1073938447

Dim FONT_TyPE As Boolean
Attribute FONT_TyPE.VB_VarUserMemId = 1073938449

Dim ExtendedSelection
Attribute ExtendedSelection.VB_VarUserMemId = 1073938450
Dim PreviousSelection
Attribute PreviousSelection.VB_VarUserMemId = 1073938451

Private Const FILE_ATTRIBUTE_DIRECTORY As Long = &H10
Dim check As Integer
Attribute check.VB_VarUserMemId = 1073938452
Dim currentHeight As Long
Attribute currentHeight.VB_VarUserMemId = 1073938456

Dim InFormDrag As Boolean
Attribute InFormDrag.VB_VarUserMemId = 1073938457

Dim cAjustarDesk As New clsDockingHandler
Attribute cAjustarDesk.VB_VarUserMemId = 1073938458
Dim cWindows As New cWindowSkin
Attribute cWindows.VB_VarUserMemId = 1073938459
Public clist As New clist
Attribute clist.VB_VarUserMemId = 1073938460
Dim range As Integer
Attribute range.VB_VarUserMemId = 1073938461
Dim lCounter As Long    'needed for timer in list display
Attribute lCounter.VB_VarUserMemId = 1073938465

Private Sub DrawBar(ImgState As eImg)
    On Error Resume Next
    Dim intY As Integer, intX As Integer

    If iMin < 0 Then iValue = -iValue Else iValue = iMax - iValue
    iY = (iValue - iMin) / (iMax - iMin) * (picBack.Height - picBar.Height)
    intX = 0: intY = iY

    picBack.cls

    '//IMAGE OVER
    If bMouseOver = True Then
        If bMouseDown = True Then
            Call BitBlt(picBack.hDC, intX, intY, picBar.ScaleWidth, picBar.ScaleHeight, _
                        picBarDown.hDC, 0, 0, vbSrcCopy)
        Else
            Call BitBlt(picBack.hDC, intX, intY, picBar.ScaleWidth, picBar.ScaleHeight, _
                        picBarOver.hDC, 0, 0, vbSrcCopy)
        End If
        picBack.Refresh
        Exit Sub
    End If

    If ImgState = Normal Then
        Call BitBlt(picBack.hDC, intX, intY, picBar.ScaleWidth, picBar.ScaleHeight, _
                    picBar.hDC, 0, 0, vbSrcCopy)
    ElseIf ImgState = Down Then
        Call BitBlt(picBack.hDC, intX, intY, picBar.ScaleWidth, picBar.ScaleHeight, _
                    picBarDown.hDC, 0, 0, vbSrcCopy)
    ElseIf ImgState = Over Then
        Call BitBlt(picBack.hDC, intX, intY, picBar.ScaleWidth, picBar.ScaleHeight, _
                    picBarOver.hDC, 0, 0, vbSrcCopy)
    End If

    picBack.Refresh
End Sub

Public Sub ActivateMe()
    PicActivate.Visible = False
    Dim i
    For i = 5 To 7
        plstButton(i).Selected = False
    Next
End Sub

Public Sub DeactivateMe()
    On Error Resume Next
    Dim iLeft
    iLeft = (Me.ScaleWidth / 2) - (PicTemp.ScaleWidth / 2) - PicActivate.Left

    PicActivate.cls
    PicActivate.PaintPicture PicTemp.Picture, iLeft, 0, PicTemp.Width, PicTemp.ScaleHeight, 0, 0
    PicActivate.Visible = True
    Dim i
    For i = 5 To 7
        plstButton(i).Selected = True
    Next
End Sub

Private Sub Form_Activate()
    If bSkinLoading Then
        cWindows.CargarSkin tAppConfig.AppConfig & "Skins\" & tAppConfig.Skin & "\PLAY_LIST\"
        picList.Move cWindows.AreaLeft, cWindows.AreaTop, cWindows.AreaWidth, cWindows.AreaHeight
        Form_Resize1
        Update_Plst_Scrollbar
        frmPLST.Scrollbar.Value = frmPLST.Scrollbar.Max
        bSkinLoading = False
    Else
        Form_Resize1
    End If
End Sub


Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
''Debug.Print str(KeyCode)
'On Error GoTo errorHandler
    Select Case KeyCode

    Case &H28:    'keydown=40
        Dim i As Integer
        If bMultiSelect Then
            For i = 0 To clist.ItemCount - 1
                Call clist.ChangeSelection(i, False)
            Next
            Call clist.ChangeSelection(NormalSelection, True)
        End If
        If NormalSelection < clist.ItemCount - 1 Then
            Call clist.ChangeSelection(NormalSelection, False)
            NormalSelection = NormalSelection + 1
            Call clist.ChangeSelection(NormalSelection, True)
            ReinitializeList
            If NormalSelection > (Scrollbar.Value + (ListRange - 1)) Then
                Scrollbar.Value = Scrollbar.Value + 1
            ElseIf NormalSelection < Scrollbar.Value Then
                Scrollbar.Value = Scrollbar.Value - 1
            End If
        End If

    Case &H26:    'keyup=38
        If bMultiSelect Then
            For i = 0 To clist.ItemCount - 1
                Call clist.ChangeSelection(i, False)
            Next
            Call clist.ChangeSelection(NormalSelection, True)
        End If
        If NormalSelection > 0 Then
            Call clist.ChangeSelection(NormalSelection, False)
            NormalSelection = NormalSelection - 1
            Call clist.ChangeSelection(NormalSelection, True)
            ReinitializeList
            If NormalSelection > (Scrollbar.Value + (ListRange - 1)) Then
                Scrollbar.Value = Scrollbar.Value + 1
            ElseIf NormalSelection < Scrollbar.Value Then
                Scrollbar.Value = Scrollbar.Value - 1
            End If
        End If

    Case &H25:
        If frmMain.PlayerIsPlaying = "true" Then frmMain.Five_Seg_Backward

    Case &H27:
        If frmMain.PlayerIsPlaying = "true" Then frmMain.Five_Seg_Forward

    Case 33    'page up
        'Listrange stores information that how many number of tracks can be shown at once in current playlist
        If clist.ItemCount <= ListRange Then Exit Sub
        If Scrollbar.Value > 0 Then
            If Scrollbar.Value - ListRange >= 0 Then
                Scrollbar.Value = Scrollbar.Value - ListRange    'Show previous listrange number of tracks
            Else
                Scrollbar.Value = 0
            End If
            ReinitializeList
        End If
    Case 34    'page down

        If clist.ItemCount <= ListRange Then Exit Sub
        If Scrollbar.Value >= 0 And Scrollbar.Value < Scrollbar.Max Then
            If Scrollbar.Value + ListRange <= Scrollbar.Max Then
                Scrollbar.Value = Scrollbar.Value + ListRange    'page Down ==Show next  listrange number of tracks
            Else
                Scrollbar.Value = Scrollbar.Max
            End If
            ReinitializeList    'redraw playlist only when required
        End If

    Case 46:
        i = 0
        While (i <= clist.ItemCount - 1)
            If clist.bSelected(i) = True Then
                REMOVE_ITEM (i)
            Else
                i = i + 1
            End If
        Wend
        NormalSelection = -1
        frmPLST.Update_Plst_Scrollbar
        ' frmPLST.Scrollbar.Value = frmPLST.Scrollbar.Max
        ReinitializeList

    Case 36:    'HOME
        If Scrollbar.Value <> Scrollbar.Min Then
            Scrollbar.Value = Scrollbar.Min
            ReinitializeList
        End If

    Case 35:    'END
        If Scrollbar.Value <> Scrollbar.Max Then
            Scrollbar.Value = Scrollbar.Max
            ReinitializeList
        End If
    End Select

errorHandler:
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        picList_DblClick
    Else
        Call frmMain.Form_KeyUp(KeyCode, Shift)
    End If
End Sub

Private Sub Form_Load()
    Dim strRes
    'Call DragAcceptFiles(Me.hWnd, True)
    bSkinLoading = True
    'Call EnableFileDrops(Me)  'Doing through OLEDRAGDRop
    Hook (Me.hwnd)

    'Initialise values
    FONT_TyPE = False
    CurrentTrack_Index = 0
    NormalSelection = -1
    LstTextHeight = picList.TextHeight("")
    ExtendedSelection = -1
    Scrollbar.Value = Scrollbar.Min
    iValue = Scrollbar.Max
    currentHeight = picList.Height

    If tAppConfig.AppPath = "" Then
        tAppConfig.AppPath = App.Path
        'check for directory or drive
        If Right(tAppConfig.AppPath, 1) <> "\" Then tAppConfig.AppPath = tAppConfig.AppPath & "\"
        'to search fromm application config ini file the path of config file
        tAppConfig.AppConfig = Read_INI("Configuration", "AppConfiguration", tAppConfig.AppPath & "MMp3Player\", , True)
        If Right(tAppConfig.AppConfig, 1) <> "\" Then tAppConfig.AppConfig = tAppConfig.AppConfig & "\"
        If Dir(tAppConfig.AppConfig, vbDirectory) = "" Then tAppConfig.AppConfig = tAppConfig.AppPath & "MMp3Player\"
    End If

    strRes = Read_INI("Configuration", "PX", 10092, , True)
    If IsNumeric(strRes) = False Then strRes = 0
    frmPLST.Left = CInt(strRes)

    strRes = Read_INI("Configuration", "PY", 2430, , True)
    If IsNumeric(strRes) = False Then strRes = 0
    frmPLST.Top = CInt(strRes)

    strRes = Read_INI("Configuration", "PW", 4065, , True)
    'If IsNumeric(strRes) = False Then strRes = 100
    frmPLST.Width = CInt(strRes)
    'If frmPLST..Width < min_plstWidth Then frmPLST.picList.Width = min_plstWidth

    strRes = Read_INI("Configuration", "PH", 3390, , True)
    'If IsNumeric(strRes) = False Then strRes = 100
    frmPLST.Height = CInt(strRes)

    picList.Height = LstTextHeight * CInt((picList.Height) / LstTextHeight)
    Update_Plst_Scrollbar
    frmPLST.Scrollbar.Value = frmPLST.Scrollbar.Max
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    cWindows.Formulario_Down X, Y
    cAjustarDesk.StartDockDrag X * Screen.TwipsPerPixelX, _
                               Y * Screen.TwipsPerPixelY
    InFormDrag = True
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    cWindows.Formulario_MouseMove Button, X, Y
    If InFormDrag And cWindows.Ajustando = False Then
        ' Continue window draggin'
        cAjustarDesk.UpdateDockDrag X * Screen.TwipsPerPixelX, _
                                    Y * Screen.TwipsPerPixelY
        'bHookForm = False
        Exit Sub
    End If

    If cWindows.Ajustando = True Then
        picList.Move cWindows.AreaLeft, cWindows.AreaTop, cWindows.AreaWidth, cWindows.AreaHeight
        Form_Resize1
        ReinitializeList
    End If
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If cWindows.Ajustando = True Then
        ReinitializeList
    End If
    cWindows.Formulario_MouseUp X, Y
    InFormDrag = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Unhook (Me.hwnd)
End Sub



Private Sub PicBottomLeft_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call ReleaseCapture
    Call SendMessage(Me.hwnd, WM_NCLBUTTONDOWN, 2, 0&)
End Sub


Private Sub picBack_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If clist.ItemCount <= ListRange Then Exit Sub
    iY = Y
    If bDrag And Button = 1 Then    '// dragging
        '// vertical
        If Y < picBar.ScaleHeight / 2 Then
            iY = 0    'iY - picBar.Height / 2
            Scrollbar.Value = Scrollbar.Min
        ElseIf Y > picBack.ScaleHeight - picBar.ScaleHeight / 2 Then
            iY = picBack.ScaleHeight - picBar.ScaleHeight
            Scrollbar.Value = Scrollbar.Max
        Else
            iY = Y - picBar.Height / 2
            '// horizontal
            If CalcValue <= Scrollbar.Max And CalcValue >= Scrollbar.Min Then Scrollbar.Value = Scrollbar.Max - CalcValue
        End If
    Else
        '// mouse over
        If bMouseOver = False Then bMouseOver = True
    End If
End Sub


Private Sub picBack_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If clist.ItemCount <= ListRange Then Exit Sub

    If bDrag = False Then
        If CalcValue <= Scrollbar.Max And CalcValue >= Scrollbar.Min Then Scrollbar.Value = Scrollbar.Max - CalcValue
    End If
    bMouseDown = False
    iY = Y
    iValue = Scrollbar.Max - Scrollbar.Value
    Call DrawBar(Normal)
    bDrag = False
End Sub

Private Sub picBack_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If clist.ItemCount <= ListRange Then Exit Sub
    If Button = 1 Then
        '// vertical
        If Y >= iY And Y <= iY + picBar.ScaleHeight And Button = 1 Then
            If Y < picBar.Height / 2 Then
                iY = 0    'iY - picBar.Height / 2
            ElseIf Y > picBack.Height - picBar.Height Then
                iY = picBack.Height - picBar.Height
            Else
                iY = Y - picBar.Height / 2
            End If
            bDrag = True
            bMouseDown = True
            If CalcValue = Scrollbar.Max - Scrollbar.Value Then
                Call DrawBar(Down)
            ElseIf CalcValue <= Scrollbar.Max And CalcValue >= Scrollbar.Min Then
                Scrollbar.Value = Scrollbar.Max - CalcValue
                'Call DrawBar(Normal)
            End If

        Else
            iY = Y
            If iY > picBack.ScaleHeight - (picBar.ScaleHeight / 2) Then iY = picBack.ScaleHeight - (picBar.ScaleHeight / 2)
            If iY < picBar.ScaleHeight / 2 Then iY = picBar.ScaleHeight / 2
            iY = iY - picBar.ScaleHeight / 2
        End If
    End If
End Sub
'
Private Function CalcValue() As Integer
'calculates current scrollbar value
    On Error Resume Next
    iValue = iY / (picBack.Height - picBar.Height) * (iMax - iMin) + iMin
    If iMin < 0 Then iValue = -iValue Else iValue = iMax - iValue
    CalcValue = iValue
End Function

Public Sub Form_Load1()
    LstTextWidth = picList.TextWidth("a")
    PicBottomLeft.Left = 0
    bSkinLoading = True
    LoadSkin
    Form_Resize1
End Sub

Public Sub Form_Resize1()
    If bLoading Then Exit Sub
    picList.Height = cWindows.AreaHeight    'lll * LstTextHeight
    picBack.Move picList.Width - picBack.Width, 0, picBack.Width, picList.Height
    picBack.Refresh
    PicBottomRight.Move (Me.ScaleWidth - PicBottomRight.Width), Me.ScaleHeight - PicBottomRight.Height, PicBottomRight.Width, PicBottomRight.Height
    Picright.Move (Me.ScaleWidth - Picright.Width), 0, Picright.Width, Picright.Height
    ListRange = Fix(picList.ScaleHeight / LstTextHeight)
    range = picList.ScaleWidth - picList.TextWidth(" 00:00 ") - 3 * picBack.Width - 2

    If clist.ItemCount > ListRange Then
        Scrollbar.Max = clist.ItemCount - ListRange    'The list items are too much to _
                                                       display we need to scroll the list
    Else
        Scrollbar.Max = 0                          'There is no need of scrolling
    End If

    iMax = Scrollbar.Max
    iValue = Scrollbar.Max - (Scrollbar.Value)
    DrawBar (Normal)
    ReinitializeList
    PicBottomLeft.Top = cWindows.AreaTop + cWindows.AreaHeight
    DoEvents
End Sub

Public Sub ClearList()
    clist.Clear
    picList.cls
    NormalSelection = -1
    CurrentTrack_Index = -1
    Update_Plst_Scrollbar
    frmPLST.Scrollbar.Value = 0
End Sub

Private Sub picBar_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    picBar.Picture = picBarOver.Picture
End Sub

Private Sub PicBottomRight_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call Form_MouseDown(Button, Shift, PicBottomRight.Left + X, PicBottomRight.Top + Y)
End Sub

Private Sub PicBottomRight_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call Form_MouseMove(Button, Shift, PicBottomRight.Left + X, PicBottomRight.Top + Y)
End Sub

Private Sub PicBottomRight_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call Form_MouseUp(Button, Shift, PicBottomRight.Left + X, PicBottomRight.Top + Y)
End Sub

Private Sub picList_DblClick()
    If clist.ItemCount <= 0 Or NormalSelection = -1 Then Exit Sub
    ' On Error Resume Next
    CurrentTrack_Index = NormalSelection
    Debug.Print CurrentTrack_Index
    ReinitializeList
    sFileMainPlaying = clist.exItem(CurrentTrack_Index)
    frmMain.PlayerIsPlaying = "false"
    If cGetInputState() <> 0 Then DoEvents

    Stream_Stop lCurrentChannel
    frmMain.Play
End Sub




Private Sub picList_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim FileName As String
    Dim icnt, ipos As Integer
    'If Effect <> 7 And Effect <> 1 Then Exit Sub
    ''Debug.Print Data.GetData(vbCFText)
    On Error Resume Next
    If Data.Files.Count = 0 Then Exit Sub
    If Data.GetFormat(vbCFFiles) = False Then Exit Sub
    Dim Index As Long
    Index = HitTestPlst(X, Y)
    If Index = -1 Then Index = clist.ItemCount

    For icnt = 1 To Data.Files.Count
        If FileExists(Data.Files(icnt)) Then
            'This function will add the index of this file added to the listview in order to
            'create a sequence of playback in normal sequential mode or shuffle mode
            If cGetInputState() <> 0 Then DoEvents
            If isMediaFile(Data.Files(icnt)) Then
                ipos = InStrRev(Data.Files(icnt), "\")
                FileName = (Mid(Data.Files(icnt), ipos + 1, Len(Data.Files(icnt)) - ipos - 4))
                If Len(FileName) >= 50 Then FileName = Left(FileName, 34) + "..."
                Call frmPLST.Add_track_to_Playlist(Data.Files(icnt), Index)
                Index = Index + 1
            ElseIf isPlaylistFile(Data.Files(icnt)) Then
                Index = Index + LoadPlaylist(Data.Files(icnt), False, Index)
            ElseIf FolderExists(Data.Files(icnt)) Then
                Index = Index + addFilesfromDir(Data.Files(icnt), False, False, Index)
                frmPLST.ReinitializeList
            Else
                Call frmPLST.Add_track_to_Playlist(Data.Files(icnt), Index)
                Index = Index + 1
            End If

        End If
1:
    Next icnt

    frmPLST.Update_Plst_Scrollbar
    If clist.ItemCount > ListRange Then
        Scrollbar.Max = clist.ItemCount - ListRange    'The list items are too much to _
                                                       display we need to scroll the list
    Else
        Scrollbar.Max = 0                           'There is no need for scrolling
    End If

    ReinitializeList
End Sub


Private Sub picList_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    bMultiSelect = True
    If clist.ItemCount <= 0 Then Exit Sub
    shiftInit = -1    'continuous selection when shift is pressed
    shiftStop = -1    'NormalSelection
    PreviousSelection = NormalSelection    'Store prev selected track for further use (multiselection)
    'If NormalSelection >= 0 And NormalSelection <= clist.ItemCount Then Call clist.ChangeSelection(NormalSelection, True)

    If Shift = 0 Then
        Dim i As Integer

        If Fix(Y / LstTextHeight) + Scrollbar.Value >= 0 And Fix(Y / LstTextHeight) + Scrollbar.Value < clist.ItemCount Then
            NormalSelection = Fix(Y / LstTextHeight) + Scrollbar.Value
            ' 'Debug.Print NormalSelection
            If clist.bSelected(NormalSelection) = True Then GoTo checkpopup:
            tmrMove.Enabled = True
            For i = 0 To clist.ItemCount - 1
                Call clist.ChangeSelection(i, False)
            Next
            Call clist.ChangeSelection(NormalSelection, True)
            If clist.ItemCount > 0 Then
                ReinitializeList
            End If
            If CurrentTrack_Index < NormalSelection Then
                check = 1
            ElseIf CurrentTrack_Index > NormalSelection Then
                check = -1
            Else
                check = 0
            End If

        Else
            NormalSelection = -1
            For i = 0 To clist.ItemCount - 1
                Call clist.ChangeSelection(i, False)
            Next
            ReinitializeList
        End If
    ElseIf Shift = 1 Then    'shift is pressed

        If Fix(Y / LstTextHeight) + Scrollbar.Value >= 0 And Fix(Y / LstTextHeight) + Scrollbar.Value < clist.ItemCount Then
            NormalSelection = Fix(Y / LstTextHeight) + Scrollbar.Value
            If PreviousSelection > -1 And NormalSelection > -1 Then    'No prev track is selected
                Dim j, k As Long
                j = PreviousSelection
                k = Min(j, NormalSelection)
                j = maxX(j, NormalSelection)
                For i = k To j
                    Call clist.ChangeSelection(i, True)    'Select all continuos tracks fro prev selection to currwnt  selection
                Next
            End If
            shiftInit = NormalSelection
            shiftStop = NormalSelection

            Call clist.ChangeSelection(NormalSelection, True)

            If clist.ItemCount > 0 Then
                ReinitializeList
            End If
        ElseIf PreviousSelection > -1 Then
            NormalSelection = clist.ItemCount - 1
            For i = PreviousSelection To NormalSelection
                Call clist.ChangeSelection(i, True)    'Select all continuos tracks fro prev selection to currwnt  selection
            Next
            ReinitializeList
        Else
            ' NormalSelection = -1
            ' ReinitializeList
        End If
    ElseIf Shift = 2 Then

        If Fix(Y / LstTextHeight) + Scrollbar.Value >= 0 And Fix(Y / LstTextHeight) + Scrollbar.Value < clist.ItemCount Then
            k = Fix(Y / LstTextHeight) + Scrollbar.Value
            If clist.bSelected(k) = False Then NormalSelection = k    'Else NormalSelection = PreviousSelection
            Call clist.ChangeSelection(k, Not (clist.bSelected(k)))
            ''Debug.Print k, clist.bSelected(k)
            ReinitializeList
        End If

    End If

checkpopup:
    If Button = 2 And NormalSelection <> -1 Then PopupMenu frmPopUp.mnuPlst_rightClick

End Sub

Private Sub picList_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error Resume Next
    Dim ItemBuffer As String
    Dim ExItemBuffer As String
    Dim ExTracklengthBuffer As String
    Dim ExSelectedBuffer As Boolean
    If NormalSelection = -1 Or clist.ItemCount <= 0 Then Exit Sub    'if no item is selected , or there is nothing in list
    '//////////CODE FOR OLEDRAGDATA
    'If Button = 1 And NormalSelection >= 0 Then picList.OLEDrag
    '//////////CODE FOR OLEDRAGDATA



    If Button <> 0 Then PreviousSelection = NormalSelection    'very important)


    If Button = 1 Then

        If Y > picList.Height Then
            If NormalSelection < clist.ItemCount - 1 Then NormalSelection = NormalSelection + 1
            If Scrollbar.Value < Scrollbar.Max Then Scrollbar.Value = Scrollbar.Value + 1
            tmrMove.Enabled = True
        ElseIf Y < 0 And Scrollbar.Value >= 1 Then    'NormalSelection <= Scrollbar.Value And Scrollbar.Value > 0 Then
            NormalSelection = NormalSelection - 1
            Scrollbar.Value = Scrollbar.Value - 1
            tmrMove.Enabled = True
        ElseIf Fix(Y / LstTextHeight) + Scrollbar.Value >= 0 And Fix(Y / LstTextHeight) + Scrollbar.Value < clist.ItemCount Then
            NormalSelection = Fix(Y / LstTextHeight) + Scrollbar.Value    'tell which track is correnly under mouse
            tmrMove.Enabled = False
        Else
            tmrMove.Enabled = False
            Exit Sub
        End If

        If Shift = 0 Then    ' normal mouse operation
            Dim i, k
            k = NormalSelection - PreviousSelection    'number of steps entry should be dragged/moved up or down
            If k <= -1 Then
                For i = 0 To clist.ItemCount - 1    '
                    If clist.bSelected(i) = True Then    'swap the selectd entires only
                        If (i + k) < 0 Or (i + k) > clist.ItemCount - 1 Then Exit Sub    'entry has moved to topmost position

                        ItemBuffer = clist.Item(i)
                        ExItemBuffer = clist.exItem(i)
                        ExTracklengthBuffer = clist.exTracklength(i)
                        ExSelectedBuffer = clist.bSelected(i)

                        clist.ChangeItem i, clist.Item(i + k)
                        clist.ChangeEXItem i, clist.exItem(i + k)
                        clist.ChangeEXTracklength i, clist.exTracklength(i + k)
                        clist.ChangeSelection i, clist.bSelected(i + k)

                        clist.ChangeItem i + k, ItemBuffer
                        clist.ChangeEXItem i + k, ExItemBuffer
                        clist.ChangeEXTracklength i + k, ExTracklengthBuffer
                        clist.ChangeSelection i + k, ExSelectedBuffer

                        'swap correnttrack position also
                        If CurrentTrack_Index = i + k Then
                            CurrentTrack_Index = i
                        ElseIf CurrentTrack_Index = i Then
                            CurrentTrack_Index = i + k
                        End If
                    End If
                Next

            ElseIf k >= 1 Then
                For i = clist.ItemCount - 1 To 0 Step -1    'check should be done in reverse order else swaping will create a problem
                    If clist.bSelected(i) = True Then
                        If (i + k) < 0 Or (i + k) > clist.ItemCount - 1 Then Exit Sub
                        'save information in buffer
                        ItemBuffer = clist.Item(i)
                        ExItemBuffer = clist.exItem(i)
                        ExTracklengthBuffer = clist.exTracklength(i)
                        ExSelectedBuffer = clist.bSelected(i)
                        'swap
                        clist.ChangeItem i, clist.Item(i + k)
                        clist.ChangeEXItem i, clist.exItem(i + k)
                        clist.ChangeEXTracklength i, clist.exTracklength(i + k)
                        clist.ChangeSelection i, clist.bSelected(i + k)

                        clist.ChangeItem i + k, ItemBuffer
                        clist.ChangeEXItem i + k, ExItemBuffer
                        clist.ChangeEXTracklength i + k, ExTracklengthBuffer
                        clist.ChangeSelection i + k, ExSelectedBuffer

                        'swap correnttrack position also
                        If CurrentTrack_Index = i + k Then
                            CurrentTrack_Index = i
                        ElseIf CurrentTrack_Index = i Then
                            CurrentTrack_Index = i + k
                        End If
                    End If
                Next
            End If
            If Abs(k) > 0 Then ReinitializeList: PreviousSelection = NormalSelection    'very important redrawing of list is required

        ElseIf Shift = 1 Then    ' shift is pressed
            If False And shiftInit = -1 Then Exit Sub
            shiftStop = NormalSelection    'final of selection limit selection
            If shiftInit > clist.ItemCount Then shiftInit = clist.ItemCount
            For i = 0 To clist.ItemCount - 1
                Call clist.ChangeSelection(i, False)    'Not (cList.bSelected(NormalSelection)))
            Next
            For i = Min(shiftInit, shiftStop) To maxX(shiftInit, shiftStop)    'select all tracks in range of shiftinit and shiftstop
                Call clist.ChangeSelection(i, True)    'Not (cList.bSelected(NormalSelection)))
                bMultiSelect = True
            Next
            ReinitializeList    'redrawing of list is required
            PreviousSelection = NormalSelection    'very important

        ElseIf Shift = 2 Then
            'Don't do anything
        End If

    End If

End Sub


Private Sub picList_OLEStartDrag(Data As DataObject, AllowedEffects As Long)
    AllowedEffects = vbDropEffectMove
    Data.SetData clist.exItem(NormalSelection), vbCFText
End Sub

Private Sub Picright_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call ReleaseCapture
    Call SendMessage(Me.hwnd, WM_NCLBUTTONDOWN, 2, 0&)
End Sub

Private Sub Picright_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'frmplst.setfocus
End Sub

Private Sub plstButton_Click(Index As Integer)
    Select Case Index

    Case 0:    'files
        PopupMenu frmPopUp.mnuadd

    Case 1:    'Remove options
        PopupMenu frmPopUp.mnuPLSTRemove

    Case 2:    'Select
        PopupMenu frmPopUp.mnuPlstSel

    Case 3:    'Miscelleneous
        PopupMenu frmPopUp.mnuPlstMis

    Case 4:    'list Options
        PopupMenu frmPopUp.mnuPlstOptions
    Case 5:    'min

    Case 6:    ' Mode

    Case 7:    'close
        Me.Hide
        frmMain.Button(26).Selected = False
        frmPopUp.mnuShowPlst.Checked = False
        Call CheckMenu(11, False)
    End Select
End Sub

Private Sub ScrollBar_Change()
    If clist.ItemCount <= ListRange Then Exit Sub

    'On Error Resume Next
    iValue = Scrollbar.Max - (Scrollbar.Value)
    iY = iValue / (iMax - iMin) * (picBack.Height - picBar.Height) - 1
    'If tmrMove.Enabled = False Then ReinitializeList
    ReinitializeList
    DrawBar (Down)
    Debug.Print Scrollbar.Value & Scrollbar.Max

End Sub

Private Sub ScrollBar_Scroll()
    Call ScrollBar_Change
End Sub

Sub ReinitializeList()
'On Error Resume Next
    Dim i
    'Sets the scrollbar max value as the list is changed
    'so the list items can be scrolled
    Dim Track As String
    iMax = Scrollbar.Max
    picList.cls
    'set user defined backcolor
    picList.BackColor = List_Backcolor
    If clist.ItemCount <= 0 Or frmPLST.Visible = False Then Exit Sub
    'Call Gradient

    For i = 0 To clist.ItemCount - 1    'if total number of items is less than display range
        ' If i = NormalSelection Then 'item is selected so draw selected bar
        If iOptPlstCase = 0 Then
            Track = clist.Item(i)
        ElseIf iOptPlstCase = 1 Then
            Track = LCase(clist.Item(i))
        Else
            Track = UCase(clist.Item(i))
        End If
        Drawtrack (i)
    Next
    lCounter = Scrollbar.Value
    TmrList.Enabled = True
    ' no DoEvents take care
End Sub
Sub DrawSelection(i As Integer)

    Dim hBrush As Long
    Dim m As RECT
    If i < Scrollbar.Value Or i > Scrollbar.Value + (ListRange - 1) Then Exit Sub

    m.Left = 1
    m.Right = picList.ScaleWidth - picBar.Width + 1
    m.Top = (i - Scrollbar.Value) * LstTextHeight
    m.Bottom = m.Top + LstTextHeight
    hBrush = CreateSolidBrush(SelectedText_Backcolor)
    FillRect picList.hDC, m, hBrush
    DeleteObject hBrush

End Sub

Private Sub EraseTrack(i As Integer)

    Dim hBrush As Long
    Dim lColor As Long
    Dim m As RECT
    'Erase only when track is visible
    If i < Scrollbar.Value Or i > Scrollbar.Value + (ListRange - 1) Then Exit Sub

    m.Left = 1
    m.Right = picList.ScaleWidth - picBar.Width + 1
    lColor = IIf(clist.bSelected(i), SelectedText_Backcolor, List_Backcolor)
    'Set the position of rectangle according to whether scrolllbar is needed or not
    i = IIf(clist.ItemCount <= ListRange, i, i - Scrollbar.Value)
    m.Top = i * LstTextHeight
    m.Bottom = m.Top + LstTextHeight
    'create brush of backcolor
    hBrush = CreateSolidBrush(List_Backcolor)
    'fill the rectangel with backcolor
    FillRect picList.hDC, m, hBrush
    'we do not need brush now
    DeleteObject hBrush

End Sub

Sub Gradient()
    Dim i
    For i = 0 To picList.ScaleHeight Step LstTextHeight / 2
        picList.Line (0, i)-(picList.ScaleWidth - 18, i + LstTextHeight / 2), RGB(i / 6 + 202, i / 6 + 201, i / 6 + 200), BF
    Next
End Sub

Public Sub ADD_Dir(Optional bRecursion As Boolean = True, Optional bNewlist As Boolean = False, Optional Startfolder As String = "E:\music")
    Dim txtReturn As String
    Dim Flags As Long
    Dim txtInstruction As String
    ' SpecialFolder = 0
    ' Startfolder = "d:\c backup\manvendra\pankaj udas"
    txtInstruction = "Add Folder for NeoMp3 Player"
    ' flags = flags + BIF_USENEWUI
    ' flags = flags + BIF_EDITBOX
    ' flags = flags + BIF_STATUSTEXT
    'Flags = Flags + BIF_NEWDIALOGSTYLE
    'Flags = Flags + BIF_BROWSEINCLUDEFILES
    txtReturn = clsDlg.FolderBrowse(Me.hwnd, txtInstruction, Flags)
    If txtReturn = "" Then Exit Sub
    ' If Right$(txtReturn, 1) <> "\" Then txtReturn = txtReturn & "\"
    bSearching = True
    If bNewlist Then ClearList
    Call addFilesfromDir(txtReturn, bRecursion)
    bSearching = False


    If clist.ItemCount >= 1 Then
        ListRange = Fix(picList.ScaleHeight / LstTextHeight)
        If clist.ItemCount > ListRange Then
            Scrollbar.Max = clist.ItemCount - ListRange    'The list items are too much to _
                                                           display we need to scroll the list
        Else
            Scrollbar.Max = 0                           'There is no need for scrolling
        End If
        iMax = Scrollbar.Max
        Scrollbar.Value = Scrollbar.Max
        If bNewlist = True Then CurrentTrack_Index = 0: sFileMainPlaying = frmPLST.clist.exItem(CurrentTrack_Index): frmMain.Play
        ReinitializeList
    End If
End Sub


Public Sub Enqueue_Track(sPath As String, sCommand As String)

    Dim play_me As Boolean
    On Error GoTo errorHandler:

    sPath = Trim(sPath)

    sCommand = UCase(sCommand)
    play_me = False

    Select Case sCommand

    Case "PLAY"
        frmPLST.ClearList
        frmPLST.Add_track_to_Playlist GetLongFilename(sPath)
        play_me = True
    Case "ADDF"
        frmPLST.Add_track_to_Playlist GetLongFilename(sPath)
    Case "RUND"
        frmPLST.ClearList
        addFilesfromDir sPath, True, True
        play_me = True
    Case "ADDD"
        addFilesfromDir sPath
    Case "ADDP"
        LoadPlaylist sPath, False
    Case "RUNP"
        LoadPlaylist sPath, True
        play_me = True
        CurrentTrack_Index = 0
    Case ""
        frmPLST.ClearList
        play_me = True
        If FolderExists(sPath) Then    'If it is folder
            Call frmPLST.Enqueue_Track(sPath, "RUND")    'Recall function with Run Directory parameter
        ElseIf FileExists(sPath) Then
            If UCase(Right(sPath, 3)) = "NPL" Then
                Call Enqueue_Track(sPath, "RUNP")    'Recall function with Run Playlist parameter
            Else
                Call Enqueue_Track(sPath, "PLAY")    'Recall function with PLAY parameter
            End If
        End If

    End Select

    If bLoading = True Then Exit Sub
    frmPLST.Update_Plst_Scrollbar
    If clist.ItemCount > 0 And play_me Then
        CurrentTrack_Index = 0
        sFileMainPlaying = clist.exItem(CurrentTrack_Index)
        frmMain.Play
    Else
        Scrollbar.Value = Scrollbar.Max
    End If
    ReinitializeList
    Exit Sub
errorHandler:
    MsgBox err.Description
End Sub

Public Sub ADD_MULTIPLE_FILES(newList As Boolean)
    On Error Resume Next
    Dim sfilter As String, sPath As String
    sfilter = "MP3 Files (*.mp3)|*.mp3|" & "All Files (*.*)|*.*|"
    Dim sfiles() As String
    Call clsDlg.OpenAsMultiFileName(frmPLST.hwnd, "Open Music files for NeoMp3 Player", sfiles(), , sfilter)
    Dim nCount As Integer
    sPath = sfiles(0)
    If sPath = "" Then Exit Sub

    If newList = True Then frmPLST.ClearList
    For nCount = 1 To UBound(sfiles)
        Call frmPLST.Add_track_to_Playlist(sPath + "\" + sfiles(nCount), -1)
        If cGetInputState() <> 0 Then DoEvents
    Next nCount

    frmPLST.Update_Plst_Scrollbar
    If newList = True And frmPLST.clist.ItemCount > 0 Then
        CurrentTrack_Index = 0
        sFileMainPlaying = frmPLST.clist.exItem(CurrentTrack_Index)
        frmMain.Play
    End If
    'Set scroll bar to max to show recently added files
    iMax = Scrollbar.Max
    Scrollbar.Value = Scrollbar.Max
    ReinitializeList

End Sub
Public Sub REMOVE_ITEM(Optional Index As Integer = -1)

    If Index = -1 Then
        If NormalSelection >= 0 Then
            Index = NormalSelection
        Else
            Exit Sub
        End If
    End If
    clist.RemoveItem (Index)

    If CurrentTrack_Index = Index Then CurrentTrack_Index = -1    'Then CurrentTrack_Index = CurrentTrack_Index - 1
    If CurrentTrack_Index > Index And CurrentTrack_Index > 0 Then CurrentTrack_Index = CurrentTrack_Index - 1
    'If index = cList.ItemCount Then NormalSelection = NormalSelection - 1
    ListRange = Fix(picList.ScaleHeight / LstTextHeight)
    If clist.ItemCount > ListRange Then
        Scrollbar.Max = clist.ItemCount - ListRange    'The list items are too much to _
                                                       display we need to scroll the list
    Else
        Scrollbar.Max = 0                           'There is no need for scrolling
    End If

    iMax = Scrollbar.Max
    If frmMain.Button(8).Selected Then frmMain.RANDOM_track
End Sub

Public Sub DroppedFiles(vFileList As Variant, newList As Boolean)

    Dim nLoopCtr As Integer
    Dim sFilename As String
    Dim successfulDrop As Boolean
    successfulDrop = True
    Dim INITCOUNT As Integer
    INITCOUNT = clist.ItemCount

    If newList = True Then frmPLST.ClearList
    ' *** Loop through the variant array to process
    ' each dropped file
    For nLoopCtr = 0 To UBound(vFileList)
        ' Get the current file name from the array
        sFilename = vFileList(nLoopCtr)

        If FileExists(sFilename) Then

            If cGetInputState() <> 0 Then DoEvents
            If isMediaFile(sFilename) Then
                'ipos = InStrRev(sFilename, "\")
                Call frmPLST.Add_track_to_Playlist(sFilename)
            ElseIf isPlaylistFile(sFilename) Then
                Call LoadPlaylist(sFilename, False)
            ElseIf FolderExists(sFilename) Then
                Call addFilesfromDir(sFilename, False, False)
                frmPLST.ReinitializeList
            Else
                Call frmPLST.Add_track_to_Playlist(sFilename)
            End If

        End If

    Next nLoopCtr

    If clist.ItemCount >= 1 Then
        ListRange = Fix(picList.ScaleHeight / LstTextHeight)
        If clist.ItemCount > ListRange Then
            Scrollbar.Max = clist.ItemCount - ListRange    'The list items are too much to _
                                                           display we need to scroll the list
        Else
            Scrollbar.Max = 0                           'There is no need for scrolling
        End If
        iMax = Scrollbar.Max
        Scrollbar.Value = Scrollbar.Max
        If newList = True Then
            CurrentTrack_Index = 0
            'frmMain.PlayerIsPlaying = "false"
            'Stream_Stop lCurrentChannel
            sFileMainPlaying = clist.exItem(CurrentTrack_Index)
            frmMain.Play
        End If
        ReinitializeList

    End If
End Sub


Sub cargar_formulario()
    Dim iX As Integer, iY As Integer
    Set cWindows.FormularioPadre = Me

    Set cAjustarDesk.ParentForm = Me
    cWindows.ColorInvisible = Read_INI("NormalMode", "ColorTrans", RGB(255, 0, 255), True)
    iX = Read_INI("Configuration", "ExitButtonX", 1, False, False, tAppConfig.AppConfig & "Skins\" & tAppConfig.Skin & "\PLAY_LIST\config.ini")
    iY = Read_INI("Configuration", "ExitButtonY", 1, False, False, tAppConfig.AppConfig & "Skins\" & tAppConfig.Skin & "\PLAY_LIST\config.ini")
    cWindows.ButtonExitXY CLng(iX), CLng(iY)
    cWindows.MinimoAlto = Read_INI("Configuration", "MinHeight", 10, False, False, tAppConfig.AppConfig & "Skins\" & tAppConfig.Skin & "\PLAY_LIST\config.ini")
    cWindows.MinimoAncho = Read_INI("Configuration", "MinWidth", 10, False, False, tAppConfig.AppConfig & "Skins\" & tAppConfig.Skin & "\PLAY_LIST\config.ini")
    'iButtonsLeft = Read_INI("Configuration", "ButtonsLeft", 5, False, False, tAppConfig.AppConfig & "Skins\" & tAppConfig.Skin & "\PLAY_LIST\config.ini")
    'iButtonsTop = Read_INI("Configuration", "ButtonsTop", 10, False, False, tAppConfig.AppConfig & "Skins\" & tAppConfig.Skin & "\PLAY_LIST\config.ini")
    cWindows.CargarSkin tAppConfig.AppConfig & "Skins\" & tAppConfig.Skin & "\PLAY_LIST\"

    Me.Left = cWindows.AreaLeft
    Me.Top = cWindows.AreaTop
    Me.Width = cWindows.AreaWidth
    Me.Height = cWindows.AreaHeight
    cWindows.Formulario_Paint

End Sub

Public Sub LoadSkin()
    Dim iX, iY As Integer
    Set cWindows.FormularioPadre = Me
    Set cAjustarDesk.ParentForm = Me
    cWindows.ColorInvisible = Read_INI("NormalMode", "ColorTrans", RGB(255, 0, 255), True)
    iX = Read_INI("Configuration", "ExitButtonX", 1, False, False, tAppConfig.AppConfig & "Skins\" & tAppConfig.Skin & "\PLAY_LIST\config.ini")
    iY = Read_INI("Configuration", "ExitButtonY", 1, False, False, tAppConfig.AppConfig & "Skins\" & tAppConfig.Skin & "\PLAY_LIST\config.ini")
    cWindows.MinimoAlto = Read_INI("Configuration", "MinHeight", 3, False, False, tAppConfig.AppConfig & "Skins\" & tAppConfig.Skin & "\PLAY_LIST\config.ini")
    cWindows.MinimoAncho = Read_INI("Configuration", "MinWidth", 3, False, False, tAppConfig.AppConfig & "Skins\" & tAppConfig.Skin & "\PLAY_LIST\config.ini")

    cWindows.CargarSkin tAppConfig.AppConfig & "Skins\" & tAppConfig.Skin & "\PLAY_LIST\"
    PreparelistSlider
    ReinitializeList
End Sub

Public Sub PreparelistSlider()
    Dim i, lHeight As Long
    On Error Resume Next
    PicSlidertemp.Picture = LoadPicture(tAppConfig.AppConfig & "Skins\" & tAppConfig.Skin & "\PLAY_LIST\sliderBack.bmp")
    PicSlidertemp.AutoSize = True
    picBack.Width = PicSlidertemp.Width
    lHeight = 100
    picBack.Picture = LoadPicture()
    picBack.Height = 900
    For i = 0 To 9    'paint picback completely SINCE IS OF LARGE height =900+
        picBack.PaintPicture PicSlidertemp, 0, i * lHeight, picBack.Width, lHeight, 0, 0, picBack.Width, lHeight
    Next
    picBack.Picture = picBack.Image

    PicSlidertemp.Picture = LoadPicture(tAppConfig.AppConfig & "Skins\" & tAppConfig.Skin & "\PLAY_LIST\sliderBar.bmp")
    PicSlidertemp.AutoSize = True
    picBar.Width = CInt(PicSlidertemp.Width / 3)
    picBar.Height = PicSlidertemp.Height
    picBarOver.Width = picBar.Width
    picBarDown.Width = picBar.Width
    picBarOver.Height = picBar.Height
    picBarDown.Height = picBar.Height

    picBar.Picture = LoadPicture()
    picBar.PaintPicture PicSlidertemp, 0, 0, picBar.Width, picBar.Height, 0, 0, picBar.Width, picBar.Height
    picBar.Picture = picBar.Image

    picBarOver.Picture = LoadPicture()
    picBarOver.PaintPicture PicSlidertemp, 0, 0, picBar.Width, picBar.Height, picBar.Width, 0, picBar.Width, picBar.Height
    picBarOver.Picture = picBarOver.Image

    picBarDown.Picture = LoadPicture()
    picBarDown.PaintPicture PicSlidertemp, 0, 0, picBar.Width, picBar.Height, 2 * picBar.Width, 0, picBar.Width, picBar.Height
    picBarDown.Picture = picBarDown.Image
    picBack.Left = picList.Width - picBack.Width
    ReinitializeList

End Sub

Public Sub Update_track_in_Playlist(Index As Long)
'Usage: it reads Mpeg Info and Id3 tag Info and replaces playlist entry with this informaton
'       It is called by timer to get system some breath. All the tracks are initially loaded by their
'       filename and then the timer ins initialised which calls it if track is not already updated
'updated: 16.2.11

    Dim sFullpath As String, sTrackName As String, sFileEx As String
    Dim sFormat As String
    Dim cfile As New cMP3
    'On Error GoTo error_handler
    Dim sTrackpath As String
    If Index >= clist.ItemCount Then Exit Sub

    sTrackpath = clist.exItem(Index)
    'If filedoes not exist then length=00:00 so that it is confirmed to be already read else it would
    ' be updated again, since lenght="" is used as flag to check where track is updated or not
    Call clist.SetTracklength(Index, "00:00")
    If FileExists(sTrackpath) = False Then: Exit Sub
    cfile.Read_MPEGInfo = True
    cfile.Read_File_Tags sTrackpath
    Call clist.SetTracklength(Index, IIf(cfile.MPEG_DurationTime <> "", cfile.MPEG_DurationTime, "00:00"))

    ' load tags
    sTrackName = GetFileTitle(sTrackpath)
    sFileEx = Right(sFullpath, 3)

    sFormat = sFormatPlayList
    '// Song Name
    If Trim(cfile.Title) = "" Then cfile.Title = sTrackName
    sFormat = Replace(sFormatPlayList, "%S", Trim(cfile.Title))
    '// Artist
    sFormat = Replace(sFormat, "%A", Trim(cfile.Artist))
    '// Album
    sFormat = Replace(sFormat, "%B", Trim(cfile.Album))
    '// Year
    sFormat = Replace(sFormat, "%Y", Trim(cfile.Year))
    '// Genre
    sFormat = Replace(sFormat, "%G", Trim(cfile.Genre))
    '// Time
    ' sFormat = Replace(sFormat, "%T", Trim(cFile.MPEG_DurationTime))
    '// File Name
    sFormat = Replace(sFormat, "%N", sTrackName)
    '// Time
    sFormat = Replace(sFormat, "%P", sTrackpath)
    '// File extencion
    sFormat = Replace(sFormat, "%F", sFileEx)

    If sFormat = sFormatPlayList Then sFormat = sTrackName

    '------------------------------------------------------------------------------
    sFormat = UpperCase_Firstletter(sFormat)    'for better appearance

    Call clist.ChangeItem(Index, sFormat)
    Exit Sub

error_handler:
    'Set duration to be zero if error in reading mpeginfo
    Call clist.SetTracklength(Index, "00:00")

End Sub

Public Sub Add_track_to_Playlist(sTrackpath As String, Optional Index As Long = -1)
    If Index = -1 Then Index = clist.ItemCount
    clist.AddItem GetFileTitle(sTrackpath), sTrackpath, "", , Index
    If Index <= CurrentTrack_Index Then CurrentTrack_Index = CurrentTrack_Index + 1
    If Index <= NormalSelection Then NormalSelection = NormalSelection + 1
    If Index <= PreviousSelection Then PreviousSelection = PreviousSelection + 1

End Sub


Private Sub TmrList_Timer()
    Dim lLimit As Long
    lLimit = Min(clist.ItemCount - 1, ListRange - 1)
    'Update only those tracks which are not already updated i.e. do not contain ID3 Info
    'It is checked by checking length entry of track in clist
    If lCounter >= Scrollbar.Value + lLimit Or clist.ItemCount = 0 Then TmrList.Enabled = False: Exit Sub

    Do While (Trim(clist.exTracklength(lCounter)) <> "" And lCounter <= Scrollbar.Value + lLimit)
        ' Debug.Print clist.exTracklength(lCounter)
        If lCounter = Scrollbar.Value + lLimit Then TmrList.Enabled = False: Exit Sub
        lCounter = lCounter + 1    ' this couter confirms no need of updation we can skip to next track
    Loop
    Update_track_in_Playlist (lCounter)
    'replace old track by new updated track, we do not need to reinilialize whole list
    EraseTrack (lCounter)
    Drawtrack (lCounter)
End Sub

Private Sub tmrMove_Timer()
'It fires event mousemove when track is dragged and cursor reaches above top of playlist or below
'bottom of list even cursor stays(does't move) at these position we should continue dragging but since
'mouse is not moving we manually have to fire mousemove even using this timer
    Dim udtCursor As POINTAPI    'stores mouse coordiantes in screen frame
    Dim udtRC As POINTAPI    'to store mouse coordiantes in piclist frame

    GetCursorPos udtCursor
    ClientToScreen picList.hwnd, udtRC    'gives topleft position of piclist in screen frame
    'fire mouse move event by sending coordinates in piclist frmae
    picList_MouseMove vbLeftButton, 0, udtCursor.X - udtRC.X, udtCursor.Y - udtRC.Y

End Sub

Private Function addFilesfromDir(Path As String, Optional SubFolder As Boolean = False, Optional bNewSearch As Boolean = False, Optional Index As Long = -1) As Long
'Return: No of files added

    Dim WFD As WIN32_FIND_DATA
    Static lFilecount As Long
    Dim hFile As Long, fPath As String, fName As String
    If bNewSearch Then iCount = 0

    fPath = AddBackSlash(Path)
    fName = fPath & "*.mp3"   'Use fPath & "*.*" for all files

    'GET HaNDLE OF FIRST FILE.
    hFile = FindFirstFile(fName, WFD)
    If (hFile > 0) And ((WFD.dwFileAttributes And FILE_ATTRIBUTE_DIRECTORY) <> FILE_ATTRIBUTE_DIRECTORY) Then
        If LCase(Right(fPath & StripNulls(WFD.cFileName), 3)) = LCase("mp3") Then
            Call Add_track_to_Playlist(GetLongFilename(fPath & StripNulls(WFD.cFileName)), Index)
            lFilecount = lFilecount + 1
            If boolSearchShow = True Then iCount = iCount + 1: frmSearch.lblFileCount.Caption = "FILES ADDED [" + str(iCount) + "]"
        End If
    End If

    While FindNextFile(hFile, WFD)
        'Solange "FindNextFile" ausfuehren, bis keine Datei mehr gefunden wird, also hFile 0 ist.
        If ((WFD.dwFileAttributes And FILE_ATTRIBUTE_DIRECTORY) <> FILE_ATTRIBUTE_DIRECTORY) Then
            If LCase(Right(fPath & StripNulls(WFD.cFileName), 3)) = LCase("mp3") Then
                Call Add_track_to_Playlist(GetLongFilename(fPath & StripNulls(WFD.cFileName)), Index)
                lFilecount = lFilecount + 1
                If boolSearchShow = True Then iCount = iCount + 1: frmSearch.lblFileCount.Caption = "FILES ADDED [" + str(iCount) + "]"
            End If
        End If
    Wend

    If SubFolder Then

        hFile = FindFirstFile(fName, WFD)
        If (hFile > 0) And ((WFD.dwFileAttributes And FILE_ATTRIBUTE_DIRECTORY) = FILE_ATTRIBUTE_DIRECTORY) And _
           StripNulls(WFD.cFileName) <> "." And StripNulls(WFD.cFileName) <> ".." Then

            addFilesfromDir fPath & StripNulls(WFD.cFileName), True, , Index
            DoEvents
        End If

        While FindNextFile(hFile, WFD)
            If ((WFD.dwFileAttributes And FILE_ATTRIBUTE_DIRECTORY) = FILE_ATTRIBUTE_DIRECTORY) And _
               StripNulls(WFD.cFileName) <> "." And StripNulls(WFD.cFileName) <> ".." Then

                DoEvents
                addFilesfromDir fPath & StripNulls(WFD.cFileName), True, , Index
            End If
        Wend

    End If
    FindClose hFile

    addFilesfromDir = lFilecount
    Debug.Print lFilecount
    lFilecount = 0

    '  bSearching = False
End Function

Public Sub Update_Plst_Scrollbar()
    If frmPLST.clist.ItemCount >= 1 Then
        ListRange = Fix(frmPLST.picList.ScaleHeight / LstTextHeight)
        If frmPLST.clist.ItemCount > ListRange Then
            frmPLST.Scrollbar.Max = frmPLST.clist.ItemCount - ListRange    'The list items are too much to _
                                                                           display we need to scroll the list
        Else
            frmPLST.Scrollbar.Max = 0                           'There is no need for scrolling
        End If
        iMax = frmPLST.Scrollbar.Max

        ReinitializeList
        iValue = Scrollbar.Max - (Scrollbar.Value)
        DrawBar (Normal)
    Else
        frmPLST.Scrollbar.Value = 0: frmPLST.Scrollbar.Max = 0:
    End If
End Sub
Public Sub UpdatePlaylist()
    Dim sFullpath As String, sTrackName As String, sFileEx As String, sTrackpath As String
    Dim sCleanStr As String, sNewString As String, sFormat As String
    Dim sSplitField() As String
    Dim iSpaces As Integer
    Dim cfile As New cMP3


    Dim i As Integer
    For i = 0 To clist.ItemCount - 1
        sTrackpath = clist.exItem(i)

        cfile.Read_MPEGInfo = True
        cfile.Read_File_Tags sTrackpath

        ' load tags
        sTrackName = GetFileTitle(sTrackpath)
        sFileEx = Right(sFullpath, 3)

        sFormat = sFormatPlayList
        '// Song Name
        ' If Trim(cFile.Title) = "" Then cFile.Title = sTrackName
        sFormat = Replace(sFormatPlayList, "%S", Trim(cfile.Title))
        '// Artist
        sFormat = Replace(sFormat, "%A", Trim(cfile.Artist))
        '// Album
        sFormat = Replace(sFormat, "%B", Trim(cfile.Album))
        '// Year
        sFormat = Replace(sFormat, "%Y", Trim(cfile.Year))
        '// Genre
        sFormat = Replace(sFormat, "%G", Trim(cfile.Genre))
        '// Time
        ' sFormat = Replace(sFormat, "%T", Trim(cFile.MPEG_DurationTime))
        '// File Name
        sFormat = Replace(sFormat, "%N", sTrackName)
        '// Time
        sFormat = Replace(sFormat, "%P", sTrackpath)
        '// File extencion
        sFormat = Replace(sFormat, "%F", sFileEx)

        If sFormat = sFormatPlayList Then sFormat = sTrackName

        '------------------------------------------------------------------------------
        sCleanStr = Trim$(sFormat)

        'Upper case and / or lower case the string correctly.
        sSplitField = Split(sCleanStr, " ", , vbTextCompare)
        sCleanStr = ""
        ' 'Debug.Print sCleanStr
        For iSpaces = 0 To UBound(sSplitField)
            If (Not iSpaces = 0 Or Not IsNumeric(sSplitField(iSpaces))) And sSplitField(iSpaces) <> "" Then
                sNewString = UCase$(Left$(sSplitField(iSpaces), 1))
                sNewString = sNewString & (Right$(sSplitField(iSpaces), Len(sSplitField(iSpaces)) - 1))
                sCleanStr = sCleanStr & sNewString & " "
            End If
        Next iSpaces
        sFormat = Trim$(sCleanStr)
        '------------------------------------------------------------------------------

        clist.ChangeItem i, sFormat
        clist.ChangeEXTracklength i, cfile.MPEG_DurationTime
    Next
    ReinitializeList
End Sub

Public Sub EnsureTrackVisible(lTrackindex As Integer)
    Dim i As Integer
    If lTrackindex < 0 Or lTrackindex >= clist.ItemCount Then Exit Sub
    If Scrollbar.Max = 0 Or (lTrackindex >= Scrollbar.Value And lTrackindex <= (Scrollbar.Value + ListRange)) Then
        EraseTrack (lTrackindex)
        Drawtrack (lTrackindex)
        Exit Sub
    End If
    If lTrackindex < Scrollbar.Value Then    ' currenttrack lies above the top of list
        i = maxX(0, lTrackindex - ListRange / 2)
        If i <= Scrollbar.Max Then
            Scrollbar.Value = i
        Else
            Scrollbar.Value = Scrollbar.Max
        End If
    ElseIf lTrackindex >= (Scrollbar.Value + ListRange) Then    '' currenttrack lies below the bottom of list
        i = Min(Scrollbar.Max, lTrackindex - ListRange / 2)
        If i >= 0 Then Scrollbar.Value = i
    End If
    ReinitializeList
End Sub

Public Function HitTestPlst(X As Single, Y As Single) As Long
    HitTestPlst = -1
    If Fix(Y / LstTextHeight) + Scrollbar.Value >= 0 And Fix(Y / LstTextHeight) + Scrollbar.Value < clist.ItemCount Then
        HitTestPlst = Fix(Y / LstTextHeight) + Scrollbar.Value
    End If
End Function


Public Sub SavePlaylist(sOutputPath As String)
'SAVE The Playlist in format m3u,pls or npl
    sOutputPath = StripNulls(sOutputPath)
    Dim sExtension As String
    'Get file extension
    If frmPLST.clist.ItemCount = 0 Then MsgBox "An Empty Playlist can not be saved": Exit Sub

    sExtension = Trim(LCase(Right(sOutputPath, Len(sOutputPath) - InStrRev(sOutputPath, "."))))
    Dim i As Long

    Select Case sExtension
    Case "m3u"
        Close #1
        Open sOutputPath For Output As #1
        Print #1, "#EXTM3U"    '// m3u header
        For i = 0 To frmPLST.clist.ItemCount - 1
            Print #1, frmPLST.clist.exItem(i)
        Next
        Close #1

    Case "pls"
        Open sOutputPath For Output As #2
        Print #2, "[playlist]"

        For i = 0 To frmPLST.clist.ItemCount - 1
            Print #2, "File" & i + 1 & "=" & frmPLST.clist.exItem(i)  'print the file's path from playlist
            Print #2, "Title" & i + 1 & "=" & frmPLST.clist.Item(i)
            Print #2, "Length" & i + 1 & "=" & Convert_TextTime_to_Seconds(frmPLST.clist.exTracklength(i))
        Next
        Print #2, "NumberOfEntries=" & i
        Print #2, "Version=2"

        Close #2

    Case "npl"
        Dim tTrack As FileTrack
        On Error Resume Next
        If FileExists(sOutputPath) Then Kill (sOutputPath)
        Open sOutputPath For Random As #3 Len = 255
        For i = 0 To frmPLST.clist.ItemCount - 1
            tTrack.trackName = frmPLST.clist.Item(i)
            tTrack.trackPath = frmPLST.clist.exItem(i)
            tTrack.Duration = frmPLST.clist.exTracklength(i)
            Put #3, i + 1, tTrack
        Next
        Close #3
    End Select

End Sub

Private Sub Drawtrack(i As Integer)
    Dim m As RECT
    'Draw only when track is visible
    If i < Scrollbar.Value Or i > Scrollbar.Value + (ListRange - 1) Then Exit Sub

    m.Left = 1
    m.Right = picList.ScaleWidth - picList.TextWidth("00:00") - picBack.Width - 7
    m.Top = (i - Scrollbar.Value) * LstTextHeight
    m.Bottom = m.Top + LstTextHeight
    'set color of track to be displayed
    picList.ForeColor = IIf(i = CurrentTrack_Index, CurrentTrack_Forecolor, NormalText_Forecolor)
    If clist.bSelected(i) = True Then DrawSelection i


    DrawText picList.hDC, " " & i + 1 & ". " & clist.Item(i), Len(" " & i + 1 & ". " & clist.Item(i)), m, DT_VCENTER Or DT_SINGLELINE Or DT_END_ELLIPSIS Or DT_LEFT

    m.Left = m.Right + 4
    m.Right = m.Right + picList.TextWidth("00:00") + 4
    'draw tracklength
    DrawText picList.hDC, clist.exTracklength(i), Len(clist.exTracklength(i)), m, DT_VCENTER Or DT_SINGLELINE Or DT_LEFT



End Sub




