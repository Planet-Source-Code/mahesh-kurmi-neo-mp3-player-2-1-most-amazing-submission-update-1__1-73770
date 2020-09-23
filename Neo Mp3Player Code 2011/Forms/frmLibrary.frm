VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL32.ocx"
Begin VB.Form frmLibrary 
   BackColor       =   &H80000005&
   Caption         =   "NeoPlayer Media Library"
   ClientHeight    =   8220
   ClientLeft      =   225
   ClientTop       =   855
   ClientWidth     =   12825
   Icon            =   "frmLibrary.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   8220
   ScaleWidth      =   12825
   StartUpPosition =   3  'Windows Default
   Begin MMPlayerXProject.vkCommand cmdSearch 
      Height          =   330
      Left            =   8970
      TabIndex        =   7
      Top             =   120
      Width           =   1005
      _extentx        =   1773
      _extenty        =   582
      caption         =   "Search..."
      font            =   "frmLibrary.frx":000C
   End
   Begin VB.FileListBox FileDropList 
      Height          =   870
      Left            =   10320
      Pattern         =   "*.mp3"
      TabIndex        =   5
      Top             =   1560
      Visible         =   0   'False
      Width           =   1935
   End
   Begin MSComctlLib.ImageList ImgIconos 
      Left            =   3450
      Top             =   6660
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   15
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLibrary.frx":0034
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLibrary.frx":0A46
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLibrary.frx":0DE0
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLibrary.frx":17F2
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLibrary.frx":39D6
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLibrary.frx":3D70
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLibrary.frx":430A
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLibrary.frx":48A4
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLibrary.frx":4C3E
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLibrary.frx":51D8
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLibrary.frx":5772
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLibrary.frx":5D0C
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLibrary.frx":62A6
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLibrary.frx":6640
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLibrary.frx":69DA
            Key             =   """Note"""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   315
      Left            =   0
      TabIndex        =   0
      Top             =   7905
      Width           =   12825
      _ExtentX        =   22622
      _ExtentY        =   556
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   19526
            Picture         =   "frmLibrary.frx":6E2C
            Text            =   "Neo Player: Media Library"
            TextSave        =   "Neo Player: Media Library"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Picture         =   "frmLibrary.frx":73C6
            Text            =   "Records: "
            TextSave        =   "Records: "
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.PictureBox PicClientArea 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   6495
      Left            =   30
      ScaleHeight     =   6465
      ScaleWidth      =   9975
      TabIndex        =   1
      Top             =   30
      Width           =   10005
      Begin MMPlayerXProject.ListView ListView2 
         Height          =   4245
         Left            =   2880
         TabIndex        =   6
         Top             =   480
         Width           =   7035
         _extentx        =   12409
         _extenty        =   7488
         forecolor       =   8421504
         reorder         =   0   'False
         font            =   "frmLibrary.frx":7960
         oledropmode     =   1
         checkboxcolor   =   4210752
         picturewidth    =   16
         pictureheight   =   16
      End
      Begin VB.PictureBox picSplitMain 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000C&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   4000
         Left            =   2790
         ScaleHeight     =   4005
         ScaleMode       =   0  'User
         ScaleWidth      =   45
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   480
         Visible         =   0   'False
         Width           =   45
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   300
         Left            =   150
         TabIndex        =   2
         Top             =   90
         Width           =   8490
      End
      Begin MSComctlLib.TreeView TreeFiles 
         Height          =   4245
         Left            =   150
         TabIndex        =   3
         Top             =   480
         Width           =   2685
         _ExtentX        =   4736
         _ExtentY        =   7488
         _Version        =   393217
         HideSelection   =   0   'False
         Indentation     =   176
         LabelEdit       =   1
         LineStyle       =   1
         Style           =   7
         ImageList       =   "ImgIconos"
         BorderStyle     =   1
         Appearance      =   0
         MousePointer    =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Image imgSplitMain 
         Height          =   4260
         Left            =   2790
         MousePointer    =   9  'Size W E
         Top             =   480
         Width           =   45
      End
   End
   Begin MSComctlLib.ImageList imgLstMp3 
      Left            =   5100
      Top             =   6630
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLibrary.frx":7988
            Key             =   "Notesss"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLibrary.frx":7D22
            Key             =   "List11"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLibrary.frx":8734
            Key             =   "List"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLibrary.frx":8ACE
            Key             =   "Note"
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Begin VB.Menu mnuNewplaylist 
         Caption         =   "&New Playlist"
      End
      Begin VB.Menu mnuImportCurrentlist 
         Caption         =   "Import Current Playlist"
      End
      Begin VB.Menu mnuImportFile 
         Caption         =   "Import Playlist from file"
      End
      Begin VB.Menu mnubar8 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExportList 
         Caption         =   "Export Playlist as.."
      End
      Begin VB.Menu mnuAddmedia 
         Caption         =   "&Add media to library.."
         Index           =   1
      End
      Begin VB.Menu mnubar2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu mnuview 
      Caption         =   "View"
      Begin VB.Menu mnurefresh 
         Caption         =   "Refresh"
         Shortcut        =   {F5}
      End
      Begin VB.Menu mnuFavTracks 
         Caption         =   "Favourite Tracks"
      End
      Begin VB.Menu mnuHistory 
         Caption         =   "History"
      End
      Begin VB.Menu mnubar 
         Caption         =   "-"
      End
      Begin VB.Menu mnuRemovemissing 
         Caption         =   "Remove missing files from Library"
      End
      Begin VB.Menu mnuEraseLibrary 
         Caption         =   "Remove All Entries from Library"
      End
      Begin VB.Menu mnubar9 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPreferences 
         Caption         =   "Preferences"
      End
   End
   Begin VB.Menu mnuhelp 
      Caption         =   "Help"
      Begin VB.Menu mnuplayerhelp 
         Caption         =   "Player help"
      End
      Begin VB.Menu mnubar4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAbout 
         Caption         =   "About"
      End
   End
   Begin VB.Menu mnuPopupListview 
      Caption         =   "PopupListview"
      Begin VB.Menu mnuPlay 
         Caption         =   "Play Selection as new Playlist"
      End
      Begin VB.Menu mnuEnqueue 
         Caption         =   "Enqueue Selection"
      End
      Begin VB.Menu mnuSend 
         Caption         =   "Send to"
         Begin VB.Menu mnuSendtoCurrntlist 
            Caption         =   "Neo player's Playlist"
         End
         Begin VB.Menu bar12 
            Caption         =   "-"
         End
         Begin VB.Menu mnuPlstname 
            Caption         =   "Playlist: New Playlist"
            Index           =   0
         End
      End
      Begin VB.Menu mnubar7 
         Caption         =   "-"
      End
      Begin VB.Menu mnuremove 
         Caption         =   "Remove Item(s)"
      End
      Begin VB.Menu mnubar6 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExplore 
         Caption         =   "Explore item(s) Folder"
      End
      Begin VB.Menu mnuviewtag 
         Caption         =   "View/Edit Tag info.."
      End
   End
End
Attribute VB_Name = "frmLibrary"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'DefLng A-Z

'/////////////////////////////////////
'The following declarations are for subclassing

Implements ICustomDraw
Implements ISubclass
Private m_clsSubclass As Subclasser
Private Const WM_SIZE As Long = &H5
Private Const WM_SIZING As Long = &H214
Private Const WM_EXITSIZEMOVE As Long = &H232
Private Const WM_GETMINMAXINFO = &H24
Private Const WM_MOVING As Long = &H216

Private XPos As Long
Private YPos As Long

Private XSize As Long
Private YSize As Long


Private Type POINTAPI
    X As Long
    Y As Long
End Type

Private Type MINMAXINFO
    ptReserved As POINTAPI
    ptMaxSize As POINTAPI
    ptMaxPosition As POINTAPI
    ptMinTrackSize As POINTAPI
    ptMaxTrackSize As POINTAPI
End Type

Private Declare Sub CopyMemoryToMinMaxInfo Lib "kernel32" Alias _
                                           "RtlMoveMemory" (hpvDest As MINMAXINFO, ByVal hpvSource As Long, _
                                                            ByVal cbCopy As Long)
Private Declare Sub CopyMemoryFromMinMaxInfo Lib "kernel32" Alias _
                                             "RtlMoveMemory" (ByVal hpvDest As Long, hpvSource As MINMAXINFO, _
                                                              ByVal cbCopy As Long)
'///////////////////////////////////////////////////////////////


'////////////////////////////////////////////////////////////
'//Playlist and popupmenu handling
Private Type tPlaylist  ' Stores playlist info
    Name As String
    TrackCount As Integer
End Type

Private Enum enumPopuptype  'for origin of popupMenu window
    enormallistview = 0
    ePlaylistview = 1
    enormaltreeview = 2
    eplaylisttreeview = 3
End Enum

Dim enumPopup As enumPopuptype
Dim tPlaylists() As tPlaylist
Dim bPlaylistActive As Boolean
Dim iActivePlaylist As Integer
Dim iActiveKeyIndex As Integer
Dim bPlaylistEdited As Boolean    'tells if playlist items are reordered or changed so that database must be updated
'/////////////////////////////////////////////////////////////

'////////////////////////////////////////////////
'/DataBase
Dim strSQL As String
Dim sConnectionString As String
Dim SQL As New ADODB.connection
Dim rs As New ADODB.Recordset

Public cnnMusic As ADODB.connection
Dim CMD As ADODB.Command
'///////////////////////////////////////


Dim mLMouseDown As Boolean
Dim sSelectedTreeKey As String

Dim bOLEdragging As Boolean
Private iDragitemIndex As Long
Private m_lngLastDropItem As Long

Dim myWindowState As Integer
Dim bInitialized As Boolean
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

'***************************************************************************************

' Name: StopFlicker
' Description:Avoid the Flickering
' Use this routine to stop a control (like a list or treeview) from flickering when it is getting it's data.
' By: Strider Solutions
' I may be using it later on if find necessary
' This code is copyrighted and has' limited warranties.Please see http://www.Planet-Source-Code.com/vb/scripts/ShowCode.asp?txtCodeId=4292&lngWId=1'for details.'**************************************

' http://www.stridersolutions.com/products/cs/

'Private Declare Function LockWindowUpdate Lib "user32" (ByVal hwndLock As Long) As Long

'Private Sub StopFlicker(ByVal lhWnd As Long)
'Dim lRet As Long
''Object will not flicker - just be blank
'lRet = LockWindowUpdate(lhWnd)
'End Sub
'Private Sub Release()
'Dim lRet As Long
'lRet = LockWindowUpdate(0)
'End Sub
'***************************************************************************************


'////CORE PART of LIBRARY: This routine Binds Listview controlwith database table
'i was very excited when i did this because i used database for the first time in any code

Public Function BindToSQL(strSQL As String, Optional iSQL As ADODB.connection, Optional iRs As ADODB.Recordset, Optional bPlaylistActive As Boolean = False)
    On Error Resume Next
    'On Error GoTo err_handler:
    Dim lWidth As Long
    If iSQL Is Nothing Then
        Set iSQL = SQL
        If iSQL.state > 0 Then iSQL.Close
        Dim connstring As String
        connstring = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & tAppConfig.AppConfig & "\Library\music.mdb;Persist Security Info=False"
        iSQL.Open connstring
    End If
    If strSQL = "" Then Exit Function

    If iRs Is Nothing Then Set iRs = rs
    If iRs.state > 0 Then iRs.Close
    iRs.Open strSQL, iSQL, adOpenKeyset, adLockOptimistic

    ListView2.Redraw = False
    ListView2.Clear
    ' listview2.
    Dim i As Integer
    For i = 0 To ListView2.ColumnCount - 1
        ListView2.RemoveColumn (0)
    Next

    Dim sF As Field

    i = 0
    lWidth = ListView2.Width / Screen.TwipsPerPixelX
    If bPlaylistActive = False Then
        For Each sF In iRs.Fields
            i = i + 1
            'ListView2.AddColumn str(i), sF.Name, , , , True
            Select Case i
            Case 1:
                ListView2.AddColumn str(i - 1), sF.Name, , 1.9 * lWidth / 9, , True

            Case 2:
                ListView2.AddColumn str(i - 1), sF.Name, , 1.2 * lWidth / 9, , True

            Case 5:
                ListView2.AddColumn str(i - 1), sF.Name, 2, 0.5 * lWidth / 9, , True

            Case 6:
                ListView2.AddColumn str(i - 1), sF.Name, 2, 0.6 * lWidth / 9, , True

            Case 7:
                ListView2.AddColumn str(i - 1), sF.Name, 2, 0.8 * lWidth / 9, , True

            Case 9:
                ListView2.AddColumn str(i - 1), sF.Name, , 2 * lWidth / 9, , True

            Case 3, 4, 8:
                ListView2.AddColumn str(i - 1), sF.Name, , lWidth / 9, , True

            End Select
        Next
    Else
        For Each sF In iRs.Fields
            i = i + 1
            Select Case i
                ' ListView2.AddColumn str(i), sF.Name, , , , True
            Case 1:    ' 0.3 * ListView2.Width / 4
                ListView2.AddColumn str(i), sF.Name, , 55, False, True
            Case 2:
                ListView2.AddColumn str(i), sF.Name, , 0.55 * (lWidth - 113), , True
            Case 3:
                ListView2.AddColumn str(i), sF.Name, 2, 58, False, True
            Case 4:
                ListView2.AddColumn str(i), sF.Name, , 0.45 * (lWidth - 113), True

            End Select
        Next
    End If
    Dim lngItem As Long


    Dim X As Integer, Item
    Do While Not iRs.EOF
        With ListView2
            lngItem = .AddItem(, , 0)
            For X = 0 To iRs.Fields.Count - 1
                If iRs(X).Value <> "" Then .ItemText(lngItem, X) = iRs(X).Value
            Next
        End With
        iRs.MoveNext
        If cGetInputState() <> 0 Then DoEvents
    Loop

    ' If ListView2.ItemCount > 0 Then ListView2.ItemSelected(0) = True
    ListView2.Redraw = True

    Exit Function

err_handler:
    ListView2.Redraw = True

    MsgBox "Error binding ListView to SQL Mahesh" & vbNewLine & vbNewLine & "Error code:" & err.Number & vbNewLine & "Error desc:" & err.Description, vbCritical
End Function

Private Sub cmdSearch_Click()

    Dim sCampos As String
    Dim sSQL As String
    Dim sWhere As String
    On Error GoTo err_handler:
    If Len(Trim(Text1)) = 0 Then Exit Sub
    bPlaylistActive = False    'IMPortant: since ttracks are displayed in normal format
    sCampos = "TITLE,ARTIST,ALBUM,GENRE,YEAR,LENGTH,PLAYCOUNT,PLAYEDLAST,FILE"
    sSQL = Replace(Text1, "'", "''", , , vbTextCompare)
    sSQL = "SELECT " & sCampos & " FROM MUSIC WHERE TITLE LIKE '%" & sSQL & "%' OR ARTIST LIKE '%" & sSQL & "%' ORDER BY TITLE"
    BindToSQL sSQL
    rs.Close
    sSQL = Replace(Text1, "'", "''", , , vbTextCompare)
    sWhere = "WHERE TITLE LIKE '%" & sSQL & "%' OR ARTIST LIKE '%" & sSQL & "%'"    '' ORDER BY TITLE"

    rs.Open "SELECT SUM(BYTES) AS TOTAL, SUM(SECONDS) AS TIEMPO FROM MUSIC " & sWhere, cnnMusic, adOpenForwardOnly, adLockReadOnly

    Dim lKilobytes As Long, lSeconds As Long
    lKilobytes = CLng(rs!total / 1024)
    lSeconds = CLng(rs!TIEMPO)
    rs.Close

    '// UPDATE STATUS BAR
    Call UpdateStatusBar(lKilobytes, lSeconds)
    Exit Sub

err_handler:
    MsgBox "Error binding ListView to SQL" & vbNewLine & vbNewLine & "Error code:" & err.Number & vbNewLine & "Error desc:" & err.Description, vbCritical

End Sub

Private Sub Form_Initialize()
    Set m_clsSubclass = New Subclasser
    If m_clsSubclass.Subclass(frmLibrary.hwnd, Me) Then
        m_clsSubclass.AddMsg frmLibrary.hwnd, WM_SIZE
        m_clsSubclass.AddMsg frmLibrary.hwnd, WM_SIZING
        m_clsSubclass.AddMsg frmLibrary.hwnd, WM_EXITSIZEMOVE
        m_clsSubclass.AddMsg frmLibrary.hwnd, WM_GETMINMAXINFO
    End If
End Sub

Private Sub Form_Load()
    On Error Resume Next
    Dim hBitmap As Long
    Dim i As Integer
    ReDim tPlaylists(0 To 0)

    ' Change the formicon to display Alpha Icons from resource file
    Call SetIcon(Me.hwnd, "LISTICON", False)

    Dim strRes
    strRes = Read_INI("Configuration", "LX", 6225, , True)
    If IsNumeric(strRes) = False Then strRes = 0
    frmLibrary.Left = CInt(strRes)

    strRes = Read_INI("Configuration", "LY", 5815, , True)
    If IsNumeric(strRes) = False Then strRes = 0
    frmLibrary.Top = CInt(strRes)

    strRes = Read_INI("Configuration", "LW", 8250, , True)
    frmLibrary.Width = CInt(strRes)

    strRes = Read_INI("Configuration", "LH", 4760, , True)
    frmLibrary.Height = CInt(strRes)


    bInitialized = False
    Set ListView2.DrawCallback = Me
    ListView2.AddPicture App.Path & "\INN.ico"
    m_lngLastDropItem = -1

    imgSplitMain.Left = TreeFiles.Left + TreeFiles.Width
    picSplitMain.Left = imgSplitMain.Left
    imgSplitMain.Top = TreeFiles.Top
    picSplitMain.Top = TreeFiles.Top
    ListView2.Top = TreeFiles.Top
    ListView2.Left = imgSplitMain.Left + imgSplitMain.Width

    bInitialized = True
    mnuPopupListview.Visible = False
    Form_Resize

    Set cnnMusic = New ADODB.connection

    With cnnMusic
        .Provider = "Microsoft.Jet.OLEDB.4.0"
        .Properties("Data Source") = tAppConfig.AppConfig & "Library\music.mdb"
        '.Properties("Jet OLEDB:Database Password") = "Licenciao159"
        'Debug.Print tAppConfig.AppConfig & "Library\music.mdb"
        .CursorLocation = adUseClient
        .Open
    End With
    Set CMD = New ADODB.Command
    CMD.ActiveConnection = cnnMusic
    LoadLibrary (True)


End Sub


Private Sub Form_Resize()
    On Error Resume Next
    ListView2.ColumnsAutoSize = True
    If WindowState <> vbMinimized Then
        If bInitialized Then myWindowState = WindowState
        PicClientArea.Width = ScaleWidth + 15    '* Screen.TwipsPerPixelX
        PicClientArea.Height = ScaleHeight - 340    '* Screen.TwipsPerPixelY
        Text1.Width = PicClientArea.ScaleWidth - 18 * Screen.TwipsPerPixelX - cmdSearch.Width
        cmdSearch.Left = PicClientArea.ScaleWidth - cmdSearch.Width - 3 * Screen.TwipsPerPixelX
        ListView2.Width = PicClientArea.ScaleWidth - ListView2.Left - 6 * Screen.TwipsPerPixelX
        ListView2.Height = PicClientArea.ScaleHeight - 30 * Screen.TwipsPerPixelY
        imgSplitMain.Height = ListView2.Height
        picSplitMain.Height = ListView2.Height
        TreeFiles.Height = ListView2.Height
    End If
End Sub



Private Sub Form_Unload(Cancel As Integer)
    frmLibrary.Visible = False
    If Not PlayerisClosing Then boolMediaLibraryShow = False: frmPopUp.mnuLibrary.Checked = boolMediaLibraryShow
    If PlayerisClosing = True Then
        cnnMusic.Close
        Set cnnMusic = Nothing
        m_clsSubclass.UnSubclass frmLibrary.hwnd
        ' libUnhook
    End If
    Cancel = Not PlayerisClosing
End Sub

Private Function ICustomDraw_CustomDraw(ByVal ItemIndex As Long, ByVal ColumnIndex As Long, BackColor As Long, ForeColor As Long) As Boolean
' Reconstruction of a complex Foobar2000 Styles
    Dim blnSelected As Boolean
    Dim lngTag As Long

    blnSelected = ListView2.ItemSelected(ItemIndex)
    lngTag = ListView2.ItemTag(ItemIndex)


    ForeColor = &H777777    '555555 '88888

    If bPlaylistActive Then
        If ColumnIndex = 0 Then ListView2.ItemText(ItemIndex, 0) = CStr(ItemIndex + 1)
        If ItemIndex Mod 2 Then
            BackColor = &HFAFAFA
            If blnSelected Then BackColor = &HFEECB4
        Else
            BackColor = &HF7F7F7
            If blnSelected Then BackColor = &HFEEFBF

        End If
        If blnSelected Then ForeColor = &HD39725

        Select Case ColumnIndex
            ' Index, Length

        Case 0, 2
            If ItemIndex Mod 2 Then
                BackColor = &HE8E8E8
                If blnSelected Then BackColor = &HF5DEAD
            Else
                BackColor = &HE5E5E5
                If blnSelected Then BackColor = &HF8E3BA
            End If
        End Select
        ICustomDraw_CustomDraw = True
        Exit Function
    End If

    If lngTag = 1 Then
        ForeColor = vbWhite
        BackColor = &H1AAFF
        If blnSelected Then BackColor = &HF9D577
    Else
        ' ForeColor = &H777777 '55555 '&H0      '&H888888
        If blnSelected Then ForeColor = &HD39725    '&HD39725

        If ItemIndex Mod 2 Then
            BackColor = &HFAFAFA
            If blnSelected Then BackColor = &HFEECB4
        Else
            BackColor = &HF7F7F7
            If blnSelected Then BackColor = &HFEEFBF
        End If
    End If

    Select Case ColumnIndex

        ' Index, Länge
    Case 0, 1, 2, 3, 4, 5, 6, 7, 8:
        If lngTag = 1 Then
            BackColor = &H19CE8
            If blnSelected Then BackColor = &HEEBF33
        Else
            If ItemIndex Mod 2 Then
                BackColor = &HFEEEEF   '&HE8E8E8
                If blnSelected Then BackColor = &HF5DEAD
            Else
                BackColor = &HEFFFFF  '&&HE5E5E5
                If blnSelected Then BackColor = &HF8E3BA
            End If
        End If

    End Select

    ICustomDraw_CustomDraw = True
End Function

Public Function CRC32(ByVal Text As String, Optional ByVal nResult As Long = &HFFFFFFFF) As Long
    Dim i As Long
    Dim Index As Long

    'If Not m_blnCRC32Init Then CRC32_Init

    For i = 1 To Len(Text)
        Index = (nResult And &HFF) Xor AscW(Mid$(Text, i, 1))
        nResult = (((nResult And &HFFFFFF00) \ &H100) And 16777215)    ' Xor m_lngCRC32LookUp(index)
    Next i

    CRC32 = Not nResult
End Function


Private Sub imgSplitMain_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    mLMouseDown = True
    picSplitMain.Visible = True
    ListView2.ColumnsAutoSize = False
End Sub

Private Sub imgSplitMain_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If mLMouseDown Then
        If imgSplitMain.Left + X - TreeFiles.Left > 1000 And PicClientArea.ScaleWidth - X - imgSplitMain.Left > 1000 Then
            picSplitMain.Move imgSplitMain.Left + X
            imgSplitMain.Move imgSplitMain.Left + X
            TreeFiles.Move TreeFiles.Left, TreeFiles.Top, imgSplitMain.Left - TreeFiles.Left, TreeFiles.Height
            ListView2.Move ListView2.Left + X, ListView2.Top, PicClientArea.ScaleWidth - imgSplitMain.Left + X - 6 * Screen.TwipsPerPixelX, ListView2.Height
        End If
    End If
End Sub

Private Sub imgSplitMain_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    mLMouseDown = False
    picSplitMain.Visible = False
End Sub


Private Sub ISubclass_WndProc(ByVal bBefore As Boolean, bHandled As Boolean, lReturn As Long, ByVal lng_hWnd As Long, ByVal uMsg As eMsg, ByVal wParam As Long, ByVal lParam As Long, lParamUser As Long)
    Dim MinMax As MINMAXINFO

    Dim pt As POINTAPI


    'Check for request for min/max window sizes.
    If uMsg = WM_GETMINMAXINFO Then
        bHandled = True

        'Retrieve default MinMax settings
        CopyMemoryToMinMaxInfo MinMax, lParam, Len(MinMax)

        'Specify new minimum size for window.
        MinMax.ptMinTrackSize.X = 568
        MinMax.ptMinTrackSize.Y = 305

        'Specify new maximum size for window.
        'MinMax.ptMaxTrackSize.x = 0
        'MinMax.ptMaxTrackSize.y = 0

        'Copy local structure back.
        CopyMemoryFromMinMaxInfo lParam, MinMax, Len(MinMax)
    ElseIf uMsg = WM_MOVING Or uMsg = WM_SIZE Then
        With frmLibrary
            If .WindowState = vbNormal And .Visible Then
                If XPos <> .Left Then _
                   XPos = .Left
                If YPos <> .Top Then _
                   YPos = .Top
            End If
        End With
    ElseIf uMsg = WM_SIZING Then
        With frmLibrary
            If .WindowState = vbNormal And .Visible Then
                If XSize <> .Width Then _
                   XSize = .Width
                If YSize <> .Height Then _
                   YSize = .Height
            End If
        End With
    ElseIf uMsg = WM_EXITSIZEMOVE Then    '
        frmLibrary.ListView2.ColumnsAutoSize = False
    End If

End Sub

Private Sub ListView2_Click(ByVal ItemIndex As Long)
    Dim sTreeviewtext As String
    sTreeviewtext = TreeFiles.Nodes(iActiveKeyIndex).Text
    If bPlaylistActive = True And ListView2.SelectedItem >= 0 Then
        StatusBar1.Panels(1).Text = sTreeviewtext & ": " _
                                    & ListView2.ItemText(ListView2.SelectedItem, 1) & "[" & ListView2.SelectedItem + 1 & "]"
    ElseIf ListView2.SelectedItem >= 0 Then
        StatusBar1.Panels(1).Text = sTreeviewtext & ": " _
                                    & ListView2.ItemText(ListView2.SelectedItem, 0) & "[" & ListView2.SelectedItem + 1 & "]"

    Else
        StatusBar1.Panels(1).Text = sTreeviewtext & ": "
    End If
End Sub

Private Sub ListView2_DblClick(ByVal ItemIndex As Long)
    Dim iTrack As Integer
    Dim bEureka As Boolean
    Dim sFile As String
    On Error Resume Next

    If bPlaylistActive = True Then
        sFile = ListView2.ItemText(ItemIndex, 3)
    Else
        sFile = ListView2.ItemText(ItemIndex, 8)
    End If
    If Dir(sFile) = "" Then Exit Sub

    For iTrack = 0 To frmPLST.clist.ItemCount - 1
        If LCase(sFile) = LCase(frmPLST.clist.exItem(iTrack)) Then
            bEureka = True
            Exit For
        End If
    Next

    If bEureka = False Then
        frmPLST.Add_track_to_Playlist (sFile)
        CurrentTrack_Index = frmPLST.clist.ItemCount - 1
        sFileMainPlaying = frmPLST.clist.exItem(CurrentTrack_Index)
        frmMain.PlayerIsPlaying = "true"
        'frmPLST.Scrollbar.Value = frmPLST.Scrollbar.Max
        frmPLST.Update_Plst_Scrollbar
        frmMain.Play
    Else
        CurrentTrack_Index = iTrack
        sFileMainPlaying = frmPLST.clist.exItem(CurrentTrack_Index)
        frmMain.PlayerIsPlaying = "true"
        frmPLST.ReinitializeList
        frmMain.Play
    End If

End Sub

Private Sub ListView2_ItemDrag()
    If Not bPlaylistActive Then ListView2.InitOLEDrag
End Sub

Private Sub ListView2_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 46 And ListView2.ItemCount > 0 Then mnuremove_Click    'DELETE
End Sub

Private Sub ListView2_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        mnuPlay_Click    'ENTER
    ElseIf KeyAscii = 12 Then
        If ListView2.ItemCount > 0 Then mnuremove_Click    'DELETE
    End If

End Sub

Private Sub ListView2_MouseDown(ByVal ItemIndex As Long, ByVal MouseButton As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Dim k As Long
    k = ListView2.RowFromPoint(X, Y, True)
    If k = -1 Then Exit Sub
    iDragitemIndex = ItemIndex    'this variable stores the index of item which is being dragged using mouse

End Sub

Private Sub ListView2_MouseUp(ByVal ItemIndex As Long, ByVal MouseButton As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
'On Error GoTo hell:
    Dim k As Long
    k = ListView2.RowFromPoint(X, Y, True)
    If k = -1 Then Exit Sub
    If MouseButton = vbRightButton Then

        If bPlaylistActive Then
            enumPopup = enormallistview
        Else
            enumPopup = ePlaylistview
        End If

        If iActivePlaylist > 0 Then mnuPlstname(iActivePlaylist).Enabled = False
        mnuremove.Caption = "Remove Selected Track(s)"
        PopupMenu mnuPopupListview
        mnuPlstname(iActivePlaylist).Enabled = True

    End If
hell:
End Sub

Private Sub ListView2_OLECompleteDrag(Effect As Long)
    bOLEdragging = False
    iDragitemIndex = -1

End Sub

Private Sub ListView2_OLEDragDrop(Data As DataObject, Effect As Long, MouseButton As Integer, Shift As Integer, X As Single, Y As Single)
    Dim i As Long
    Dim lngItem As Long
    Dim lngOver As Long
    Dim FileName As String
    Dim extension As String
    Dim bfilefound As Boolean
    Dim tt, icnt, ipos As Integer
    Dim sFile As String
    Dim Index As Integer

    If Data.GetFormat(vbCFFiles) = False Then Exit Sub
    If bOLEdragging = True Then Exit Sub
    If Data.Files.Count = 0 Then Exit Sub

    lngOver = ListView2.RowFromPoint(X, Y)
    If ListView2.RowFromPoint(X, Y + 4) > lngOver Then lngOver = lngOver + 1
    If lngOver > ListView2.ItemCount - 1 Then lngOver = -1
    If lngOver < 0 Then lngOver = -1

    ListView2.Redraw = False

    For icnt = 1 To Data.Files.Count
        If FileExists(Data.Files(icnt)) Then
            'This function will add the index of this file added to the listview in order to
            'create a sequence of playback in normal sequential mode or shuffle mode
            ipos = InStrRev(Data.Files(icnt), ".")
            extension = Mid(Data.Files(icnt), ipos + 1, Len(Data.Files(icnt)) - ipos)
            If UCase(extension) = "MP3" Or UCase(extension) = "MP2" Or UCase(extension) = "MP1" Or UCase(extension) = "WAV" Then
                ipos = InStrRev(Data.Files(icnt), "\")
                FileName = (Mid(Data.Files(icnt), ipos + 1, Len(Data.Files(icnt)) - ipos - 4))
                If iActivePlaylist > 0 And bPlaylistActive = True Then
                    Dim sTitle As String
                    Dim cfile As New cMP3
                    cfile.Read_MPEGInfo = True
                    cfile.Read_File_Tags Data.Files(icnt)
                    sTitle = cfile.Title
                    If sTitle = "" Then sTitle = GetFileTitle(Data.Files(icnt))
                    Call AddTracktoPlaylist(Data.Files(icnt), sTitle, tPlaylists(iActivePlaylist).Name, cfile.MPEG_DurationTime, 0)
                    Call DisplayTrack_in_List(Data.Files(icnt), lngOver)
                    Index = Index + 1
                Else
                    Call AddTracktoLibrary(Data.Files(icnt), True)
                    Call DisplayTrack_in_List(Data.Files(icnt), lngOver)
                    Index = Index + 1
                End If
            ElseIf FolderExists(Data.Files(icnt)) Then
                FileDropList.Path = Data.Files(icnt)

                For i = 1 To FileDropList.ListCount
                    sFile = AddBackSlash(FileDropList.Path) & FileDropList.list(i - 1)
                    If iActivePlaylist > 0 And bPlaylistActive = True Then
                        cfile.Read_MPEGInfo = True
                        cfile.Read_File_Tags sFile
                        sTitle = cfile.Title
                        If sTitle = "" Then sTitle = GetFileTitle(sFile)
                        Call AddTracktoPlaylist(sFile, sTitle, tPlaylists(iActivePlaylist).Name, cfile.MPEG_DurationTime, 0)
                        Call DisplayTrack_in_List(sFile, lngOver)
                        Index = Index + 1
                    Else
                        Call AddTracktoLibrary(sFile, True)
                        Call DisplayTrack_in_List(sFile, lngOver)
                        Index = Index + 1
                    End If

                Next
            End If

        End If
1:
    Next icnt


    ListView2.Redraw = True
End Sub

Private Sub ListView2_OLEDragOver(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single, state As Integer)
    Dim lngDropItem As Long
    If bOLEdragging Then Exit Sub
    ' user is dragging some data over the listview
    '
    ' show where the data would be dropped

    lngDropItem = ListView2.RowFromPoint(X, Y)

    If lngDropItem <> m_lngLastDropItem Then
        If m_lngLastDropItem > -1 Then
            ListView2.ItemSelected(m_lngLastDropItem) = False
            m_lngLastDropItem = lngDropItem
        End If
    End If

    If lngDropItem >= 0 And lngDropItem <= ListView2.ItemCount - 1 Then
        ListView2.ItemSelected(lngDropItem) = True
        m_lngLastDropItem = lngDropItem
    Else
        m_lngLastDropItem = -1
    End If
End Sub

Private Sub ListView2_OLEStartDrag(Data As DataObject, AllowedEffects As Long)
    AllowedEffects = vbDropEffectCopy
    Data.Clear
    Dim i As Integer
    Dim sFile As String

    If ListView2.ItemCount = 0 Then Exit Sub

    For i = 0 To ListView2.ItemCount - 1
        If ListView2.ItemSelected(i) Then
            'Check if call is from Noraml list view entry or playlist entry
            If bPlaylistActive = True Then
                sFile = ListView2.ItemText(i, 3)
            Else
                sFile = ListView2.ItemText(i, 8)
            End If
            Data.Files.Add sFile
        End If
    Next
    Data.SetData , vbCFFiles
    bOLEdragging = True

End Sub

Private Sub ListView2_Reorder()
'we need upating database of playlist
    If bPlaylistActive Then bPlaylistEdited = True
End Sub

Private Sub mnuAbout_Click()
    frmMain.ScrollText(1).CaptionText = "--MAHESH MP3 PLAYER--"    '& (index + 1) & " " & Str(Eq_SliderCtrl(index).Value) & " kHz"
    frmAbout.Show
End Sub

Private Sub mnuAddmedia_Click(Index As Integer)
    boolSearchShow = True
    frmSearch.bAddtracktoPlaylist = False    'Default addition to playlist to removed since it is call from library
    frmSearch.Show
End Sub

Private Sub mnuEnqueue_Click()
    On Error GoTo hell

    Dim i As Integer
    Dim X As Integer
    Dim sFile As String

    Select Case enumPopup
    Case enormallistview, ePlaylistview

        If ListView2.ItemCount = 0 Then Exit Sub

        For i = 0 To ListView2.ItemCount - 1
            If ListView2.ItemSelected(i) Then
                'Check if call is from Noraml list view entry or playlist entry
                If bPlaylistActive = True Then
                    sFile = ListView2.ItemText(i, 3)
                Else
                    sFile = ListView2.ItemText(i, 8)
                End If
                Call frmPLST.Add_track_to_Playlist(sFile, -1)
            End If
        Next
        frmPLST.Update_Plst_Scrollbar

    Case enormaltreeview
        Dim Key As String
        Key = sSelectedTreeKey
        If TreeFiles.SelectedItem.Children > 0 Then
            Dim oNode    'As node
            Set oNode = TreeFiles.Nodes(Key)
            EnqueueParentNode (oNode)
        Else
            EnqueueTreeView_Entry (sSelectedTreeKey)
        End If

    Case eplaylisttreeview
        EnqueueTreeView_Entry (sSelectedTreeKey)

    End Select

    frmPLST.Update_Plst_Scrollbar
    frmPLST.Scrollbar.Value = frmPLST.Scrollbar.Max
    Exit Sub
hell:
    MsgBox err.Description
End Sub

Private Sub mnuEraseLibrary_Click()
    Dim k
    If k = 7 Then Exit Sub
    k = MsgBox("Are you sure you want to Erase all data from library?", vbYesNo, "NeoPlayer Media Library")
    'k=6=Yes,  k=7=no
    If k = 6 Then
        On Error GoTo hell
        Dim rs As New ADODB.Recordset
        rs.Open "SELECT FILE FROM MUSIC", cnnMusic, adOpenDynamic, adLockOptimistic

        Do Until rs.EOF
            rs.Delete
            rs.MoveNext
        Loop
        rs.Update
        rs.Close
        Set rs = Nothing
    End If
hell:
    LoadLibrary
    StatusBar1.Panels(1).Text = "Neo Player: Media Library"

End Sub

Private Sub mnuExit_Click()
    Call Form_Unload(0)
End Sub

Private Sub mnuExplore_Click()
    Dim strPathExplore As String
    If ListView2.ItemCount < 1 Then Exit Sub

    'Check if call is from Normal list view entry or playlist entry
    If bPlaylistActive = True Then
        strPathExplore = ListView2.ItemText(ListView2.SelectedItem, 3)
    Else
        strPathExplore = ListView2.ItemText(ListView2.SelectedItem, 8)
    End If

    strPathExplore = Left(strPathExplore, InStrRev(strPathExplore, "\"))
    Shell "explorer.exe " & strPathExplore, vbMaximizedFocus
End Sub

Private Sub mnuExportList_Click()
    If ListView2.ItemCount = 0 Then MsgBox "No files selected to export, make sure that there are files in the ": Exit Sub
    Dim sFile As String, sfilter As String
    sfilter = "Playlist Files (*.npl)|*.npl|" & "Winamp Playlistfile (.m3u)|*.m3u|" & "MediaPlayer Pls file (.PLS)|*.pls|"
    sFile = clsDlg.GetSaveAsName(frmLibrary.hwnd, "Export current Files in Listview as..", , sfilter)
    If sFile = vbNullString Then Exit Sub
    ExportPlaylist (sFile)
End Sub

Private Sub mnuFavTracks_Click()
    On Error Resume Next
    Dim sNode
    Set TreeFiles.SelectedItem = TreeFiles.Nodes("kTopHits")
    Set sNode = TreeFiles.SelectedItem
    TreeFiles_NodeClick (sNode)
    '  TreeFiles.Nodes.Add "kList", tvwChild, "kTopHits", "Top Hits", 12
    '  TreeFiles.Nodes.Add "kList", tvwChild, "kRecP", "Recently Played", 11
    ' TreeFiles.Nodes.Add "kList", tvwChild, "kRecI", "Recently Imported", 10

End Sub

Private Sub mnuHistory_Click()
    On Error Resume Next
    Dim sNode
    Set TreeFiles.SelectedItem = TreeFiles.Nodes("kRecP")
    Set sNode = TreeFiles.SelectedItem
    TreeFiles_NodeClick (sNode)

End Sub

Private Sub mnuImportCurrentlist_Click()
    Dim strPlstName As String    ' Name of playlist
    Dim importCount As Integer    'Number of imported playlist
    Dim bCheckAgain As Boolean
    Dim i As Integer
    importCount = 0
    bCheckAgain = True

    Do While bCheckAgain
        bCheckAgain = False
        strPlstName = "Imported Playlist " & importCount + 1
        For i = 0 To UBound(tPlaylists)    ' Note:Ubound here gives  tPlaylist.count-1
            If UCase(tPlaylists(i).Name) = UCase(strPlstName) Then importCount = importCount + 1: bCheckAgain = True    ': Exit For
        Next
    Loop

    '8793659763 nakul suhane
    ReDim Preserve tPlaylists(LBound(tPlaylists) To UBound(tPlaylists) + 1)
    tPlaylists(UBound(tPlaylists)).Name = strPlstName
    tPlaylists(UBound(tPlaylists)).TrackCount = 0
    TreeFiles.Nodes.Add "kPlaL", tvwChild, "kPlsT" & strPlstName, strPlstName, 14
    Load mnuPlstname(UBound(tPlaylists))
    mnuPlstname(UBound(tPlaylists)).Caption = "Playlist: " & strPlstName

    With frmPLST.clist
        For i = 0 To .ItemCount - 1
            Call AddTracktoPlaylist(.exItem(i), .Item(i), strPlstName, .exTracklength(i), 0)
        Next
    End With

End Sub

Private Sub mnuImportFile_Click()
    On Error Resume Next
    Dim sFile As String, sfilter As String

    sfilter = "All Supported Playlist Files|*.npl;*.m3u;*.pls|" & "Playlist Files (*.npl)|*.npl|" & "Winamp Playlistfile (.m3u)|*.m3u|" & "MediaPlayer Pls file (.PLS)|*.pls|"
    sFile = clsDlg.GetOpenAsName(frmPopUp.hwnd, "Import Playlist to Library", , sfilter)

    If FileExists(sFile) = False Then Exit Sub

    Dim strPlstName As String    ' Name of playlist
    Dim importCount As Integer    'Number of imported playlist
    Dim bCheckAgain As Boolean
    Dim i As Integer

    importCount = 0
    bCheckAgain = True    'Flag to check if playlistname already exists

    Do While bCheckAgain
        bCheckAgain = False
        strPlstName = "Imported Playlist " & importCount + 1
        For i = 0 To UBound(tPlaylists)    ' Note:Ubound here gives  tPlaylist.count-1
            If UCase(tPlaylists(i).Name) = UCase(strPlstName) Then importCount = importCount + 1: bCheckAgain = True    ': Exit For
        Next
    Loop

    'Update popupmenu and treefile entry
    ReDim Preserve tPlaylists(LBound(tPlaylists) To UBound(tPlaylists) + 1)
    tPlaylists(UBound(tPlaylists)).Name = strPlstName
    tPlaylists(UBound(tPlaylists)).TrackCount = 0
    TreeFiles.Nodes.Add "kPlaL", tvwChild, "kPlsT" & strPlstName, strPlstName, 14
    Load mnuPlstname(UBound(tPlaylists))
    mnuPlstname(UBound(tPlaylists)).Caption = "Playlist: " & strPlstName
    'AddPlaylist to library

    ImportPlaylist sFile, strPlstName
End Sub

'Description: Imports playlists into Library (.m3u,.pls,.npl)
'Parameter: Openpath: Path of playlist file to be openned
Private Sub ImportPlaylist(openPath As String, strPlstName As String)
'On Error GoTo errorHandler:
    Dim tTrack As FileTrack
    Dim i As Integer
    i = 1
    If bLoading = True And openPath = "" Then openPath = App.Path + "\mplayerlist.npl"    'load playlist when loading application
    If FileExists(openPath) = False Then Exit Sub

    Dim sExtension As String
    'Get file extension
    sExtension = Trim(LCase(Right(openPath, Len(openPath) - InStrRev(openPath, "."))))
    'Debug.Print sExtension & openPath
    'Debug.Print sExtension & "mahesh"

    Select Case sExtension
    Case "npl"
        Open openPath For Random As #5 Len = 255

        Do While (1)
            Get #5, i, tTrack
            If tTrack.trackPath = "" Then Exit Do
            'frmPLST.clist.AddItem tTrack.trackName, tTrack.trackPath, tTrack.Duration
            Call AddTracktoPlaylist(tTrack.trackPath, tTrack.trackName, strPlstName, tTrack.Duration, -1)
            i = i + 1
            If cGetInputState() <> 0 Then DoEvents
        Loop
        Close #5

    Case "m3u"
        Dim sBuff As String, M3UChk As String * 7

        '// Check for M3U Header
        Open openPath For Binary As #1
        Get 1#, 1, M3UChk
        If M3UChk <> "#EXTM3U" Then Exit Sub
        Close #1
        DoEvents
        '// Adding procedure
        Open openPath For Input As #1
        Do While Not EOF(1)
            Line Input #1, sBuff
            'If IsBlank(sBuff) Then GoTo 1
            If Mid(sBuff, 1, 1) = "#" Then GoTo 1
            Call AddTracktoPlaylist(sBuff, GetFileTitle(sBuff), strPlstName, "", -1)
            If cGetInputState() <> 0 Then DoEvents

1
        Loop
        Close #1
    Case "pls"
        Dim lMax As Long
        Dim sFile As String, sTitle As String, sLength As String
        '// Check the file

        '// Get # of files in this playlist
        lMax = CLng(GetFromINI("playlist", "NumberOfEntries", openPath, 0))
        'Debug.Print lMax
        If lMax = 0 Then Exit Sub

        For i = 0 To lMax    '// Get the files INI values
            sFile = GetFromINI("playlist", "File" & i + 1, openPath, "")
            If IsBlank(sFile) Then GoTo 2
            sTitle = GetFromINI("playlist", "Title" & i + 1, openPath, GetFileTitle(sFile))
            sLength = GetFromINI("playlist", "Length" & i + 1, openPath, "00:00")
            sLength = Convert_Time_to_string(CLng(sLength))
            Call AddTracktoPlaylist(sFile, sTitle, strPlstName, sLength, -1)
            If cGetInputState() <> 0 Then DoEvents

2
        Next

    End Select

errorHandler:

End Sub


Private Sub mnuNewplaylist_Click()
    Dim i As Integer
    Dim strPlstName As String
    Dim Count As String    'check for suplicate entries
    'strPlstName = InputBox  "New Playlist", "Neo Player", , , , "Enter the name of new playlist")
    strPlstName = InputBox("Enter the name of new playlist")
    If strPlstName = "" Then Exit Sub

    For i = 0 To UBound(tPlaylists)
        If UCase(tPlaylists(i).Name) = UCase(strPlstName) Then
            Dim k
            k = MsgBox("Please Choose different name of playlist, playlist with this name already exist in the library", vbRetryCancel, "Add new playlist")
            If k = 4 Then    'retry
                mnuNewplaylist_Click
            Else    'cancelled
                Exit Sub
            End If
        End If
    Next

    ReDim Preserve tPlaylists(LBound(tPlaylists) To UBound(tPlaylists) + 1)
    tPlaylists(UBound(tPlaylists)).Name = strPlstName
    tPlaylists(UBound(tPlaylists)).TrackCount = 0
    TreeFiles.Nodes.Add "kPlaL", tvwChild, "kPlsT" & strPlstName, strPlstName, 14
    Load mnuPlstname(mnuPlstname.UBound + 1)
    mnuPlstname(mnuPlstname.UBound).Caption = "Playlist: " & strPlstName
End Sub

Private Sub mnuPlay_Click()
'On Error GoTo hell

    Dim i As Integer
    Dim X As Integer
    Dim sFile As String

    Select Case enumPopup
    Case enormallistview, ePlaylistview

        If ListView2.ItemCount = 0 Then Exit Sub
        frmPLST.ClearList

        For i = 0 To ListView2.ItemCount - 1
            If ListView2.ItemSelected(i) Then
                'Check if call is from Noraml list view entry or playlist entry
                If bPlaylistActive = True Then
                    sFile = ListView2.ItemText(i, 3)
                Else
                    sFile = ListView2.ItemText(i, 8)
                End If
                Call frmPLST.Add_track_to_Playlist(sFile, -1)
            End If
        Next
        frmPLST.Update_Plst_Scrollbar

    Case enormaltreeview
        Dim Key As String
        Key = sSelectedTreeKey
        frmPLST.ClearList
        If TreeFiles.SelectedItem.Children > 0 Then
            Dim oNode    'As node
            Set oNode = TreeFiles.Nodes(Key)
            EnqueueParentNode (oNode)
        Else
            Call EnqueueTreeView_Entry(sSelectedTreeKey, True)
        End If

    Case eplaylisttreeview
        Call EnqueueTreeView_Entry(sSelectedTreeKey, True)

    End Select

    CurrentTrack_Index = 0
    NormalSelection = 0
    frmPLST.Scrollbar.Value = 0
    sFileMainPlaying = frmPLST.clist.exItem(CurrentTrack_Index)
    frmMain.PlayerIsPlaying = "true"
    frmPLST.Update_Plst_Scrollbar    'to update Listrange and Scroll bar Value
    frmMain.Play
hell:
End Sub

Private Sub mnuPlstname_Click(Index As Integer)
    Dim strPlstName As String
    Dim i As Integer

    '*****************************************
    '*****************************************
    'THE FOLLOWING CODE IS TO TAKE NEW NAME OF PLAYLIST IF USER CLICKED
    'ON Send to new List
    If Index = 0 Then
        strPlstName = InputBox("Enter the name of new playlist")
        If strPlstName = "" Then Exit Sub
        'the following loop is to check if name already exists
        For i = 0 To UBound(tPlaylists)
            If UCase(tPlaylists(i).Name) = UCase(strPlstName) Then
                Dim k
                k = MsgBox("Please Choose different name of playlist, playlist with this name already exist in the library", vbRetryCancel, "Add new playlist")
                If k = 4 Then    'retry
                    mnuNewplaylist_Click
                Else    'cancelled
                    Exit Sub
                End If
            End If
        Next i

        ReDim Preserve tPlaylists(LBound(tPlaylists) To UBound(tPlaylists) + 1)
        tPlaylists(UBound(tPlaylists)).Name = strPlstName
        tPlaylists(UBound(tPlaylists)).TrackCount = 0
        TreeFiles.Nodes.Add "kPlaL", tvwChild, "kPlsT" & strPlstName, strPlstName, 14
        Load mnuPlstname(mnuPlstname.UBound + 1)
        mnuPlstname(mnuPlstname.UBound).Caption = "Playlist: " & strPlstName
    Else
        strPlstName = tPlaylists(Index).Name
    End If
    '*****************************************
    '*****************************************


    '--------------------------------
    '*****************************************
    'The following code is for Sending Playlist from  different nodes or listview etc
    Select Case enumPopup
    Case enormallistview, ePlaylistview

        Dim sFile As String

        If ListView2.ItemCount = 0 Then Exit Sub

        For i = 0 To ListView2.ItemCount - 1
            If ListView2.ItemSelected(i) Then
                'add all selected items to selected playlist
                With ListView2    '.ListItems.Item(i)

                    If bPlaylistActive = True Then
                        Call AddTracktoPlaylist(.ItemText(i, 3), .ItemText(i, 1), strPlstName, .ItemText(i, 2), 0)
                    Else
                        Call AddTracktoPlaylist(.ItemText(i, 8), .ItemText(i, 0), strPlstName, .ItemText(i, 5), 0)
                    End If

                End With
            End If
        Next i

    Case enormaltreeview, eplaylisttreeview

        Dim Key As String
        Key = sSelectedTreeKey
        If TreeFiles.SelectedItem.Children > 0 Then
            Dim oNode    'As node
            Set oNode = TreeFiles.Nodes(Key)
            Call SendParentNodetoList(oNode, strPlstName)
        Else
            Call SendtoList(sSelectedTreeKey, strPlstName)
        End If

    End Select
    '*****************************************
    '--------------------------------
End Sub

Public Sub mnurefresh_Click()
    LoadLibrary
    Dim sSQL As String
    sSQL = "SELECT " & "TITLE,ARTIST,ALBUM,GENRE,YEAR,LENGTH,PLAYCOUNT,PLAYEDLAST,FILE" & " FROM MUSIC " & "WHERE ONCD=FALSE"
    BindToSQL sSQL
    bPlaylistActive = False
    iActivePlaylist = 0
End Sub

Private Sub mnuremove_Click()
'On Error GoTo hell:
    Dim sFile, sSQL As String
    Dim i As Integer, Index As Integer
    Dim k
    If bPlaylistActive = False Then
        k = MsgBox("Are you sure you want to remove selected item(s) from library?", vbYesNo, "Confirmation")
        If k = 7 Then Exit Sub
    End If

    Dim Key As String
    Key = sSelectedTreeKey
    Index = 1
    Select Case enumPopup
    Case enormallistview, ePlaylistview

        For i = ListView2.ItemCount - 1 To 0 Step -1
            If ListView2.ItemSelected(i) Then

                If bPlaylistActive = True Then
                    sFile = ListView2.ItemText(i, 3)
                    sSQL = Replace(sFile, "'", "''", , , vbTextCompare)
                    CMD.CommandText = "DELETE FROM PLAYLIST WHERE TRACKPATH='" & sSQL & "'"
                    CMD.Execute
                    bPlaylistEdited = True
                Else
                    sFile = ListView2.ItemText(i, 8)
                    sSQL = Replace(sFile, "'", "''", , , vbTextCompare)
                    CMD.CommandText = "DELETE FROM MUSIC WHERE FILE='" & sSQL & "'"
                    CMD.Execute
                End If
                ListView2.RemoveItem i
            End If

        Next i


    Case enormaltreeview
        If TreeFiles.SelectedItem.Children > 0 Then
            Dim oNode    'As node
            Set oNode = TreeFiles.Nodes(Key)
            RemoveParentNode (oNode)
        Else
            RemoveTreeView_Entry (Key)
        End If
        TreeFiles.Nodes.Remove Key
        ' LoadLibrary

    Case eplaylisttreeview
        RemoveTreeView_Entry (Key)
        TreeFiles.Nodes.Remove Key
        ' LoadLibrary
    End Select


hell:
End Sub
Private Sub mnuRemovemissing_Click()
    On Error Resume Next
    Dim rs As New ADODB.Recordset
    rs.Open "SELECT FILE FROM MUSIC", cnnMusic, adOpenDynamic, adLockOptimistic
    Dim i As Integer
    i = 0
    Do Until rs.EOF
        If Dir(rs!File) = "" Then
            rs.Delete
            i = i + 1
        End If
        rs.MoveNext
    Loop
    'rs.Update
    rs.Close
    Set rs = Nothing
    LoadLibrary
    MsgBox str(i) & "Files are removed from the library"
hell:
End Sub




Private Sub mnuSendtoCurrntlist_Click()
    Call mnuEnqueue_Click
End Sub

Private Sub mnuviewtag_Click()
'On Error Resume Next
    Dim sSQL As String
    Dim rsAct As New ADODB.Recordset

    boolTagsShow = True
    frmPopUp.mnuTagEditor.Checked = boolTagsShow
    frmTags.Show
    DoEvents
    'frmTags.fileTags.Clear
    frmTags.vkFiletags.Clear
    frmTags.listRef.ListItems.Clear
    'don't let  vkfiletags draw itself repeatedly which it does on adding an item to it
    frmTags.vkFiletags.UnRefreshControl = True


    Select Case enumPopup

    Case enormallistview, ePlaylistview
        Dim i
        For i = 0 To ListView2.ItemCount - 1
            If ListView2.ItemSelected(i) Then
                If bPlaylistActive = True Then    'check call from listview if its is from playlist
                    frmTags.Load_Tags ListView2.ItemText(i, 3)    'add path to filetags
                Else
                    frmTags.Load_Tags ListView2.ItemText(i, 8)    'add path to filetags
                End If
            End If
        Next
    Case enormaltreeview
        sSQL = SQLfromNode(sSelectedTreeKey)
        rsAct.Open sSQL, cnnMusic, adOpenDynamic, adLockPessimistic

        For i = 0 To rsAct.RecordCount - 1
            frmTags.Load_Tags rsAct!File  'add path to filetags
            rsAct.MoveNext
        Next
        rsAct.Close
    Case eplaylisttreeview
        sSQL = SQLfromNode(sSelectedTreeKey)
        rsAct.Open sSQL, cnnMusic, adOpenDynamic, adLockPessimistic

        For i = 0 To rsAct.RecordCount - 1
            frmTags.Load_Tags rsAct!trackPath  'add path to filetags
            rsAct.MoveNext
        Next
        rsAct.Close
    End Select

    If frmTags.vkFiletags.ListCount = 0 Then Exit Sub
    'Show tags of first item
    frmTags.Show_tags (1)
    'Show first item as selected
    frmTags.vkFiletags.Selected(1) = True    'index starts from 1 in vkListbox instead of zero
    'Now we can draw vkfiletags listbox
    frmTags.vkFiletags.UnRefreshControl = False
    frmTags.vkFiletags.Refresh
    Exit Sub
errorHandler:
    MsgBox "mnuViewTag_click:  " & err.Description
End Sub







Private Sub Text1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then Call cmdSearch_Click
End Sub

Public Sub LoadLibrary(Optional bFormloading As Boolean = False)
    Dim sClave As String
    Dim sLastNode As String
    Dim sLAlbum As String
    Dim sPlaylist As String
    Dim sLArtist As String
    Dim sNode As String
    Dim sAlbum As String
    Dim sArtist As String
    Dim sGenre As String
    Dim rsFiles As New ADODB.Recordset
    Dim stipo As String
    On Error GoTo hell
    TreeFiles.Nodes.Clear
    ListView2.Clear
    TreeFiles.Nodes.Add , , "kAll", "Local Audio", 3
    TreeFiles.Nodes.Add , , "kMediaLibrary", "Media Library", 5
    TreeFiles.Nodes.Add "kMediaLibrary", tvwChild, "kAlbum", "Album", 6
    TreeFiles.Nodes.Add "kMediaLibrary", tvwChild, "kArtist\Album", "Artist\Album", 7
    TreeFiles.Nodes.Add "kMediaLibrary", tvwChild, "kCDMedia", "CD Media", 8
    TreeFiles.Nodes.Add "kMediaLibrary", tvwChild, "kPath", "File Location", 9
    TreeFiles.Nodes.Add "kPath", tvwChild, "kFullPath", "Full Path", 9
    TreeFiles.Nodes.Add "kPath", tvwChild, "kFolder", "by Folder", 2
    TreeFiles.Nodes.Add "kMediaLibrary", tvwChild, "kGenre", "Genre", 13
    TreeFiles.Nodes.Add , , "kList", "Play List", 14
    TreeFiles.Nodes.Add "kList", tvwChild, "kTopHits", "Top Hits", 12
    TreeFiles.Nodes.Add "kList", tvwChild, "kRecP", "Recently Played", 11
    TreeFiles.Nodes.Add "kList", tvwChild, "kRecI", "Recently Imported", 10
    TreeFiles.Nodes.Add "kList", tvwChild, "kPlaL", "Play Lists", 14


    CMD.CommandText = "SELECT DISTINCT GENRE,ARTIST FROM MUSIC WHERE ONCD =FALSE ORDER BY GENRE"

    Set rsFiles = CMD.Execute

    Dim rsArtist As New ADODB.Recordset
    Dim rsAlbum As New ADODB.Recordset

    ''  // GENRE
    Do Until rsFiles.EOF
        sGenre = CStr(Trim(rsFiles!Genre))
        If sGenre = "" Then sGenre = "Desconocido"

        If "GAAGE" & LCase(sGenre) <> sLastNode Then
            sLastNode = "GAAGE" & LCase(sGenre)
            TreeFiles.Nodes.Add "kGenre", tvwChild, sLastNode, sGenre, 13
        End If
        rsFiles.MoveNext
    Loop


    CMD.CommandText = "SELECT DISTINCT ALBUM FROM MUSIC WHERE ONCD=FALSE ORDER BY ALBUM"
    Set rsFiles = CMD.Execute
    sArtist = ""
    sAlbum = ""

    ''  // ALBUM
    Do Until rsFiles.EOF
        sAlbum = CStr(Trim(rsFiles!Album))
        If sAlbum = "" Then sAlbum = "Desconocido"

        If "A  AL" & LCase(sAlbum) <> sLastNode Then
            sLastNode = "A  AL" & LCase(sAlbum)
            TreeFiles.Nodes.Add "kAlbum", tvwChild, sLastNode, sAlbum, 6
        End If
        rsFiles.MoveNext
    Loop

    '  // ARTIST - ALBUMS




    CMD.CommandText = "SELECT DISTINCT ARTIST FROM MUSIC WHERE ONCD=FALSE ORDER BY ARTIST"
    Set rsFiles = CMD.Execute
    sLastNode = ""
    sLAlbum = ""
    Do Until rsFiles.EOF
        sArtist = CStr(Trim(rsFiles!Artist))
        If sArtist = "" Then sArtist = "Desconocido"
        If "AA AR" & LCase(sArtist) <> sLastNode Then
            sLastNode = "AA AR" & LCase(sArtist)
            TreeFiles.Nodes.Add "kArtist\Album", tvwChild, sLastNode, sArtist, 7
            CMD.CommandText = "SELECT DISTINCT ALBUM FROM MUSIC WHERE ONCD=FALSE AND ARTIST='" & rsFiles!Artist & "'"
            Set rsAlbum = CMD.Execute
            If rsAlbum.RecordCount > 1 Then
                Do Until rsAlbum.EOF
                    sAlbum = CStr(Trim(rsAlbum!Album))
                    If sAlbum = "" Then sAlbum = "Unknown"
                    If "AA AL|" & LCase(sArtist) & "|" & LCase(sAlbum) <> sLAlbum Then
                        sLAlbum = "AA AL|" & LCase(sArtist) & "|" & LCase(sAlbum)
                        TreeFiles.Nodes.Add sLastNode, tvwChild, sLAlbum, sAlbum, 6
                    End If
                    rsAlbum.MoveNext
                    If cGetInputState() <> 0 Then DoEvents
                Loop
            End If
            rsAlbum.Close

        End If
        rsFiles.MoveNext
    Loop

    '// FILE LOCATION
    Dim sKey As String, s As String
    Dim sPath() As String

    On Error Resume Next

    CMD.CommandText = "SELECT DISTINCT FILEPATH FROM MUSIC WHERE ONCD=FALSE"
    Set rsFiles = CMD.Execute
    sLastNode = ""
    sArtist = ""
    sAlbum = ""

    '// add albums folders
    Do Until rsFiles.EOF
        s = rsFiles!FilePath

        sPath = Split(s, "\", , vbTextCompare)
        TreeFiles.Nodes.Add "kFolder", tvwChild, "FL FO" & CStr(s & "\"), sPath(UBound(sPath)), 2

        If sLastNode <> sPath(0) Then
            TreeFiles.Nodes.Add "kFullPath", tvwChild, "FL CA" & CStr(sPath(0) & "\"), sPath(0), 1
            sLastNode = sPath(0)
        End If

        sKey = "FL CA" & sPath(0) & "\"
        Dim i As Integer
        For i = 1 To UBound(sPath)
            'If TreeFiles.Nodes(sKey).Children = 0 Then
            TreeFiles.Nodes.Add sKey, tvwChild, sKey & sPath(i) & "\", sPath(i), 2
            'End If
            If i = UBound(sPath) Then
                sKey = sKey & sPath(i)
            Else
                sKey = sKey & sPath(i) & "\"
            End If
        Next i
        rsFiles.MoveNext
        sKey = ""
    Loop


    '// CD MEDIA

    CMD.CommandText = "SELECT DISTINCT FILEPATH FROM MUSIC WHERE ONCD=TRUE"
    Set rsFiles = CMD.Execute
    sLastNode = ""
    sArtist = ""
    sAlbum = ""

    '// add albums folders
    Do Until rsFiles.EOF
        s = rsFiles!FilePath

        sPath = Split(s, "\", , vbTextCompare)

        If sLastNode <> sPath(0) Then
            TreeFiles.Nodes.Add "kCDMedia", tvwChild, "CDMCA" & CStr(sPath(0) & "\"), sPath(0), 1
            sLastNode = sPath(0)
        End If

        sKey = "CDMCA" & sPath(0) & "\"

        For i = 1 To UBound(sPath)
            'If TreeFiles.Nodes(sKey).Children = 0 Then
            TreeFiles.Nodes.Add sKey, tvwChild, sKey & sPath(i) & "\", sPath(i), 2
            'End If
            If i = UBound(sPath) Then
                sKey = sKey & sPath(i)
            Else
                sKey = sKey & sPath(i) & "\"
            End If
        Next i
        rsFiles.MoveNext

        sKey = ""
    Loop




    '-----------------------------------------------------------------------------------
    '// buskar los archivos de playlist y agragarlos
    'fPlayList.Pattern = "*.pls"


    'If Dir(tAppConfig.AppConfig & "Library\", vbDirectory) <> "" Then
    'fPlayList.Path = tAppConfig.AppConfig & "Library\"

    'For i = 0 To fPlayList.ListCount - 1
    '   TreeFiles.Nodes.Add "kPlaL", tvwChild, "kPlaE" & Left(fPlayList.list(i), Len(fPlayList.list(i)) - 4), Left(fPlayList.list(i), Len(fPlayList.list(i)) - 4), 14
    ' Next i
    'End If


    ' '// AGREGAR CD ROMS Y OTROS
    '    Dim FS As New FileSystemObject
    '    Dim dDrive As Drive
    '    Dim dDrives As Drives
    '
    '
    '    Set dDrives = FS.Drives
    '
    '    For Each dDrive In dDrives
    '       'If dDrive.IsReady = True Then
    '          Select Case dDrive.DriveType
    '
    '             Case 0 '/* Desconocido
    '             Case 1 '/* Separable
    '             Case 2 '/* Fijo
    ''                cboDrives.AddItem dDrive.DriveLetter & ": [" & dDrive.VolumeName & "]"
    '             Case 3 '/* Red
    '             Case 4 '/* CDROM
    '                 If dDrive.IsReady = True Then
    '                     sGenre = dDrive.DriveLetter & ": [" & dDrive.VolumeName & "]"
    '                 Else
    '                    sGenre = dDrive.DriveLetter
    '                 End If
    '
    '                 TreeFiles.Nodes.Add "kCDS", tvwChild, "CDSFI" & dDrive.DriveLetter, sGenre, 1
    '
    '             Case 5 '/* Disco RAM
    '          End Select
    '      ' End If
    '    Next
    '
    ' Set FS = Nothing



    '////PLAYLIST
    If bFormloading = True Then    'Retrieve Playlist information from Database
        Dim rsPlaylist As New ADODB.Recordset
        CMD.CommandText = "SELECT DISTINCT PLAYLIST FROM PLAYLIST "
        Set rsPlaylist = CMD.Execute
        If rsPlaylist.RecordCount >= 1 Then
            Do Until rsPlaylist.EOF
                sPlaylist = CStr(Trim(rsPlaylist!Playlist))
                ReDim Preserve tPlaylists(LBound(tPlaylists) To UBound(tPlaylists) + 1)
                tPlaylists(UBound(tPlaylists)).Name = sPlaylist
                tPlaylists(UBound(tPlaylists)).TrackCount = 0
                TreeFiles.Nodes.Add "kPlaL", tvwChild, "kPlsT" & sPlaylist, sPlaylist, 14
                'Load mnuPlstname(UBound(tPlaylists))
                Load mnuPlstname(mnuPlstname.UBound + 1)
                mnuPlstname(mnuPlstname.UBound).Caption = "Playlist: " & sPlaylist
                rsPlaylist.MoveNext
            Loop

            For i = 1 To UBound(tPlaylists)
                sPlaylist = tPlaylists(i).Name
                CMD.CommandText = "SELECT index FROM Playlist WHERE PLAYLIST='" & sPlaylist & "'"
                Set rsPlaylist = CMD.Execute
                tPlaylists(i).TrackCount = rsPlaylist.RecordCount
            Next
        End If


        rsPlaylist.Close
        Set rsPlaylist = Nothing
    Else    'Retrieve info from tPlaylists and no need to update menu
        For i = 1 To UBound(tPlaylists)
            sPlaylist = tPlaylists(i).Name
            tPlaylists(i).TrackCount = rsPlaylist.RecordCount
            TreeFiles.Nodes.Add "kPlaL", tvwChild, "kPlsT" & sPlaylist, sPlaylist, 14
        Next
    End If
    '///................

    'Close all recordsets
    rsFiles.Close
    rsArtist.Close
    rsAlbum.Close

    Set rsFiles = Nothing
    Set rsArtist = Nothing
    Set rsAlbum = Nothing

    'Set local Audio Node to be selected
    TreeFiles.Nodes("kMediaLibrary").Expanded = True
    TreeFiles.Nodes("kPlsT").Expanded = True
    Set TreeFiles.SelectedItem = TreeFiles.Nodes(1)
    sSelectedTreeKey = TreeFiles.Nodes(1).Key
    iActiveKeyIndex = TreeFiles.SelectedItem.Index

    Call TreeFiles_NodeClick(TreeFiles.Nodes(1))
    'If ListView2.ItemCount > 0 Then ListView2.ItemSelected(0) = True
    bPlaylistActive = False
    Exit Sub
hell:

    MsgBox err.Description
End Sub

Public Sub UpdatePlaycount(sFile As String, Optional bcheckExistinLIBRARY As Boolean)
    Dim rsAct As New ADODB.Recordset
    Dim iContar As Integer
    On Error GoTo hell
    Dim s As String

    If bcheckExistinLIBRARY = True Then Call AddTracktoLibrary(sFile, True)    'no need to show in list
    s = Replace(sFile, "'", "''", , , vbTextCompare)
    rsAct.Open "SELECT PLAYCOUNT,PLAYEDLAST FROM MUSIC WHERE FILE='" & s & "'", cnnMusic, adOpenDynamic, adLockPessimistic

    If rsAct.RecordCount = 1 Then
        'rsAct!PlayCount = rsAct!PlayCount + 1
        'rsAct!PLayedLast = Now()
        'rsAct.UpdateBatch adAffectCurrent
        iContar = rsAct!Playcount + 1
        CMD.CommandText = "UPDATE MUSIC SET PLAYCOUNT=" & iContar & ",PLAYEDLAST='" & Now() & "' WHERE FILE='" & s & "'"
        CMD.Execute
    End If
    rsAct.Close
    Set rsAct = Nothing
hell:
    'MsgBox err.Description
End Sub
Public Function AddTracktoLibrary(sFile As String, Optional bcheckExistinLIBRARY As Boolean = False) As Boolean
    Dim cfile As New cMP3
    Dim rst As New ADODB.Recordset
    Dim s As String
    Dim lSeconds As Long
    Dim sTitle As String, sArtist As String, sAlbum As String, sGenre As String, sYear As String, sComment As String
    Dim bUpdateLib As Boolean
    bUpdateLib = True

    On Error GoTo errorHandler:
    sFile = Replace(sFile, "'", "''", , , vbTextCompare)
    If bcheckExistinLIBRARY Then
        bUpdateLib = Not Exist_in_Library(sFile)
    Else
        bUpdateLib = True
    End If

    'no need to check for id3 info if not to be added eg. dblclick playlist event
    If bUpdateLib = False Then AddTracktoLibrary = False: Exit Function

    cfile.Read_MPEGInfo = True
    cfile.Read_File_Tags sFile
    sTitle = Replace(cfile.Title, "'", " ", , , vbTextCompare)
    sArtist = Replace(cfile.Artist, "'", " ", , , vbTextCompare)
    sAlbum = Replace(cfile.Album, "'", " ", , , vbTextCompare)
    sYear = Replace(cfile.Year, "'", " ", , , vbTextCompare)
    sGenre = Replace(cfile.Genre, "'", " ", , , vbTextCompare)
    sComment = Replace(cfile.Comment, "'", " ", , , vbTextCompare)

    If sTitle = "" Then sTitle = GetFileTitle(sFile)
    If sArtist = "" Then sArtist = "Unknown"
    If sAlbum = "" Then sAlbum = "Unknown"
    If sYear = "" Then sYear = Year(Now())
    If sGenre = "" Then sGenre = "Other"
    If sComment = "" Then sComment = "Uncommented"

    If bUpdateLib = True Then
        rst.Open "SELECT * FROM Music", cnnMusic, adOpenDynamic, adLockOptimistic
        rst.AddNew
        rst!File = sFile
        rst!Title = sTitle
        rst!Artist = sArtist
        rst!Album = sAlbum
        rst!Year = sYear
        rst!Genre = sGenre
        rst!Comments = sComment
        rst!Length = cfile.MPEG_DurationTime
        rst!bytes = cfile.FileSize
        rst!Seconds = cfile.DurationInSecs
        ' rst!LastUpdate = cFile.LastUpdate
        rst!Playcount = 0
        rst!Quality = cfile.Quality
        rst!Situation = cfile.Situation
        ' rst!Mood = cFile.Mood
        rst!FilePath = getDirFromPath(sFile)
        'rst!OnCD = bCDROM
        rst!drive = Left(sFile, 3)
        rst.Update
        rst.Close
        AddTracktoLibrary = True    'Successfully added return TRUE
    End If

    If cGetInputState() <> 0 Then DoEvents
    AddTracktoLibrary = True
    Exit Function
errorHandler:
    MsgBox err.Description
End Function


Public Function AddTracktoPlaylist(sPath As String, sTitle As String, sPlaylist As String, sDuration As String, Index As Long) As Boolean
    Dim rst As New ADODB.Recordset

    'this code checks index of playlist
    Dim flag As Boolean
    Dim i As Integer
    i = 0
    flag = False
    Do While (flag = False)
        If UCase(tPlaylists(i).Name) = UCase(sPlaylist) Or i >= UBound(tPlaylists) Then flag = True Else i = i + 1
    Loop
    'i is location of splaylist
    tPlaylists(i).TrackCount = tPlaylists(i).TrackCount + 1
    'If Index = -1 Then Index = tPlaylists(i).clist.Index
    'tPlaylists(i).clist.AddItem sTitle, sPath, sDuration, , Index

    rst.Open "SELECT * FROM Playlist", cnnMusic, adOpenDynamic, adLockOptimistic
    rst.AddNew
    rst!Index = tPlaylists(i).TrackCount
    rst!trackName = sTitle
    rst!trackPath = sPath
    rst!Duration = sDuration
    rst!Playlist = sPlaylist
    rst.Update
    rst.Close
    If cGetInputState() <> 0 Then DoEvents

End Function
Public Sub DisplayTrack_in_List(sFile As String, Optional Index As Long = -1)
'Just display track with id3 info in listview without changing any data in library
    Dim cfile As New cMP3
    Dim lSeconds As Long, lngItem As Long
    Dim sTitle As String, sArtist As String, sAlbum As String, sGenre As String, sYear As String, sComment As String

    cfile.Read_MPEGInfo = True

    cfile.Read_File_Tags sFile
    sTitle = Replace(cfile.Title, "'", " ", , , vbTextCompare)
    sArtist = Replace(cfile.Artist, "'", " ", , , vbTextCompare)
    sAlbum = Replace(cfile.Album, "'", " ", , , vbTextCompare)
    sYear = Replace(cfile.Year, "'", " ", , , vbTextCompare)
    sGenre = Replace(cfile.Genre, "'", " ", , , vbTextCompare)
    sComment = Replace(cfile.Comment, "'", " ", , , vbTextCompare)

    If sTitle = "" Then sTitle = GetFileTitle(sFile)
    If sArtist = "" Then sArtist = "Unknown"
    If sAlbum = "" Then sAlbum = "Unknown"
    If sYear = "" Then sYear = Year(Now())
    If sGenre = "" Then sGenre = "Other"
    If sComment = "" Then sComment = "Uncommented"

    If bPlaylistActive = False Then
        lngItem = ListView2.AddItem(Index, sTitle, 0)
        ListView2.ItemText(lngItem, 1) = sArtist
        ListView2.ItemText(lngItem, 2) = sAlbum
        ListView2.ItemText(lngItem, 3) = sGenre
        ListView2.ItemText(lngItem, 4) = sYear
        ListView2.ItemText(lngItem, 5) = cfile.MPEG_DurationTime
        ListView2.ItemText(lngItem, 6) = ""
        ListView2.ItemText(lngItem, 7) = ""
        ListView2.ItemText(lngItem, 8) = sFile
    Else
        lngItem = ListView2.AddItem(Index, , 0)
        ListView2.ItemText(lngItem, 1) = sTitle
        ListView2.ItemText(lngItem, 2) = cfile.MPEG_DurationTime
        ListView2.ItemText(lngItem, 3) = sFile
    End If
End Sub

Public Function Exist_in_Library(sFile As String) As Boolean
    Dim rst As New ADODB.Recordset
    If sFile = "" Then Exit Function
    sFile = Replace(sFile, "'", "''", , , vbTextCompare)
    rst.Open "SELECT PLAYCOUNT,PLAYEDLAST FROM MUSIC WHERE FILE='" & sFile & "'", cnnMusic, adOpenDynamic, adLockPessimistic
    'If Recodcount>0 then file is existing somewhere in library
    If rst.RecordCount <> 0 Then Exist_in_Library = True
    rst.Close
End Function

Public Sub UpdateStatusBar(lKilobytes As Long, lSeconds As Long)
'On Error Resume Next
    Dim sTreeviewtext As String
    sTreeviewtext = TreeFiles.Nodes(iActiveKeyIndex).Text

    StatusBar1.Panels(2).Text = "RECORDS:[ " & ListView2.ItemCount & " ]   -   "
    If bPlaylistActive = True And ListView2.SelectedItem >= 0 Then
        StatusBar1.Panels(1).Text = sTreeviewtext & ": " & ListView2.ItemText(ListView2.SelectedItem, 1) & "[" & ListView2.SelectedItem + 1 & "]"
    ElseIf ListView2.SelectedItem >= 0 Then
        StatusBar1.Panels(1).Text = sTreeviewtext & ": " & ListView2.ItemText(ListView2.SelectedItem, 0) & "[" & ListView2.SelectedItem + 1 & "]"
    Else
        StatusBar1.Panels(1).Text = sTreeviewtext & ": "
    End If

    Dim i As Integer, ArrTime() As String
    Dim DD As Long, HH As Long, MM As Long, ss As Long, sTempTime As String

    If lKilobytes > 0 Then sTempTime = "SIZE: [ " & Format(lKilobytes, "000,000") & " KB. ]   -   "

    StatusBar1.Panels(2).Text = StatusBar1.Panels(2).Text + sTempTime
    sTempTime = ""

    If bPlaylistActive = True Then
        For i = 0 To ListView2.ItemCount - 1
            sTempTime = ListView2.ItemText(i, 2)
            ArrTime = Split(sTempTime, ":", , vbTextCompare)
            lSeconds = lSeconds + Val(Right(sTempTime, 2)) + 60 * Val(Left(sTempTime, 2))    '+ 3600 * CLng(Left(sTempTime, Len(sTempTime) - 5))
            'Val (Left(sTemptime, InStrRev(sTemptime, ":") + 1))
            '  'Debug.Print lSeconds
        Next
    End If


    DD = lSeconds \ 86400     ' Days
    lSeconds = Abs(lSeconds - (DD * 86400))
    HH = lSeconds \ 3600      ' Hours
    MM = lSeconds \ 60 Mod 60    ' Minutes
    ss = lSeconds Mod 60      ' Seconds
    sTempTime = "TIME:[ "
    If DD > 0 Then sTempTime = sTempTime & DD & " days. "
    If HH > 0 Then sTempTime = sTempTime & HH & " Hr. "
    StatusBar1.Panels(2).Text = StatusBar1.Panels(2).Text + sTempTime & MM & " Min. " & Format$(ss, "00") & " Sec. ]"

End Sub


Private Sub TreeFiles_DblClick()
    If Not TreeFiles.SelectedItem Is Nothing Then
        Dim stipo As String
        stipo = Left(TreeFiles.SelectedItem.Key, 5)
        If stipo <> "kPlaE" And stipo <> "kRecI" And stipo <> "kRecI" And stipo <> "kPlsT" And stipo <> "kRecP" And stipo <> "kTopH" And stipo <> "FL FO" And stipo <> "CDMCA" And stipo <> "A  AL" And stipo <> "AA AR" And stipo <> "AA AL" And stipo <> "FL CA" And stipo <> "GAAGE" Then Exit Sub
        If TreeFiles.SelectedItem.Children = 0 Then mnuPlay_Click
    End If
End Sub

Private Sub TreeFiles_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'set selected item is needed to avoid doubleclick item even if doubleclicking in blank
    Set TreeFiles.SelectedItem = TreeFiles.HitTest(X, Y)
    If Not TreeFiles.HitTest(X, Y) Is Nothing Then
        sSelectedTreeKey = TreeFiles.SelectedItem
        'bPlaylistActive = False
        enumPopup = enormaltreeview
        If Left(TreeFiles.SelectedItem.Key, 5) = "kPlsT" Then enumPopup = eplaylisttreeview: bPlaylistActive = True
    End If
End Sub

Private Sub TreeFiles_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim tNode
    Dim i As Integer
    Set tNode = TreeFiles.HitTest(X, Y)
    If tNode Is Nothing Then Exit Sub
    ''Debug.Print tNode.Key
    Dim stipo As String
    stipo = Left(TreeFiles.SelectedItem.Key, 5)

    If stipo <> "kPlaE" And stipo <> "kRecI" And stipo <> "kRecI" And stipo <> "kPlsT" And stipo <> "kRecP" And stipo <> "kTopH" And stipo <> "FL FO" And stipo <> "CDMCA" And stipo <> "A  AL" And stipo <> "AA AR" And stipo <> "AA AL" And stipo <> "FL CA" And stipo <> "GAAGE" Then Exit Sub
    If Button = vbRightButton And TreeFiles.SelectedItem.Key <> "" Then
        If bPlaylistActive Then
            enumPopup = eplaylisttreeview
            mnuremove.Caption = "Delete Playlist: " & tNode.Text
            For i = 1 To UBound(tPlaylists)
                'If UCase(Right(mnuPlstname(i).Caption, Len(mnuPlstname(i).Caption) - Len("Playlist: "))) = UCase(tNode.Text) Then
                '    mnuPlstname(i).Enabled = False
                ' End If
                mnuPlstname(iActivePlaylist).Enabled = False
            Next
        Else
            enumPopup = enormaltreeview
            mnuremove.Caption = "Remove All Entries from the node '" & tNode.Text & "'"
        End If

        PopupMenu mnuPopupListview

        mnuPlstname(iActivePlaylist).Enabled = True

        'For i = 1 To UBound(tPlaylists)
        '  mnuPlstname(i).Enabled = True
        'Next

    End If
End Sub

Private Sub TreeFiles_NodeClick(ByVal node As MSComctlLib.node)
    Dim Color As Single
    Dim sAlbum As String
    Dim sArtist As String
    Dim sGenre As String
    Dim sPlaylist As String

    Dim sSQL As String
    Dim stipo As String
    Dim aEle() As String
    Dim sCampos As String
    Dim sWhere As String
    On Error Resume Next
    'Stipo FOR SELECTED  KEY IN TREEVIEW

    If (bPlaylistActive And bPlaylistEdited) = True And iActivePlaylist > 0 And TreeFiles.SelectedItem.Text <> tPlaylists(iActivePlaylist).Name Then UpdatePlaylist (tPlaylists(iActivePlaylist).Name)
    sSelectedTreeKey = TreeFiles.SelectedItem.Key
    iActiveKeyIndex = TreeFiles.SelectedItem.Index

    stipo = Left(sSelectedTreeKey, 5)
    'Debug.Print TreeFiles.SelectedItem.index & ":" & sSelectedTreeKey

    bPlaylistActive = False
    iActivePlaylist = 0
    bPlaylistEdited = False

    If stipo <> "kPlaE" And stipo <> "kRecI" And stipo <> "kRecP" And stipo <> "kTopH" And stipo <> "FL FO" And stipo <> "CDMCA" And stipo <> "kAll" And stipo <> "A  AL" And stipo <> "AA AR" And stipo <> "AA AL" And stipo <> "FL CA" And stipo <> "kPlsT" And stipo <> "GAAGE" Then Exit Sub

    sCampos = "TITLE,ARTIST,ALBUM,GENRE,YEAR,LENGTH,PLAYCOUNT,PLAYEDLAST,FILE"

    If stipo = "kAll" Then
        sWhere = "WHERE ONCD=FALSE"
        sSQL = "SELECT " & sCampos & " FROM MUSIC " & sWhere
    End If

    '//CLICK ON ALBUMS
    If stipo = "A  AL" Then
        sAlbum = Right(sSelectedTreeKey, Len(sSelectedTreeKey) - 5)
        sWhere = "WHERE ONCD=FALSE AND ALBUM='" & sAlbum & "'"
        sSQL = "SELECT " & sCampos & " FROM MUSIC " & sWhere
    End If

    '//CLICK ON ARTIST
    If stipo = "AA AR" Then
        sArtist = Right(sSelectedTreeKey, Len(sSelectedTreeKey) - 5)
        sWhere = "WHERE ONCD=FALSE AND ARTIST='" & sArtist & "'"
        sSQL = "SELECT " & sCampos & " FROM MUSIC " & sWhere
    End If
    sArtist = aEle(1)

    '//CLICK ON ARTIST - ALBUM
    If stipo = "AA AL" Then   'Here .Key= Artist,Album   so we have to split name of artist and album
        aEle = Split(sSelectedTreeKey, "|", , vbTextCompare)  'aELe IS AN ARRAY which stores artist(ele(2)) and album name(ele(1))
        sArtist = aEle(1)
        sAlbum = aEle(2)
        sWhere = "WHERE ONCD=FALSE AND ARTIST='" & sArtist & "' AND ALBUM='" & sAlbum & "'"
        sSQL = "SELECT " & sCampos & " FROM MUSIC " & sWhere
    End If

    '// CLICK ON FILE LOCATION
    If stipo = "FL CA" Or stipo = "FL FO" Then
        If TreeFiles.SelectedItem.Children = 0 Then
            sGenre = Right(sSelectedTreeKey, Len(sSelectedTreeKey) - 5)
            sGenre = Left(sGenre, Len(sGenre) - 1)
            sWhere = "WHERE ONCD=FALSE AND FILEPATH='" & sGenre & "'"
            sSQL = "SELECT " & sCampos & " FROM MUSIC " & sWhere

        End If
    End If

    '// CD MEDIA
    If stipo = "CDMCA" Then
        If TreeFiles.SelectedItem.Children = 0 Then
            sGenre = Right(sSelectedTreeKey, Len(sSelectedTreeKey) - 5)
            sGenre = Left(sGenre, Len(sGenre) - 1)
            sWhere = "WHERE ONCD=TRUE AND FILEPATH='" & sGenre & "'"
            sSQL = "SELECT " & sCampos & " FROM MUSIC " & sWhere
        Else
            If Len(sSelectedTreeKey) = 8 Then
                sGenre = Right(sSelectedTreeKey, 3)
                sWhere = "WHERE ONCD=TRUE AND DRIVE='" & sGenre & "'"
                sSQL = "SELECT " & sCampos & " FROM MUSIC " & sWhere
            End If
        End If
    End If


    '//CLICK ON GENRES
    If stipo = "GAAGE" Then
        sGenre = Right(sSelectedTreeKey, Len(sSelectedTreeKey) - 5)
        sWhere = "WHERE ONCD=FALSE AND GENRE='" & sGenre & "'"
        sSQL = "SELECT " & sCampos & " FROM MUSIC " & sWhere
    End If

    '//CLICK ON TOP HITS
    If stipo = "kTopH" Then
        sWhere = "WHERE ONCD=FALSE AND PLAYCOUNT>0 ORDER BY PLAYCOUNT DESC "
        'sSQL = "SELECT TOP 20 PLAYCOUNT," & sCampos & " FROM MUSIC " & sWhere
        sSQL = "SELECT " & sCampos & " FROM MUSIC " & sWhere
    End If


    '//CLICK ON RECENTLY PLAYED
    If stipo = "kRecP" Then
        sWhere = "WHERE ONCD=FALSE AND PLAYEDLAST IS NOT NULL " & "ORDER BY PLAYEDLAST DESC"
        'sSQL = "SELECT TOP 10 PLAYEDLAST, " & sCampos & " FROM MUSIC " & sWhere
        sSQL = "SELECT " & sCampos & " FROM MUSIC " & sWhere
    End If
    '*****iMPORTANT*****SPACE in sql SATAEMENET IS VERY IMP. "SELECT " IS CORRECT BUT "SELECT"IS WRONG
    '*****"SELECT TOP 20 LASTUPDATE, " IS CORRECT "SELECT TOP 20 LASTUPDATE," WRONG

    '//CLICK ON RECENTLY ADDED
    If stipo = "kRecI" Then
        sWhere = "WHERE ONCD=FALSE AND LASTUPDATE IS NOT NULL " & "ORDER BY LASTUPDATE DESC"
        'sSQL = "SELECT TOP 20 LASTUPDATE, " & sCampos & " FROM MUSIC " & sWhere 'CORRECT STATEMENT TO GET TOP 20 ENTRIES
        ' sSQL = "SELECT TOP 20 LASTUPDATE, " & sCampos & " FROM MUSIC " & sWhere 'CORRECT STATEMENT TO GET TOP 20 ENTRIES
        sCampos = "TITLE,ARTIST,ALBUM,GENRE,YEAR,LENGTH,PLAYCOUNT,LASTUPDATE,FILE"
        sSQL = "SELECT " & sCampos & " FROM MUSIC " & sWhere
    End If

    '//CLICK EN PLAYLISTS
    If stipo = "kPlaE" Then
        sGenre = Right(sSelectedTreeKey, Len(sSelectedTreeKey) - 5)
        'Cargar_PlayListTracks sGenre
        Exit Sub
    End If

    If stipo = "kPlaE" Then
        sGenre = Right(sSelectedTreeKey, Len(sSelectedTreeKey) - 5)
        'Cargar_PlayListTracks sGenre
        Exit Sub
    End If

    If stipo = "kPlaE" Then
        sGenre = Right(sSelectedTreeKey, Len(sSelectedTreeKey) - 5)
        'Cargar_PlayListTracks sGenre
        Exit Sub
    End If

    'stipo = sSelectedTreeKey  'Key=kPlsTpLAYLISTPOSITION/INDEX"

    If stipo = "kPlsT" Then
        bPlaylistEdited = False
        bPlaylistActive = True
        sPlaylist = TreeFiles.SelectedItem.Text
        'Right(sSelectedTreeKey, Len(sSelectedTreeKey) - 10) 'Key starts with "kPlsTPlaylistName"
        sWhere = "WHERE PLAYLIST='" & sPlaylist & "'" & " ORDER BY INDEX ASC"
        ' 'Debug.Print TreeFiles.SelectedItem.Index & ":" & sSelectedTreeKey
        iActivePlaylist = TreeFiles.SelectedItem.Index - TreeFiles.SelectedItem.Parent.Child.Index + 1
        'Debug.Print iActivePlaylist
        'sSQL = "SELECT TOP 20 LASTUPDATE, " & sCampos & " FROM MUSIC " & sWhere 'CORRECT STATEMENT TO GET TOP 20 ENTRIES
        ' sSQL = "SELECT TOP 20 LASTUPDATE, " & sCampos & " FROM MUSIC " & sWhere 'CORRECT STATEMENT TO GET TOP 20 ENTRIES
        sCampos = "Index,TrackName,Duration,TrackPath"
        sSQL = "SELECT " & sCampos & " FROM Playlist " & sWhere
    End If

    ListView2.Reorder = bPlaylistActive

    If sSQL = "" Then Exit Sub
    BindToSQL sSQL, , , bPlaylistActive

    rs.Close
    'The following lines are just to remove order by descending which creates extra field in database which i dont ant to be displayed in listview
    If stipo = "kTopH" Then sWhere = "WHERE ONCD=FALSE AND PLAYCOUNT>0 "
    If stipo = "kRecP" Then sWhere = "WHERE ONCD=FALSE AND PLAYEDLAST IS NOT NULL "
    If stipo = "kRecI" Then sWhere = "WHERE ONCD=FALSE AND LASTUPDATE IS NOT NULL "
    'On Error Resume Next

    rs.Open "SELECT SUM(BYTES) AS TOTAL, SUM(SECONDS) AS TIEMPO FROM MUSIC " & sWhere, cnnMusic, adOpenForwardOnly, adLockReadOnly

    '// UPDATE STATUS BAR
    Dim lKilobytes As Long, lSeconds As Long
    lKilobytes = CLng(rs!total / 1024)
    lSeconds = CLng(rs!TIEMPO)
    rs.Close
    Call UpdateStatusBar(lKilobytes, lSeconds)


End Sub

Public Sub RemoveTreeView_Entry(Key As String)

    Dim sAlbum As String
    Dim sArtist As String
    Dim sGenre As String
    Dim sPlaylist As String

    Dim sSQL As String
    Dim stipo As String
    Dim aEle() As String

    On Error GoTo errorHandler

    stipo = Left(Key, 5)
    Dim sCampos As String

    sCampos = "TITLE,ARTIST,ALBUM,GENRE,YEAR,LENGTH,PLAYCOUNT,PLAYEDLAST,FILE"

    Select Case stipo

    Case "A  AL"    '//CLICK ON ALBUM
        sAlbum = Right(Key, Len(Key) - 5)
        sSQL = "DELETE FROM MUSIC WHERE ALBUM='" & sAlbum & "'"

    Case "AA AR"    '//CLICK ON ARTIST
        sArtist = Right(Key, Len(Key) - 5)
        sSQL = "DELETE FROM MUSIC WHERE ARTIST='" & sArtist & "'"

    Case "AA AL"  '//CLICK ON ARTIST - ALBUM
        'Here .Key= Artist,Album   so we have to split name of artist and album
        aEle = Split(Key, "|", , vbTextCompare)  'aELe IS AN ARRAY which stores artist(ele(2)) and album name(ele(1))
        sArtist = aEle(1)
        sAlbum = aEle(2)
        sSQL = "DELETE FROM MUSIC WHERE ARTIST='" & sArtist & "' AND ALBUM='" & sAlbum & "'"

    Case "FL CA", "FL FO"  '// CLICK ON FILE LOCATION
        If TreeFiles.Nodes(Key).Children = 0 Then
            sGenre = Right(Key, Len(Key) - 5)
            sGenre = Left(sGenre, Len(sGenre) - 1)
            sSQL = "DELETE FROM MUSIC WHERE FILEPATH='" & sGenre & "'"
        End If

    Case "CDMCA"    '// CLICK ON CD
        If TreeFiles.Nodes(Key).Children = 0 Then
            sGenre = Right(Key, Len(Key) - 5)
            sGenre = Left(sGenre, Len(sGenre) - 1)
            sSQL = "DELETE FROM MUSIC WHERE FILEPATH='" & sGenre & "'"
        Else
            If Len(Key) = 8 Then
                sGenre = Right(Key, 3)
                sSQL = "DELETE FROM MUSIC WHERE ONCD=TRUE AND DRIVE='" & sGenre & "'"
            End If
        End If

    Case "GAAGE"  '// CLICK ON GENRE
        sGenre = Right(Key, Len(Key) - 5)
        sSQL = "DELETE FROM MUSIC WHERE ONCD=FALSE AND GENRE='" & sGenre & "'"

    Case "kTopH"  '//CLICK ON TOP HITS
        sSQL = "DELETE FROM MUSIC WHERE ONCD=FALSE AND PLAYCOUNT>0 "

    Case "kRecP"   '//CLICK ON RECENTLY PLAYED
        sSQL = "DELETE FROM MUSIC WHERE ONCD=FALSE AND PLAYEDLAST IS NOT NULL "

    Case "kRecI"  '//CLICK ON RECENTLY ADDED
        sSQL = "DELETE FROM MUSIC WHERE ONCD=FALSE AND LASTUPDATE IS NOT NULL "


    Case "kPlsT"   '//CLICK EN PLAYLIST
        sPlaylist = Right(Key, Len(Key) - 5)
        sSQL = "DELETE FROM PLAYLIST WHERE PLAYLIST='" & sPlaylist & "'"
        Dim i As Integer, j As Integer
        For i = 1 To UBound(tPlaylists)
            If UCase(tPlaylists(i).Name) = UCase(sPlaylist) Then j = i: Exit For
        Next

        For i = j To UBound(tPlaylists) - 1
            tPlaylists(i).Name = tPlaylists(i + 1).Name
            tPlaylists(i).TrackCount = tPlaylists(i + 1).TrackCount
            mnuPlstname(i).Caption = "Playlist: " & tPlaylists(i + 1).Name
        Next

        Unload mnuPlstname(mnuPlstname.UBound)    '(UBound(tPlaylists))
        iActivePlaylist = 0
        ReDim Preserve tPlaylists(LBound(tPlaylists) To UBound(tPlaylists) - 1)

    End Select
    CMD.CommandText = sSQL: CMD.Execute
    Exit Sub
errorHandler:

    MsgBox sSQL & err.Description
End Sub

Public Sub EnqueueTreeView_Entry(Key As String, Optional bNewlist As Boolean = False)
    On Error Resume Next    'GoTo errorHandler
    Dim sSQL As String
    sSQL = SQLfromNode(Key)
    Dim rsAct As New ADODB.Recordset
    rsAct.Open sSQL, cnnMusic, adOpenDynamic, adLockPessimistic

    Dim sCleanStr As String, sNewString As String, sFormat As String
    Dim sSplitField() As String
    Dim iSpaces As Integer

    Dim i As Integer
    If bNewlist = True And rsAct.RecordCount > 0 Then frmPLST.ClearList:
    For i = 0 To rsAct.RecordCount - 1
        If enumPopup = enormaltreeview Then
            ' load tags
            sFormat = sFormatPlayList
            '// Song Name
            sFormat = Replace(sFormatPlayList, "%S", Trim(rsAct!Title))
            '// Artist
            sFormat = Replace(sFormat, "%A", Trim(rsAct!Artist))
            '// Album
            sFormat = Replace(sFormat, "%B", Trim(rsAct!Album))
            '// Year
            sFormat = Replace(sFormat, "%Y", Trim(rsAct!Year))
            '// Genre
            sFormat = Replace(sFormat, "%G", Trim(rsAct!Genre))
            '// Time
            ' sFormat = Replace(sFormat, "%T", Trim(cFile.MPEG_DurationTime))
            '// File Name
            sFormat = Replace(sFormat, "%N", GetFileTitle(rsAct!File))
            '// PATH
            sFormat = Replace(sFormat, "%P", rsAct!File)
            '// File extencion
            ' sFormat = Replace(sFormat, "%F", sFileEx)

            If sFormat = sFormatPlayList Then sFormat = rsAct!Title    'If no fomat is there / tag info fails then display trackname
            '------------------------------------------------------------------------------
            sCleanStr = Trim$(sFormat)

            'Upper case and / or lower case the string correctly.
            sSplitField = Split(sCleanStr, " ", , vbTextCompare)
            sCleanStr = ""
            For iSpaces = 0 To UBound(sSplitField)
                If (Not iSpaces = 0 Or Not IsNumeric(sSplitField(iSpaces))) And sSplitField(iSpaces) <> "" Then
                    sNewString = UCase$(Left$(sSplitField(iSpaces), 1))
                    sNewString = sNewString & (Right$(sSplitField(iSpaces), Len(sSplitField(iSpaces)) - 1))
                    sCleanStr = sCleanStr & sNewString & " "
                End If
            Next iSpaces
            sFormat = Trim$(sCleanStr)
            '------------------------------------------------------------------------------

            frmPLST.clist.AddItem sFormat, rsAct!File, rsAct!Length
        Else
            ' frmPLST.Add_track_to_Playlist (rsAct!trackPath)
            frmPLST.clist.AddItem rsAct!trackName, rsAct!trackPath, rsAct!Duration
        End If

        rsAct.MoveNext
    Next

    rsAct.Close

    Exit Sub

errorHandler:
    MsgBox sSQL & err.Description
End Sub



Public Sub SendtoList(Key As String, sPlaylist As String)
    On Error GoTo errorHandler
    Dim sSQL As String
    sSQL = SQLfromNode(Key)
    Dim rsAct As New ADODB.Recordset
    rsAct.Open sSQL, cnnMusic, adOpenDynamic, adLockPessimistic
    Dim i As Integer
    For i = 0 To rsAct.RecordCount - 1
        If enumPopup = enormaltreeview Then
            Call AddTracktoPlaylist(rsAct!File, rsAct!Title, sPlaylist, rsAct!Length, 0)
        Else
            Call AddTracktoPlaylist(rsAct!trackPath, rsAct!trackName, sPlaylist, rsAct!Duration, 0)
        End If
        rsAct.MoveNext
    Next

    rsAct.Close

    Exit Sub

errorHandler:
    MsgBox sSQL & "sendtoList" & err.Description
End Sub


Public Function SQLfromNode(Key As String) As String
    Dim sAlbum As String
    Dim sArtist As String
    Dim sGenre As String
    Dim sPlaylist As String

    Dim sSQL As String
    Dim stipo As String
    Dim aEle() As String



    stipo = Left(Key, 5)
    Dim sCampos As String
    Dim sWhere As String

    sCampos = "TITLE,ARTIST,ALBUM,GENRE,YEAR,LENGTH,PLAYCOUNT,PLAYEDLAST,FILE"

    Select Case stipo

    Case "A  AL"    '//CLICK ON ALBUM
        sAlbum = Right(Key, Len(Key) - 5)
        sWhere = "WHERE ONCD=FALSE AND ALBUM='" & sAlbum & "'"
        sSQL = "SELECT " & sCampos & " FROM MUSIC " & sWhere
    Case "AA AR"    '//CLICK ON ARTIST
        sArtist = Right(Key, Len(Key) - 5)
        sWhere = "WHERE ONCD=FALSE AND ARTIST='" & sArtist & "'"
        sSQL = "SELECT " & sCampos & " FROM MUSIC " & sWhere
    Case "AA AL"  '//CLICK ON ARTIST - ALBUM
        'Here .Key= Artist,Album   so we have to split name of artist and album
        aEle = Split(Key, "|", , vbTextCompare)  'aELe IS AN ARRAY which stores artist(ele(2)) and album name(ele(1))
        sArtist = aEle(1)
        sAlbum = aEle(2)
        sWhere = "WHERE ONCD=FALSE AND ARTIST='" & sArtist & "' AND ALBUM='" & sAlbum & "'"
        sSQL = "SELECT " & sCampos & " FROM MUSIC " & sWhere

    Case "FL CA", "FL FO"  '// CLICK ON FILE LOCATION
        If TreeFiles.Nodes(Key).Children = 0 Then
            sGenre = Right(Key, Len(Key) - 5)
            sGenre = Left(sGenre, Len(sGenre) - 1)
            sWhere = "WHERE ONCD=FALSE AND FILEPATH='" & sGenre & "'"
            sSQL = "SELECT " & sCampos & " FROM MUSIC " & sWhere
        End If

    Case "CDMCA"    '// CLICK ON CD
        If TreeFiles.Nodes(Key).Children = 0 Then
            sGenre = Right(Key, Len(Key) - 5)
            sGenre = Left(sGenre, Len(sGenre) - 1)
            sWhere = "WHERE ONCD=TRUE AND FILEPATH='" & sGenre & "'"
            sSQL = "SELECT " & sCampos & " FROM MUSIC " & sWhere
        Else
            If Len(Key) = 8 Then
                sGenre = Right(Key, 3)
                sWhere = "WHERE ONCD=TRUE AND DRIVE='" & sGenre & "'"
                sSQL = "SELECT " & sCampos & " FROM MUSIC " & sWhere
            End If
        End If

    Case "GAAGE"  '// CLICK ON GENRE
        sGenre = Right(Key, Len(Key) - 5)
        sWhere = "WHERE ONCD=FALSE AND GENRE='" & sGenre & "'"
        sSQL = "SELECT " & sCampos & " FROM MUSIC " & sWhere

    Case "kTopH"  '//CLICK ON TOP HITS
        sWhere = "WHERE ONCD=FALSE AND PLAYCOUNT>0 ORDER BY PLAYCOUNT DESC "
        sSQL = "SELECT " & sCampos & " FROM MUSIC " & sWhere

    Case "kRecP"   '//CLICK ON RECENTLY PLAYED
        sWhere = "WHERE ONCD=FALSE AND PLAYEDLAST IS NOT NULL " & "ORDER BY PLAYEDLAST DESC"
        sSQL = "SELECT " & sCampos & " FROM MUSIC " & sWhere

    Case "kRecI"  '//CLICK ON RECENTLY ADDED
        sWhere = "WHERE ONCD=FALSE AND LASTUPDATE IS NOT NULL " & "ORDER BY LASTUPDATE DESC"
        sCampos = "TITLE,ARTIST,ALBUM,GENRE,YEAR,LENGTH,PLAYCOUNT,LASTUPDATE,FILE"
        sSQL = "SELECT " & sCampos & " FROM MUSIC " & sWhere

    Case "kPlsT"   '//CLICK EN PLAYLIST
        sPlaylist = Right(Key, Len(Key) - 5)
        sWhere = "WHERE PLAYLIST='" & sPlaylist & "'"
        sCampos = "TRACKNAME,DURATION,TRACKPATH"
        sSQL = "SELECT " & sCampos & " FROM Playlist " & sWhere
    End Select

    SQLfromNode = sSQL
End Function


Public Sub EnqueueParentNode(ByVal oParentnode As MSComctlLib.node)
'This is recursive function used to add node which has more subnodes and so on
    Dim oNode    'Dim oNode as node ives error type mismatch perhaps because of mscomctllib versons SP5 and SP@
' Get the first child node
    Set oNode = oParentnode.Child
    ' Loop through the child nodes of this node
    ' until there are none left...
    Do While Not oNode Is Nothing
        ' Check/Uncheck the node
        If oNode.Children = 0 Then EnqueueTreeView_Entry (oNode.Key)
        ' Call this function again for the child node, so that it's child nodes can get checked/uncheckedbe enqueued.
        EnqueueParentNode (oNode)
        DoEvents
        ' Get the next child node of this node
        Set oNode = oNode.Next
    Loop
End Sub


Public Sub RemoveParentNode(ByVal oParentnode As MSComctlLib.node)
'This is recursive function used to add node which has more subnodes and so on
    Dim oNode    'Dim oNode as node ives error type mismatch perhaps because of mscomctllib versons SP5 and SP@
' Get the first child node
    Set oNode = oParentnode.Child
    ' Loop through the child nodes of this node
    ' until there are none left...
    Do While Not oNode Is Nothing
        ' Check/Uncheck the node
        If oNode.Children = 0 Then RemoveTreeView_Entry (oNode.Key)
        ' Call this function again for the child node, so that it's child nodes can get checked/uncheckedbe enqueued.
        RemoveParentNode (oNode)
        DoEvents
        ' Get the next child node of this node
        Set oNode = oNode.Next
    Loop
End Sub


Public Sub SendParentNodetoList(ByVal oParentnode As MSComctlLib.node, sPlaylist As String)
'This is recursive function used to add node which has more subnodes and so on
    Dim oNode    'Dim oNode as node ives error type mismatch perhaps because of mscomctllib versons SP5 and SP@
' Get the first child node
    Set oNode = oParentnode.Child
    ' Loop through the child nodes of this node
    ' until there are none left...
    Do While Not oNode Is Nothing
        ' Check/Uncheck the node
        If oNode.Children = 0 Then SendtoList oNode.Key, sPlaylist
        ' Call this function again for the child node, so that it's child nodes can get checked/uncheckedbe enqueued.
        SendParentNodetoList oNode, sPlaylist
        DoEvents
        ' Get the next child node of this node
        Set oNode = oNode.Next
    Loop
End Sub


Public Sub UpdatePlaylist(sPlaylist As String)
    Dim rst As New ADODB.Recordset
    Dim i As Long
    'tPlaylists(i).clist.AddItem sTitle, sPath, sDuration, , Index
    Dim sSQL As String
    sSQL = "DELETE FROM PLAYLIST WHERE PLAYLIST='" & sPlaylist & "'"
    CMD.CommandText = sSQL: CMD.Execute

    rst.Open "SELECT * FROM Playlist", cnnMusic, adOpenDynamic, adLockOptimistic


    For i = 0 To ListView2.ItemCount - 1
        rst.AddNew
        rst!Index = i
        rst!trackName = ListView2.ItemText(i, 1)
        rst!trackPath = ListView2.ItemText(i, 3)
        rst!Duration = ListView2.ItemText(i, 2)
        rst!Playlist = sPlaylist
        rst.Update
        If cGetInputState() <> 0 Then DoEvents
    Next

    rst.Close

End Sub

Public Sub ExportPlaylist(sOutputPath As String)
'SAVE The CUrrent displayed files in listview as M3U FILE
    Dim k As Integer
    Dim i As Long
    Dim sExtension As String
    'Get file extension
    sExtension = Trim(LCase(Right(sOutputPath, Len(sOutputPath) - InStrRev(sOutputPath, "."))))

    k = IIf(bPlaylistActive, 1, 0)
    sOutputPath = StripNulls(sOutputPath)
    Select Case sExtension

    Case "m3u"
        Close #1
        Open sOutputPath For Output As #1
        Print #1, "#EXTM3U"    '// m3u header
        For i = 0 To ListView2.ItemCount - 1
            If bPlaylistActive Then
                Print #1, ListView2.ItemText(i, 3)  'print the file's path from playlist
            Else
                Print #1, ListView2.ItemText(i, 8)  'print the file's path from normal list view
            End If
        Next
        Close #1

    Case "pls"
        Close #2
        Open sOutputPath For Output As #2
        Print #2, "[playlist]"

        For i = 0 To ListView2.ItemCount - 1
            Print #2, "File" & i + 1 & "=" & ListView2.ItemText(i, 8 * (1 - k) + 3 * k)
            Print #2, "Title" & i + 1 & "=" & IIf(bPlaylistActive, ListView2.ItemText(i, 1), ListView2.ItemText(i, 1))
            Print #2, "Length" & i + 1 & "=" & Convert_TextTime_to_Seconds(ListView2.ItemText(i, 5 * (1 - k) + 2 * k))
        Next
        Print #2, "NumberOfEntries=" & i
        Print #2, "Version=2"
        Close #2

    Case "npl"
        Dim tTrack As FileTrack
        If FileExists(sOutputPath) Then Kill (sOutputPath)
        On Error Resume Next
        Open sOutputPath For Random As #3 Len = 255
        For i = 0 To ListView2.ItemCount - 1
            tTrack.trackName = IIf(bPlaylistActive, ListView2.ItemText(i, 1), ListView2.ItemText(i, 1))
            tTrack.trackPath = ListView2.ItemText(i, 8 * (1 - k) + 3 * k)
            tTrack.Duration = ListView2.ItemText(i, 5 * (1 - k) + 2 * k)
            Put #3, i + 1, tTrack
        Next
        Close #3
    End Select

End Sub

