Attribute VB_Name = "modSubclassMneu"
Option Explicit
Public hmenu1&, Result&
Attribute Result.VB_VarUserMemId = 1073741824
Private Declare Function SetMenuItemBitmaps Lib "user32" _
                                            (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long, _
                                             ByVal hBitmapUnchecked As Long, ByVal hBitmapChecked As Long) As Long
Private Declare Function GetMenu Lib "user32" (ByVal hwnd As Long) As Long

Private Declare Function LoadImage Lib "user32" Alias "LoadImageA" _
                                   (ByVal hInst As Long, ByVal lpsz As String, ByVal un1 As Long, _
                                    ByVal n1 As Long, ByVal n2 As Long, ByVal un2 As Long) As Long

'Constant for SetMenuItemBitmaps
Private Const MF_BYPOSITION As Long = &H400&

'Constants for LoadImage
Private Const IMAGE_BITMAP As Long = &O0
Private Const LR_LOADFROMFILE As Long = 16
Private Const LR_CREATEDIBSECTION As Long = 8192


Public Declare Function GetMenuState Lib "user32" (ByVal hMenu As Long, ByVal wID As Long, ByVal wFlags As Long) As Long
Public Declare Function DrawMenuBar Lib "user32" (ByVal hwnd As Long) As Long

Public Declare Function GetSystemMenu Lib "user32" (ByVal hwnd As Long, ByVal bRevert As Long) As Long
Private Declare Function CheckMenuItem Lib "user32" (ByVal hMenu As Long, ByVal wIDCheckItem As Long, ByVal wCheck As Long) As Long
Private Declare Function EnableMenuItem Lib "user32" (ByVal hMenu As Long, ByVal wIDEnableItem As Long, ByVal wEnable As Long) As Long
Private Declare Function GetSubMenu Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Private Declare Function GetMenuItemID Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Public Declare Function GetMenuItemCount Lib "user32" (ByVal hMenu As Long) As Long

Private Declare Function AppendMenu Lib "user32" Alias "AppendMenuA" (ByVal hMenu As Long, ByVal wFlags As Long, ByVal wIDNewItem As Long, ByVal lpNewItem As Any) As Long
Private Declare Function ModifyMenu Lib "user32" Alias "ModifyMenuA" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long, ByVal wIDNewItem As Long, ByVal lpString As Any) As Long
Public Declare Function RemoveMenu Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long) As Long
Private Declare Function DeleteMenu Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long) As Long

' Menu flags for Add/Check/EnableMenuItem().
Private Const MF_INSERT = &H0&
Private Const MF_CHANGE = &H80&
Private Const MF_APPEND = &H100&
Private Const MF_DELETE = &H200&
Private Const MF_REMOVE = &H1000&

Private Const MF_BYCOMMAND = &H0&

Private Const MF_SEPARATOR = &H800&

Private Const MF_ENABLED = &H0&
Private Const MF_GRAYED = &H1&
Private Const MF_DISABLED = &H2&

Private Const MF_UNCHECKED = &H0&
Private Const MF_CHECKED = &H8&
Private Const MF_USECHECKBITMAPS = &H200&

Private Const MF_STRING = &H0&
Private Const MF_BITMAP = &H4&
Private Const MF_OWNERDRAW = &H100&

Private Const MF_POPUP = &H10&
Private Const MF_MENUBARBREAK = &H20&
Private Const MF_MENUBREAK = &H40&

Private Const MF_UNHILITE = &H0&
Private Const MF_HILITE = &H80&

Private Const MF_SYSMENU = &H2000&
Private Const MF_HELP = &H4000&
Private Const MF_MOUSESELECT = &H8000&
Private Const MF_LINKS As Long = &H20000000


Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Declare Function CreateMenu Lib "user32" () As Long

Private Const GWL_WNDPROC = (-4&)
Private Const WM_SYSCOMMAND = &H112
Private Const WM_ACTIVATE As Long = &H6
Private Const WM_ACTIVATEAPP As Long = &H1C
Private Const HWND_TOP As Long = 0
Private Const HWND_TOPMOST As Long = -1
Private Const SWP_NOACTIVATE As Long = &H10
Private Const SWP_NOMOVE As Long = &H2
Private Const SWP_NOREDRAW As Long = &H8
Private Const SWP_NOSIZE As Long = &H1
Private Const SWP_SHOWWINDOW As Long = &H40

Private Const IDM_SEP1 = 11
Private Const IDM_SEP2 = 12

'Private Const MF_POPUP = &H10&
'Private Const MF_SEPARATOR = &H800&
'Private Const MF_STRING = &H0&
'Private Const MF_BYCOMMAND = &H0&
Private Declare Function SetWindowPos Lib "user32.dll" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long

Dim PrevWndProc As Long
Attribute PrevWndProc.VB_VarUserMemId = 1073741826
Dim mnuHandle As Long
Attribute mnuHandle.VB_VarUserMemId = 1073741827
Dim nmenuHandle As Long
Attribute nmenuHandle.VB_VarUserMemId = 1073741828
Dim hmenu11&, hmenu12&, hmenu13&, hmenu14&, hmenu15&, hmenu16&, hmenu131&
Attribute hmenu11.VB_VarUserMemId = 1073741829
Attribute hmenu12.VB_VarUserMemId = 1073741829
Attribute hmenu13.VB_VarUserMemId = 1073741829
Attribute hmenu14.VB_VarUserMemId = 1073741829
Attribute hmenu15.VB_VarUserMemId = 1073741829
Attribute hmenu16.VB_VarUserMemId = 1073741829
Attribute hmenu131.VB_VarUserMemId = 1073741829


Public Sub InitAppend_Sys_menu(hwnd As Long)
    mnuHandle = GetSystemMenu(hwnd, False)
    ' Add menu
    'Call FormOnTop(frmMain)
    Dim i As Integer
    Dim cnt As Long
    cnt = GetMenuItemCount(mnuHandle)
    RemoveMenu mnuHandle, cnt - 1, MF_BYPOSITION Or MF_REMOVE    'Remove Close Menu

    hmenu1& = CreateMenu    'main menu for mahersh mp3 player
    'add the items to the first sub menu

    hmenu11& = CreateMenu    'Play
    Result& = AppendMenu(hmenu11&, MF_STRING, 2, "File...")
    Result& = AppendMenu(hmenu11&, MF_STRING, 3, "Folder..")
    Result& = AppendMenu(hmenu11&, MF_STRING, 4, "Playlist...")
    Result& = AppendMenu(hmenu11&, MF_SEPARATOR, 0, "")
    Result& = AppendMenu(hmenu11&, MF_STRING, 5, "Removable Media")
    Result& = AppendMenu(hmenu11&, MF_SEPARATOR, 0, "")

    hmenu12& = CreateMenu    'Visualisation
    Result& = AppendMenu(hmenu12&, MF_STRING, 12, "Show Visualisation")
    Result& = AppendMenu(hmenu12&, MF_STRING, 13, "Display")
    Result& = AppendMenu(hmenu12&, MF_STRING, 14, "Configure Visualisation")

    hmenu13& = CreateMenu    'Player Control
    hmenu131& = CreateMenu    'Volume
    Result& = AppendMenu(hmenu131&, MF_STRING, 15, "+ Volume Up")
    Result& = AppendMenu(hmenu131&, MF_STRING, 16, "- Volume Down")
    Result& = AppendMenu(hmenu13&, MF_POPUP, hmenu131&, "Volume")
    Result& = AppendMenu(hmenu13&, MF_SEPARATOR, 0, "")
    Result& = AppendMenu(hmenu13&, MF_STRING, 17, "Previous Track      <P>")
    Result& = AppendMenu(hmenu13&, MF_STRING, 18, "Play       <P>")
    Result& = AppendMenu(hmenu13&, MF_STRING, 19, "Pause    <Space>")
    Result& = AppendMenu(hmenu13&, MF_STRING, 20, "Stop      <N>")
    Result& = AppendMenu(hmenu13&, MF_STRING, 21, "Next Track      <B>")
    Result& = AppendMenu(hmenu13&, MF_SEPARATOR, 0, "")
    Result& = AppendMenu(hmenu13&, MF_STRING, 22, "Seek 5 Seconds Forward    ")
    Result& = AppendMenu(hmenu13&, MF_STRING, 23, "Seek 5 Seconds Backwards  ")

    hmenu14& = CreateMenu    'Options
    Result& = AppendMenu(hmenu14&, MF_STRING, 24, "Always on Top")
    Result& = AppendMenu(hmenu14&, MF_STRING, 25, "Register Mp3 Files")
    Result& = AppendMenu(hmenu14&, MF_SEPARATOR, 0, "")
    Result& = AppendMenu(hmenu14&, MF_STRING, 26, "Intro 10 Seconds")
    Result& = AppendMenu(hmenu14&, MF_STRING, 27, "Repeat Track")
    Result& = AppendMenu(hmenu14&, MF_STRING, 28, "Mute")
    Result& = AppendMenu(hmenu14&, MF_STRING, 29, "Random")
    Result& = AppendMenu(hmenu14&, MF_SEPARATOR, 0, "")
    Result& = AppendMenu(hmenu14&, MF_STRING, 30, "Preferences")

    hmenu15& = CreateMenu    'Skins
    Result& = AppendMenu(hmenu15&, MF_STRING, 100, "<<Skin Browser>>")
    Call AppendMenu(hmenu15&, MF_SEPARATOR, 0, "")
    For i = 1 To frmPopUp.mnuSkinsAdd.Count - 1
        Result& = AppendMenu(hmenu15&, MF_STRING, 46 + i, frmPopUp.mnuSkinsAdd(i).Caption)
    Next

    hmenu16& = CreateMenu    'Opacity
    Result& = AppendMenu(hmenu16&, MF_STRING, 33, frmPopUp.mnuAlpha(0).Caption)
    For i = 1 To frmPopUp.mnuAlpha.Count - 1
        Result& = AppendMenu(hmenu16&, MF_STRING, 33 + i, frmPopUp.mnuAlpha(i).Caption)
    Next
    Call AppendMenu(hmenu16&, MF_SEPARATOR, 0, "")
    Result& = AppendMenu(hmenu16&, MF_STRING, 32 + i, "Custom..")

    Result& = AppendMenu(hmenu1&, MF_STRING, 1, "MaheshMp3 Player...")
    Result& = AppendMenu(hmenu1&, MF_SEPARATOR, 0, "")
    Call AppendMenu(hmenu1&, MF_POPUP, hmenu11&, "play")
    Call AppendMenu(hmenu1&, MF_STRING, 8, "Lyrics Viwer")
    Call AppendMenu(hmenu1&, MF_STRING, 7, "Tag Editor")
    Call AppendMenu(hmenu1&, MF_STRING, 6, "MP3 Searcher")

    Call AppendMenu(hmenu1&, MF_SEPARATOR, 0, "")
    Call AppendMenu(hmenu1&, MF_STRING, 9, "DFX Editor")
    Call AppendMenu(hmenu1&, MF_STRING, 10, "Equalisers")
    Call AppendMenu(hmenu1&, MF_STRING, 11, "Playlist Editor")
    Call AppendMenu(hmenu1&, MF_POPUP, hmenu12&, "Visualisations")

    Call AppendMenu(hmenu1&, MF_SEPARATOR, 0, "")
    Call AppendMenu(hmenu1&, MF_POPUP, hmenu13&, "Player Controls")
    Call AppendMenu(hmenu1&, MF_SEPARATOR, 0, "")

    Call AppendMenu(hmenu1&, MF_POPUP, hmenu14&, "Options")
    Call AppendMenu(hmenu1&, MF_STRING, 31, "Change Mask")
    Call AppendMenu(hmenu1&, MF_POPUP, hmenu15&, "Skins")
    Call AppendMenu(hmenu1&, MF_POPUP, hmenu16&, "Window Opacity")
    Call AppendMenu(hmenu1&, MF_SEPARATOR, 0, "")
    Call AppendMenu(hmenu1&, MF_STRING, 45, "Player Help")
    Call AppendMenu(hmenu1&, MF_SEPARATOR, 0, "")
    Call AppendMenu(hmenu1&, MF_STRING, 46, "Exit")

    'Call AppendMenu(mnuHandle, MF_SEPARATOR, 0, "")
    Call AppendMenu(mnuHandle, MF_POPUP, hmenu1&, "NeoMP3 Player")

End Sub

Public Sub Terminate(hwnd As Long)
    Call SetWindowLong(hwnd, GWL_WNDPROC, PrevWndProc)
End Sub


Private Sub FormOnTop(frm As Form)
    SetWindowPos frm.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
End Sub



Public Sub Process_SystemMenu(wParam As Long)
    Dim Result


    Select Case wParam
    Case 1:    'Mahesh MP3 Player
        frmMain.Show_Message "--MAHESH MP3 PLAYER--"
        frmAbout.Show
    Case 2:    'File
        Call frmPopUp.mnuPlayfile_Click

    Case 3:    'Folder
        Call frmPopUp.mnuPlayfolder_Click

    Case 4:    'Playlist
        Call frmPopUp.mnuPlay_playlist_Click

    Case 5:    'removable drives

    Case 6:    'Search Media files
        boolSearchShow = True
        frmSearch.Show

    Case 7:    'Tag Editor
        Call frmPopUp.mnuTagEditor_Click

    Case 8:    'lyrics
        Call frmPopUp.mnuLyrics_Click

    Case 9:    'Dsp Editor
        Call frmPopUp.mnuDSP_Click

    Case 10:    'Equalizers
        Call frmMain.Button_Click(24)

    Case 11:    'Playlist Editor
        Call frmPopUp.mnuShowPlst_Click

    Case 12:    'Show  Visualisation
        Result = GetMenuState(mnuHandle, wParam, MF_BYCOMMAND)
        If Result And MF_CHECKED Then    ' Checking Checked Menu
            CheckMenu wParam, False
        Else
            CheckMenu wParam, True
        End If
        Call frmPopUp.mnuShowVis_Click

    Case 13:    'vis

    Case 14:    'vis

    Case 15:    'volume up
        Call frmMain.Form_KeyDown(&H26, 0)
        Call frmMain.Form_KeyDown(&H26, 0)

    Case 16:    'volume down
        Call frmMain.Form_KeyDown(&H28, 0)
        Call frmMain.Form_KeyDown(&H28, 0)

    Case 17:    'Previous Track
        frmMain.Previous_Track

    Case 18:    'Play
        frmMain.Play

    Case 19:    'Pause
        frmMain.Pause_Play

    Case 20:    'Stop
        frmMain.Stop_Player

    Case 21:    'Next Track
        frmMain.Next_Track

    Case 22:    'seek
        frmMain.Five_Seg_Forward

    Case 23:    'seek
        frmMain.Five_Seg_Backward

    Case 24:    'Always on top
        Call frmPopUp.Atop_Click

    Case 25:    'Register mp3 files
        Call frmPopUp.mnuRegisterfiles_Click

    Case 26:    'Intro 10 seconds
        Call frmPopUp.mnuIntro_Click

    Case 27:    'Repeat
        Call frmPopUp.mnuRepeatTrack_Click

    Case 28:    'mute
        Call frmPopUp.mnuSilencio_Click

    Case 29:    'Random
        Call frmPopUp.mnuOrdenAleatorio_Click

    Case 30:    'preference
        Call frmPopUp.mnuPreferences_Click

    Case 31:    'Change mask
        Change_Mask False, True: frmMain.Image_State_Rep

    Case 32:    'not used  yet


    Case 33, 34, 35, 36, 37, 38, 39, 40, 41, 42:    'opacity
        frmPopUp.mnuAlpha_Click (wParam - 33)

    Case 43:    'Custom opacity

    Case 45:    'player help


    Case 46:    'exit
        frmMain.Button_Click (15)

    Case Else:
        If wParam > 46 And wParam <= 46 + frmPopUp.mnuSkinsAdd.Count - 1 Then
            Call frmPopUp.mnuSkinsAdd_Click(wParam - 46)
        End If
        '   Call FormOnTop(frmMain)
        '   MsgBox "Unchecked Menu", vbExclamation
        'Else
        '   Call CheckMenuItem(mnuHandle, &H200, MF_BYCOMMAND Or MF_CHECKED)

        '  Call FormOnTop(frmMain)
        ' MsgBox "Checked Menu", vbCritical
        '    End If
        '
        'C'ase &H201 ' Delete Menu
        '   DeleteMenu mnuHandle, &H201, MF_BYCOMMAND Or MF_DELETE
        '  MsgBox "Again Click On Caption Bar You'll see that the menu is deleted", vbInformation

        'Case &H202 ' Show frmAbout Form
        '   SetWindowPos frmAbout.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE Or SWP_SHOWWINDOW

    End Select

    'Debug.Print wParam

End Sub

Public Sub CheckMenu(mnuPos As Long, check As Boolean)
    Dim hMenuImg As Long
    Dim sFilename As String

    '................................................................
    ' I will work on it later on
    '...................................................................
    ' Get the bitmap.
    'sFileName = App.Path & "\2.bmp"
    'hMenuImg = LoadImage(0, sFileName, IMAGE_BITMAP, 0, 0, LR_LOADFROMFILE Or LR_CREATEDIBSECTION)

    ' Get the menu item handle.
    'hMenu = GetMenu(Me.hwnd)
    ' hSubMenu = GetSubMenu(hMenu, 0)

    ' Set the "mnuTwo" bitmap to the one that is loaded in memory.
    'Call SetMenuItemBitmaps(mnuHandle, 1, MF_BYPOSITION, hMenuImg, 0)
    '..........................................................................
    '.......................................................................

    If check = True Then
        Call CheckMenuItem(mnuHandle, mnuPos, MF_BYCOMMAND Or MF_CHECKED)
    Else
        Call CheckMenuItem(mnuHandle, mnuPos, MF_BYCOMMAND Or MF_UNCHECKED)
    End If
End Sub

Public Sub RenameSystemMenu(mnuPos As Long, sName As String)
'Call ModifyMenu(hmenu131&, 0, MF_CHANGE, 0, sName)
End Sub
