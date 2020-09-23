VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDialog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private Declare Function GetOpenFileName Lib "COMDLG32.DLL" Alias "GetOpenFileNameA" (pOpenfilename As OPENFILENAME) As Long
Private Declare Function GetSaveFileName Lib "COMDLG32.DLL" Alias "GetSaveFileNameA" (pOpenfilename As OPENFILENAME) As Long
'Public Declare Function GetFileVersionInfoSize Lib "Version.dll" Alias "GetFileVersionInfoSizeA" (ByVal lptstrFilename$, lpdwHandle&) As Long
'Public Declare Function GetFileVersionInfo Lib "Version.dll" Alias "GetFileVersionInfoA" (ByVal lptstrFilename$, ByVal dwHandle&, ByVal dwLen&, lpData As Any) As Long
'Private Declare Function VerQueryValue Lib "Version.dll" Alias "VerQueryValueA" (pBlock As Any, ByVal lpSubBlock As String, lplpBuffer As Any, puLen As Long) As Long
'*******************
Private Declare Function SHBrowseForFolder Lib "shell32" (lpBI As BrowseInfo) As Long
Private Declare Function SHGetPathFromIDList Lib "shell32" (ByVal pidList As Long, ByVal lpBuffer As String) As Long

Private Declare Sub CoTaskMemFree Lib "ole32.dll" (ByVal hMem As Long)
Private Declare Function lstrcat Lib "kernel32" Alias "lstrcatA" (ByVal lpString1 As String, ByVal lpString2 As String) As Long


Private cdlg As OPENFILENAME
Private LastFileName As String
Private Type OPENFILENAME
    lStructSize As Long
    hwndOwner As Long
    hInstance As Long
    lpstrFilter As String
    lpstrCustomFilter As String
    nMaxCustFilter As Long
    nFilterIndex As Long
    lpstrFile As String
    nMaxFile As Long
    lpstrFileTitle As String
    nMaxFileTitle As Long
    lpstrInitialDir As String
    lpstrTitle As String
    Flags As Long
    nFileOffset As Integer
    nFileExtension As Integer
    lpstrDefExt As String
    lCustData As Long
    lpfnHook As Long
    lpTemplateName As String
End Type

Private Const OFN_ALLOWMULTISELECT = &H200
Private Const OFN_CREATEPROMPT = &H2000
Private Const OFN_ENABLEHOOK = &H20
Private Const OFN_ENABLETEMPLATE = &H40
Private Const OFN_ENABLETEMPLATEHANDLE = &H80
Private Const OFN_EXPLORER = &H80000                         '  new look commdlg
Private Const OFN_EXTENSIONDIFFERENT = &H400
Private Const OFN_FILEMUSTEXIST = &H1000
Private Const OFN_HIDEREADONLY = &H4
Private Const OFN_LONGNAMES = &H200000                       '  force long names for 3.x modules
Private Const OFN_NOCHANGEDIR = &H8
Private Const OFN_NODEREFERENCELINKS = &H100000
Private Const OFN_NOLONGNAMES = &H40000                      '  force no long names for 4.x modules
Private Const OFN_NONETWORKBUTTON = &H20000
Private Const OFN_NOREADONLYRETURN = &H8000
Private Const OFN_NOTESTFILECREATE = &H10000
Private Const OFN_NOVALIDATE = &H100
Private Const OFN_OVERWRITEPROMPT = &H2
Private Const OFN_PATHMUSTEXIST = &H800
Private Const OFN_READONLY = &H1
Private Const OFN_SHAREAWARE = &H4000
Private Const OFN_SHAREFALLTHROUGH = 2
Private Const OFN_SHARENOWARN = 1
Private Const OFN_SHAREWARN = 0
Private Const OFN_SHOWHELP = &H10

Public Enum DialogFlags
    ALLOWMULTISELECT = OFN_ALLOWMULTISELECT
    CREATEPROMPT = OFN_CREATEPROMPT
    ENABLEHOOK = OFN_ENABLEHOOK
    ENABLETEMPLATE = OFN_ENABLETEMPLATE
    ENABLETEMPLATEHANDLE = OFN_ENABLETEMPLATEHANDLE
    EXPLORER = OFN_EXPLORER
    EXTENSIONDIFFERENT = OFN_EXTENSIONDIFFERENT
    FILEMUSTEXIST = OFN_FILEMUSTEXIST
    HIDEREADONLY = OFN_HIDEREADONLY
    LONGNAMES = OFN_LONGNAMES
    NOCHANGEDIR = OFN_NOCHANGEDIR
    NODEREFERENCELINKS = OFN_NODEREFERENCELINKS
    NOLONGNAMES = OFN_NOLONGNAMES
    NONETWORKBUTTON = OFN_NONETWORKBUTTON
    NOREADONLYRETURN = OFN_NOREADONLYRETURN
    NOTESTFILECREATE = OFN_NOTESTFILECREATE
    NOVALIDATE = OFN_NOVALIDATE
    OVERWRITEPROMPT = OFN_OVERWRITEPROMPT
    PATHMUSTEXIST = OFN_PATHMUSTEXIST
    'ReadOnly = OFN_READONLY
    SHAREAWARE = OFN_SHAREAWARE
    SHAREFALLTHROUGH = OFN_SHAREFALLTHROUGH
    SHARENOWARN = OFN_SHARENOWARN
    SHAREWARN = OFN_SHAREWARN
    ShowHelp = OFN_SHOWHELP
End Enum

'BrowseInfo ulFlags
Private Const BIF_RETURNONLYFSDIRS = &H1         'Only return file system directories.
Private Const BIF_DONTGOBELOWDOMAIN = &H2       'Do not include network folders below the domain level in the dialog box's tree view control.
Private Const BIF_STATUSTEXT = &H4              'Include a status area in the dialog box.
Private Const BIF_RETURNFSANCESTORS = &H8       'Only return file system ancestors. An ancestor is a subfolder that is beneath the root folder in the namespace hierarchy.
Private Const BIF_BROWSEFORCOMPUTER = &H1000    'Only return computers.
Private Const BIF_BROWSEFORPRINTER = &H2000     'Only return network printers.
Private Const BIF_BROWSEINCLUDEFILES = &H4000   '(SHELL32.DLL Version 4.71). The browse dialog will display files as well as folders.
Private Const BIF_SHAREABLE = &H8000            '(SHELL32.DLL Version 5.0). The browse dialog box can display shareable resources on remote systems. The BIF_USENEWUI flag must also be set.

'BrowseInfo pIDLRoot(Do not use these with new style dialog)
Const Default = 0
Const Internet = 1
Const Programs = 2
Const ControlPanel = 3
Const Printers = 4
Const MyDocuments = 5
Const Favorites = 6
Const StartUp = 7
Const Recent = 8
Const SendTo = 9
Const RecycleBin = 10
Const StartMenu = 11
Const Desktop = 16
Const MyComputer = 17
Const Network = 18
Const Nethood = 19
Const Fonts = 20
Const Templates = 21
Const ApplicationData = 26
Const PrintHood = 27
Const TemporaryInternetFiles = 32
Const Cookies = 33
Const History = 34

Const BFFM_ENABLEOK = &H465
Const BFFM_SETSELECTION = &H466
Const BFFM_SETSTATUSTEXT = &H464

Const BFFM_INITIALIZED = 1
Const BFFM_SELCHANGED = 2
Const BFFM_VALIDATEFAILED = 3

Const MAX_PATH = 260

Private Type BrowseInfo
    hwndOwner As Long
    pIDLRoot As Long
    pszDisplayName As String
    lpszTitle As String
    ulFlags As Long
    lpfnCallBack As Long
    lParam As Long
    iImage As Integer
End Type

Private CurrentSelection As String * MAX_PATH
Public DIALOGUE_OPEN As Boolean

'Function'   FileOpen
'            Display standard OpenSaveDialogue
'
' Parameters'
'            Form_hWnd       - Parent's Handle
'            Title           - Dialogue Title
'            Filter          - Specify Filter as with FileSelectFile
'                              for use of multiple Filter seperate them with |
'            IniDir          - Specify Initialisation Directory
'                              chosen Directory will be set as WorkingDir in
'                              A_WorkingDir
'            flags           - Flags for Dialogue
'                              when no supplied
'                               OFN_FILEMUSTEXIST - OFN_HIDEREADONLY  - OFN_ENABLEXPLORER
'                              are used as default
'
'  Returns'
'            Selected FileName or Emtpy when cancelled
'
'  Remarks'
'            When OFN_ALLOWMULTISELECT flag is set, be sure to set OFN_ENABLEXPLORER flag,
'            too. Without it the Old-Style FileOpen Dialogue appears and handling multiple
'            fileNames might be different and wont work as expected
'            See http'//msdn2.microsoft.com/en-us/library/ms646839.aspx for details

Public Function GetOpenAsName(ByVal Form_hWnd As Long, ByVal Title As String, Optional ByVal InitDir As String = "", Optional ByVal Filter As String = "All Files (*.*)|*.*|", Optional Flags As DialogFlags = FILEMUSTEXIST) As String

    Dim lpOFN As OPENFILENAME, sTemp$, nStrEnd&
    If DIALOGUE_OPEN = True Then GetOpenAsName = "": Exit Function

    Dim i As Integer
    Filter = Replace(Filter, "|", Chr(0))
    If Right(Filter, 1) <> Chr(0) Then Filter = Filter & Chr(0)

    If Len(InitDir) = 0 Then InitDir = LastFileName
    ' initialize the struct params

    With lpOFN
        DIALOGUE_OPEN = True
        .lStructSize = Len(lpOFN)
        ' if the 2K version of the common dialog dll is not present, subtract the byte count for the
        ' last three members of the struct
        'If Is2KShell() = False Then .lStructSize = .lStructSize - 12
        .lpstrTitle = Title
        .hwndOwner = Form_hWnd
        .hInstance = App.hInstance
        .lpstrFilter = Filter
        .lpstrFile = Space(254)
        .nMaxFile = 255
        .lpstrFileTitle = Space(254)
        .nMaxFileTitle = 255
        .lpstrInitialDir = InitDir
        .lpstrTitle = Title
        .Flags = Flags Or FILEMUSTEXIST Or PATHMUSTEXIST Or OFN_LONGNAMES
    End With

    If GetOpenFileName(lpOFN) Then
        sTemp = lpOFN.lpstrFile
        nStrEnd = InStr(sTemp, vbNullChar)
        If nStrEnd > 1 Then
            GetOpenAsName = Left$(sTemp, nStrEnd - 1)
        Else
            GetOpenAsName = vbNullString
        End If
    Else
        GetOpenAsName = vbNullString
    End If
    DIALOGUE_OPEN = False
    If Len(GetOpenAsName) > 0 Then LastFileName = GetOpenAsName

End Function

Public Property Let FileName(szFileName)
    LastFileName = szFileName
End Property

Public Property Get FileName()
    FileName = LastFileName
End Property

Public Function GetSaveAsName(ByVal Form_hWnd As Long, ByVal Title As String, Optional ByVal InitDir As String = "", Optional ByVal Filter As String = "All Files (*.*)|*.*|", Optional ByVal DefExt As String = "", Optional Flags As DialogFlags = OVERWRITEPROMPT) As String

    Dim lpOFN As OPENFILENAME, sTemp$, nStrEnd&
    If DIALOGUE_OPEN = True Then GetSaveAsName = "": Exit Function

    Filter = Replace(Filter, "|", Chr(0))
    If Right(Filter, 1) <> Chr(0) Then Filter = Filter & Chr(0)
    If Len(InitDir) = 0 Then InitDir = LastFileName
    ' initialize the struct params

    With lpOFN
        DIALOGUE_OPEN = True
        .lStructSize = Len(lpOFN)

        ' if the 2K version of the common dialog dll is not present, subtract the byte count for the
        ' last three members of the struct
        ' If Is2KShell() = False Then .lStructSize = .lStructSize - 12

        .lpstrTitle = Title
        .hwndOwner = Form_hWnd
        .hInstance = App.hInstance
        .lpstrFilter = Filter
        .lpstrDefExt = DefExt
        .lpstrFile = Space(254)  '1 less than .nmaxfile
        .nMaxFile = 255
        .lpstrFileTitle = Space(254)    '1 less than .nmaxfiletitle
        .nMaxFileTitle = 255
        .lpstrInitialDir = InitDir
        .Flags = Flags

        If GetSaveFileName(lpOFN) Then
            sTemp = lpOFN.lpstrFile
            nStrEnd = InStr(sTemp, vbNullChar)
            If nStrEnd > 1 Then
                GetSaveAsName = Left$(sTemp, nStrEnd - 1)
            Else
                GetSaveAsName = vbNullString
            End If
        Else
            GetSaveAsName = vbNullString
        End If
    End With
    DIALOGUE_OPEN = False
    ' If .Flags And OFN_EXTENSIONDIFFERENT Then MsgBox "choose correct ext"

    If Len(GetSaveAsName) > 0 Then LastFileName = GetSaveAsName

End Function

Public Sub OpenAsMultiFileName(ByVal Form_hWnd As Long, ByVal Title As String, ByRef sfiles() As String, Optional ByVal InitDir As String = "", Optional ByVal Filter As String = "All Files (*.*)|*.*|", Optional ByVal DefExt As String = "", Optional Flags As DialogFlags = OVERWRITEPROMPT Or OFN_EXTENSIONDIFFERENT)
    On Error Resume Next
    If DIALOGUE_OPEN = True Then ReDim sfiles(0 To 1) As String: sfiles(0) = "": sfiles(1) = "": Exit Sub

    Filter = Replace(Filter, "|", Chr(0))
    If Right(Filter, 1) <> Chr(0) Then Filter = Filter & Chr(0)

    If Len(InitDir) = 0 Then InitDir = LastFileName

    Dim lpOFN As OPENFILENAME, sTemp1, nStrEnd&

    ' initialize the struct params
    DIALOGUE_OPEN = True
    With lpOFN
        .lStructSize = Len(lpOFN)
        ' if the 2K version of the common dialog dll is not present, subtract the byte count for the
        ' last three members of the struct
        ' If Is2KShell() = False Then .lStructSize = .lStructSize - 12
        .hwndOwner = frmPopUp.hwnd

        ' tell it we want to display all files.  filter strings are explained in
        ' the OPENFILENAME documentation in the MSDN
        .lpstrFilter = "MP3 Files (*.mp3)" & vbNullChar & "*.mp3" & vbNullChar & "All Files (*.*)" & vbNullChar & "*.*" & vbNullChar & vbNullChar
        .lpstrFile = String$(10700, 0)
        .nMaxFile = 10700
        .lpstrFileTitle = String$(1260, 0)
        .nMaxFileTitle = 1260
        .lpstrInitialDir = sInitialDir
        .lpstrTitle = "Open Music File for MaheshMp3"
        .Flags = OFN_LONGNAMES Or OFN_EXPLORER Or OFN_ALLOWMULTISELECT

        '  enable the hook proc and change the view before the dialog is displayed
        '  Set .Flags = .Flags Or OFN_ENABLEHOOK
        '  .lpfnHook = ReturnProcAddress(AddressOf WindowProc)
    End With

    If GetOpenFileName(lpOFN) Then
        sTemp1 = lpOFN.lpstrFile
        nStrEnd = InStr(sTemp1, vbNullChar)
    End If
    DIALOGUE_OPEN = False

    If Trim(lpOFN.lpstrFile = "") Then ReDim sfiles(0 To 1) As String: sfiles(0) = "": sfiles(1) = "": Exit Sub

    'Remove trailing null characters.
    Dim nDoubleNullPos As Long
    nDoubleNullPos = InStr(lpOFN.lpstrFile & vbNullChar, String$(2, 0))
    Dim sfullname, sPath, sExtension As String
    Dim sfiletitle As String
    If nDoubleNullPos Then
        'Get the file name including the path name.
        sfullname = Left$(lpOFN.lpstrFile, nDoubleNullPos - 1)
        If sfullname = "" Then Exit Sub
        'Get the file name without the path name.
        sfiletitle = Left$(lpOFN.lpstrFileTitle, InStr(lpOFN.lpstrFileTitle, vbNullChar) - 1)
        'Get the path name.
        sPath = Left$(sfullname, lpOFN.nFileOffset - 1)
        sInitialDir = sPath
        'Get the extension.
        If lpOFN.nFileExtension Then
            sExtension = Mid$(sfullname, lpOFN.nFileExtension + 1)
        End If
        DoEvents
        'If sFileTitle is a string, we have a single selection.
        If Len(sfiletitle) Then
            'Add to the collections.
            Dim m As String
            m = Mid$(sfullname, lpOFN.nFileOffset + 1)
            m = sfullname
            ReDim sfiles(0 To 1) As String
            'frmPLST.Add_track_to_Playlist (m)
            sfiles(0) = sPath
            sfiles(1) = sfiletitle

        Else    'Tear multiple selection apart.
            Dim sTemp() As String
            Dim nCount As Long
            sfiles = Split(sfullname, vbNullChar)
            'If array contains no elements, UBound returns -1.
            If UBound(sfiles) > LBound(sfiles) Then
                'We have more than one array element!
                'Remove backslash if sPath is the root folder.
                If Len(sPath) = 3 Then sPath = Left$(sPath, 2)
                'Loop through the array, and create the
                'collections; skip the first element
                '(containing the path name), so start the
                'counter at 1, not at 0.
                For nCount = 1 To UBound(sfiles)
                    'frmPLST.Add_track_to_Playlist (sPath + "\" + sfiles(nCount))
                    'If the string already contains a backslash,
                    'the user must have selected a shortcut
                    'file, so we don't add the path.
                    'colFileNames.Add IIf(InStr(sTemp(nCount),sBackSlash), sTemp(nCount), sPath & sBackSlash & sTemp(nCount))
                Next nCount
                'Clear this variable.
                'sfullname = vbNullString
            End If
        End If
        'Add backslash if sPath is the root folder.
        If Len(sPath) = 2 Then
            'spath = spath & sBackSlash
        End If

2:
    End If

End Sub


Public Function FolderBrowse(ByVal lngHwnd As Long, ByVal strMsg As String, Optional lFlags As Long) As String
    On Error GoTo hell
    Dim intNull As Integer
    Dim lngIDList As Long, lngResult As Long
    Dim strPath As String
    Dim udtBI As BrowseInfo
    '// Set the properties of the api with the appropriate structure
    With udtBI
        .hwndOwner = lngHwnd
        .lpszTitle = strMsg    'lstrcat(strMsg, "")
        .ulFlags = lFlags Or BIF_RETURNONLYFSDIRS
    End With
    '// show Explorer ...
    lngIDList = SHBrowseForFolder(udtBI)
    DIALOGUE_OPEN = False
    If lngIDList <> 0 Then
        '//Create string of 260 zero space to store the address / path
        strPath = String(MAX_PATH, 0)

        '// Store the selected address strPath
        lngResult = SHGetPathFromIDList(lngIDList, strPath)

        '// free memory from taskmanager
        Call CoTaskMemFree(lngIDList)

        '// Find the first null character in the string
        intNull = InStr(strPath, vbNullChar)
        '//Ensure that the goods
        If intNull > 0 Then
            strPath = Left(strPath, intNull - 1)
        End If
    End If

    '// Return the directory name
    FolderBrowse = strPath
    Exit Function

hell:
    '// Return null on error
    FolderBrowse = Empty

End Function

'///////////////////
'+-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-+
'|  FUNCION TO CHECK cmdlgversion                                             |
'+-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-+


'Public Function Is2KShell() As Boolean
' this function returns the version of the Comdlg32.dll on the system
' this info is used to determine which version of the OPENFILENAME struct
' should be passed to the dialog functions

'Dim nBuffSize&, nDiscard&, lpBuffer&, nVerMajor&, abytBuffer() As Byte

'Const FILE_NAME As String = "Comdlg32.dll"

'nBuffSize = GetFileVersionInfoSize(FILE_NAME, nDiscard)

'If nBuffSize > 0 Then
' ReDim abytBuffer(nBuffSize - 1) As Byte

'Call GetFileVersionInfo(FILE_NAME, 0&, nBuffSize, abytBuffer(0))

'If VerQueryValue(abytBuffer(0), "\", lpBuffer, nDiscard) Then
'CopyMemory nVerMajor, ByVal lpBuffer + 10, 2&

'If nVerMajor >= 5 Then Is2KShell = True
'End If
' End If

'End Function
'+-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-+
'|  FUNCION PARA EXPLORADOR DE DIRECTORIOS                                                |
'+-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-+

