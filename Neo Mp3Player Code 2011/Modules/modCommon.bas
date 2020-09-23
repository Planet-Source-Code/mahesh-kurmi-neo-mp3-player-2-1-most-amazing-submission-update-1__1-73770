Attribute VB_Name = "modFileSystem"
Option Explicit

'
' module containing commonly used functions and subroutines
'

' last updated: 2011 Feb 12

Global Fsys As New FileSystemObject         'global FileSystemObject

Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Declare Function SendMessage Lib "user32" Alias "SendMessageA" _
                             (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, _
                              lParam As Any) As Long
Public Const FILE_ATTRIBUTE_DIRECTORY As Long = &H10
Public Declare Function ShellExecute Lib "shell32" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Public Declare Function SetCapture Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function InitCommonControls Lib "comctl32" () As Long
Public Declare Function FindNextFile Lib "kernel32.dll" Alias "FindNextFileA" (ByVal hFindFile As Long, lpFindFileData As WIN32_FIND_DATA) As Long
Public Declare Function FindClose Lib "kernel32" (ByVal hFindFile As Long) As Long
Public Declare Function FindFirstFile Lib "kernel32" Alias "FindFirstFileA" (ByVal lpFileName As String, lpFindFileData As WIN32_FIND_DATA) As Long
Public Declare Function GetFileAttributes Lib "kernel32" Alias "GetFileAttributesA" (ByVal lpFileName$) As Long
Private Declare Function GetFileSize Lib "kernel32" (ByVal hFile As Long, lpFileSizeHigh As Long) As Long


Public Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, lpvParam As Any, ByVal fuWinIni As Long) As Long

Public Const SPI_GETWORKAREA = 48
Const CB_FINDSTRING = &H14C
Const CB_FINDSTRINGEXACT = &H158
Const LB_FINDSTRING = &H18F
Const LB_FINDSTRINGEXACT = &H1A2

Public Type FILETIME
    dwLowDateTime As Long
    dwHighDateTime As Long
End Type
Public Type WIN32_FIND_DATA
    dwFileAttributes As Long
    ftCreationTime As FILETIME
    ftLastAccessTime As FILETIME
    ftLastWriteTime As FILETIME
    nFileSizeHigh As Long
    nFileSizeLow As Long
    dwReserved0 As Long
    dwReserved1 As Long
    cFileName As String * 260
    cAlternate As String * 14
End Type
Public infoStr As String, Pos As Single, infoStr1 As String
Attribute Pos.VB_VarUserMemId = 1073741825
Attribute infoStr1.VB_VarUserMemId = 1073741825
Public InfoParts() As String, currentPart As Integer
Attribute InfoParts.VB_VarUserMemId = 1073741828
Attribute currentPart.VB_VarUserMemId = 1073741828

Public sSize As Single, tSize As Single
Attribute sSize.VB_VarUserMemId = 1073741830
Attribute tSize.VB_VarUserMemId = 1073741830
'flag to stop scanning for media
Public stopScan As Boolean
Attribute stopScan.VB_VarUserMemId = 1073741832

' type to encapsulate errors
Public Type ErrStruct
    errNum As Long
    errShortDesc As String
    errLongDesc As String
End Type


Public Type FireAMPoptions
    ' general
    enableVisualizations As Byte

    'start up
    showSplashScreen As Byte
    loadDefaultSkin As Byte
    checkAssociationsAtStartUp As Byte
    defaultSepChar As Byte
    ' file types
    MIDI As Byte
    WAV As Byte
    MP3 As Byte
    MPG As Byte
    WMA As Byte

End Type
Public theOptions As FireAMPoptions
Attribute theOptions.VB_VarUserMemId = 1073741833
Public sepChar As Byte, SelectedIndex As Integer
Attribute sepChar.VB_VarUserMemId = 1073741834
Attribute SelectedIndex.VB_VarUserMemId = 1073741834

Public sleepFactor As Byte
Attribute sleepFactor.VB_VarUserMemId = 1073741836
Public oldPath As String
Attribute oldPath.VB_VarUserMemId = 1073741837

Public Const HTTPGET = "GET {PATH} HTTP/1.1" & vbCrLf & _
       "Host: {HOST}" & vbCrLf & _
       "Accept: */*" & vbCrLf & _
       "User-Agent: FireAMP Update System" & vbCrLf & _
       vbCrLf
Public Const HTTPGETRANGE = "GET {PATH} HTTP/1.1" & vbCrLf & _
       "Range: bytes={RANGE}" & vbCrLf & _
       "Host: {HOST}" & vbCrLf & _
       "Accept: */*" & vbCrLf & _
       "User-Agent: FireAMP Update System" & vbCrLf & _
       vbCrLf
Private Const MAX_PATH As Long = 260&
Private Const API_FALSE As Long = 0&
Private Const INVALID_HANDLE_VALUE As Long = (-1&)
Private Const ERROR_NO_MORE_FILES As Long = 18&
Public clsDlg As New clsDialog
Attribute clsDlg.VB_VarUserMemId = 1073741838
Public xFlags As DialogFlags
Attribute xFlags.VB_VarUserMemId = 1073741839

'subroutine to log errors
Public Sub logError(theError As ErrStruct)

    Dim Fout As TextStream
    'With frmFireTrap
    '.lblError = theError.errShortDesc
    '.lblReason = theError.errLongDesc
    '.lblNum = "Error #" & theError.errNum
    'End With
    'frmFireTrap.Show vbModal

    ' log error to file
    Set Fout = Fsys.OpenTextFile(App.Path & "\FireAMP.Errors.Log", ForAppending, True)
    If Fsys.GetFile(App.Path & "\FireAMP.Errors.Log").Size > 10& * 1024& Then    ' greater than 10kb
        Fout.Close
        Kill App.Path & "\FireAMP.Errors.Log"
        Set Fout = Fsys.OpenTextFile(App.Path & "\FireAMP.Errors.Log", ForAppending, True)
    End If
    Fout.WriteLine "FireAMP error #" & theError.errNum
    Fout.WriteLine "Error occured on: " & Now
    Fout.WriteLine "Short Desc: " & theError.errShortDesc
    Fout.WriteLine "Long Desc: " & theError.errLongDesc
    Fout.WriteLine String(40, "-")
    Fout.Close

End Sub

' function to check if the given char is valid or not
Private Function isAllowedChar(testStr As String) As Boolean
    isAllowedChar = InStr(1, "ABCDEFGHIJKLMNOPQRSTUVWXYZ abcdefghijklmnopqrstuvwxyz1234567890(){}[]!;:'"",.$%*+#|\/~&", testStr)
End Function

' function to remove unwnated chars
Public Function toStdString(theString As String) As String
    Dim retStr As String, i As Integer, j As Integer
    retStr = Space(Len(theString))    ' fill up return string
    Let j = 1
    For i = 1 To Len(theString)

        If isAllowedChar(Mid(theString, i, 1)) Then
            ' mid is more faster and efficient than '&'
            Mid(retStr, j, 1) = Mid(theString, i, 1)
            j = j + 1
        End If
    Next i

    toStdString = Trim(retStr)

End Function

' bunch of useful functions
Public Function getFileNameFromPath(dirPath As String) As String
    getFileNameFromPath = Right(dirPath, Len(dirPath) - InStrRev(dirPath, "\"))
End Function
Public Function getDirFromPath(dirPath As String) As String
    getDirFromPath = Left(dirPath, InStrRev(dirPath, "\") - 1)
End Function

Public Function getFileTitleFromPath(FilePath As String) As String
    Dim Temp As String
    Temp = Mid(FilePath, Len(Left(FilePath, InStrRev(FilePath, "\"))) + 1)
    getFileTitleFromPath = Left(Temp, InStrRev(Temp, ".") - 1)
End Function

Public Function getFileExtensionFromPath(FilePath As String) As String
    getFileExtensionFromPath = Trim(LCase(Right(FilePath, Len(FilePath) - InStrRev(FilePath, "."))))
    'If InStrRev(FilePath, ".") = 0 Then getFileExtensionFromPath = ""
End Function

Public Function getFolderTitleFromPath(folderPath As String) As String

    getFolderTitleFromPath = Mid(folderPath, InStrRev(folderPath, "\") + 1)
End Function

' function to convert seconds to HH:MM:SS format
Public Function convertToStdTime(ByVal Seconds As Long) As String
    On Error Resume Next
    'Format input value to "00:00:00"
    Dim HH As Long                   'Hours
    Dim MM As Long                   'Minutes
    Dim ss As Long                   'Seconds
    Dim Tmp As String                'Temporary value

    'Old values time is made of
    HH = Seconds \ 3600
    MM = Seconds \ 60 Mod 60
    ss = Seconds Mod 60

    'If there is hour
    If HH > 0 Then Tmp = Format$(HH, "00:")
    'Format input
    convertToStdTime = Tmp & Format$(MM, "00:") & Format$(ss, "00")
End Function

' function to translate the bar position into a position value
Public Function getBarPosition(picBar As PictureBox, picBarBack As PictureBox, iMax As Single) As Single
    getBarPosition = ((picBar.Left) * iMax) / (picBarBack.ScaleWidth - picBar.ScaleWidth)

End Function

'updates the seek bar
Public Sub updateBar(picBar As PictureBox, picBarBack As PictureBox, ByVal iMax As Double, ByVal ipos As Double)
    On Error Resume Next
    'Set the position bar to player position
    picBar.Move ((picBarBack.ScaleWidth - picBar.ScaleWidth) * ((ipos) / iMax))    ' bar position depends on the maximum value
    picBarBack.CurrentX = picBarBack.ScaleWidth / 2
    picBarBack.CurrentY = 1
    picBar.CurrentX = picBar.ScaleWidth / 2
    picBar.CurrentY = 1
    DoEvents
End Sub




' general purpose sub to scan a path for media files, can display progress
' and status in picture boxes and labels

Sub scanFolder(FolderSpec As String, lstPaths As ListBox, lstPl As ListView, Optional lblBar As PictureBox = Nothing, Optional lblBarBack As PictureBox = Nothing, Optional lblStatus As Label = Nothing, Optional lblData As Label = Nothing)
    On Error GoTo e
    DoEvents
    Dim i As Integer

    Dim thisFolder As Folder
    Dim sFolders As Folders
    Dim fileItem As File, folderItem As Folder
    Dim allFiles As Files

    Set thisFolder = Fsys.GetFolder(FolderSpec)
    Set sFolders = thisFolder.SubFolders
    Set allFiles = thisFolder.Files

    If stopScan Then Exit Sub

    For Each folderItem In sFolders
        DoEvents
        If Not lblData Is Nothing Then lblData.Caption = "Looking in:" & vbNewLine & vbNewLine & folderItem.Path
        scanFolder folderItem.Path, lstPaths, lstPl, lblBar, lblBarBack, lblStatus, lblData

    Next

    For Each fileItem In allFiles
        sSize = sSize + fileItem.Size
        If isMediaFile(fileItem.Path) Then
            'lstPaths.AddItem fileItem.Path
            'lstPl.ListItems.Add , , getFileTitleFromPath(fileItem.Path)

        End If
    Next
    DoEvents
    If Not lblBar Is Nothing And Not lblBarBack Is Nothing Then updateBar lblBar, lblBarBack, tSize, sSize
    If Not lblStatus Is Nothing Then lblStatus.Caption = "Scanned " & Round(sSize / (1024! * 1024!)) & "MB of " & Round(tSize / (1024! * 1024!)) & "MB so far."
    Exit Sub
e:

End Sub

' function used by ScanFolder to determine if the file is a media file
Public Function isMediaFile(FilePath As String) As Boolean
    Dim ext As String

    ext = getFileExtensionFromPath(FilePath)
    'isMediaFile = CBool(InStr("mp3 mp2 wma wmv mpg mpeg mpe rm rmvb mid rmi avi mov", ext))
    If UCase(ext) = "MP3" Or UCase(ext) = "MP2" Or UCase(ext) = "WAV" Or UCase(ext) = "WMA" Or UCase(ext) = "OGG" Then isMediaFile = True

End Function

Public Function isVideoFile(FilePath As String)
    Dim ext As String

    ext = getFileExtensionFromPath(FilePath)
    isVideoFile = CBool(InStr("wmv mpg mpeg mpe rm rmvb avi dat mov", ext))
End Function

' function used by ScanFolder to determine if the file is a media file
Public Function isPlaylistFile(FilePath As String) As Boolean
    Dim ext As String

    ext = LCase(getFileExtensionFromPath(FilePath))
    'isMediaFile = CBool(InStr("mp3 mp2 wma wmv mpg mpeg mpe rm rmvb mid rmi avi mov", ext))
    isPlaylistFile = CBool(InStr("npl m3u pls", ext))

End Function

Public Function parseString(Src As String, Start As Integer, Finish As Integer) As String
    On Error Resume Next
    Dim str As String, str1 As String
    Src = Trim(Src)
    str = Left(Src, Len(Src) - Start)
    str1 = Right(str, Len(str) - Finish)
    parseString = str1
End Function



'Description: loads playlists into frmplst (.m3u,.pls,.npl)
'Parameter: Openpath: Path of playlist file to be openned   bNewlist: If Playlist is to be added as new list or to be enqueued
'           Insert_at: position where all tracks of playlist are to be inserted(Defaualt=-1 if tracks are to be inserted at last
'Return: No. of tracks added
Public Function LoadPlaylist(Optional openPath As String, Optional bNewlist As Boolean = False, Optional Insert_at As Long = -1) As Long
    On Error Resume Next
    Dim tTrack As FileTrack
    Dim i As Integer
    i = 1
    If bLoading = True And openPath = "" Then openPath = App.Path + "\mplayerlist.npl"    'load playlist when loading application
    If bNewlist = True Then frmPLST.ClearList

    Dim sExtension As String
    'Get file extension
    sExtension = Trim(LCase(Right(openPath, Len(openPath) - InStrRev(openPath, "."))))
    Select Case sExtension
    Case "npl"
        Open openPath For Random As #5 Len = 255

        Do While (1)
            Get #5, i, tTrack
            If tTrack.trackPath = "" Then Exit Do
            frmPLST.clist.AddItem tTrack.trackName, tTrack.trackPath, tTrack.Duration, , IIf(Insert_at = -1, -1, Insert_at + i)
            i = i + 1
        Loop
        Close #5

    Case "m3u"
        Dim sBuff As String, M3UChk As String * 7

        '// Check for M3U Header
        Open openPath For Binary As #1
        Get 1#, 1, M3UChk
        If M3UChk <> "#EXTM3U" Then Exit Function
        Close #1
        DoEvents
        '// Adding procedure
        i = 0    'variable for insert position of track
        Open openPath For Input As #1
        Do While Not EOF(1)
            Line Input #1, sBuff
            'If IsBlank(sBuff) Then GoTo 1
            If Mid(sBuff, 1, 1) = "#" Then GoTo 1
            frmPLST.clist.AddItem GetFileTitle(sBuff), sBuff, , , IIf(Insert_at = -1, -1, Insert_at + i)
            i = i + 1
1
        Loop
        Close #1
    Case "pls"
        Dim lMax As Long
        Dim sFile As String, sTitle As String, sLength As String
        '// Check the file
        If FileExists(openPath) = False Then Exit Function

        '// Get # of files in this playlist
        lMax = CLng(GetFromINI("playlist", "NumberOfEntries", openPath, 0))
        If lMax = 0 Then Exit Function

        For i = 0 To lMax    '// Get the files INI values
            sFile = GetFromINI("playlist", "File" & i + 1, openPath, "")
            If IsBlank(sFile) Then GoTo 2
            sTitle = GetFromINI("playlist", "Title" & i + 1, openPath, GetFileTitle(sFile))
            sLength = GetFromINI("playlist", "Length" & i + 1, openPath, "")
            'I slength is zero then trk is either corrupt or hasn't been read
            sLength = IIf(sLength = "" Or sLength = "0", "", Convert_Time_to_string(CLng(sLength)))
            frmPLST.clist.AddItem sTitle, sFile, sLength, , IIf(Insert_at = -1, -1, Insert_at + i)
2
        Next

    End Select
3:
    frmPLST.Update_Plst_Scrollbar
    LoadPlaylist = i
    If bNewlist = True And frmPLST.clist.ItemCount > 0 Then CurrentTrack_Index = 0: sFileMainPlaying = frmPLST.clist.exItem(CurrentTrack_Index): frmMain.Play
errorHandler:
End Function

Public Function FormatFileSize( _
       ByVal dblFileSize As Double, _
       Optional ByVal strFormatMask As String _
       ) As String

    Select Case dblFileSize
    Case 0 To 1023    ' Bytes
        FormatFileSize = Format(dblFileSize) & " bytes"

    Case 1024 To 1048575    ' KB
        If strFormatMask = Empty Then strFormatMask = "###0"
        FormatFileSize = Format(dblFileSize / 1024#, strFormatMask) & " KB"

    Case 1024# ^ 2 To 1073741823    ' MB
        If strFormatMask = Empty Then strFormatMask = "###0.0"
        FormatFileSize = Format(dblFileSize / (1024# ^ 2), strFormatMask) & " MB"

    Case Is > 1073741823#    ' GB
        If strFormatMask = Empty Then strFormatMask = "###0.0"
        FormatFileSize = Format(dblFileSize / (1024# ^ 3), strFormatMask) & " GB"
    End Select
End Function

'///GET MAXIMUM OF 2
Public Function maxX( _
       ByVal val1 As Long, _
       ByVal val2 As Long _
       ) As Long

    maxX = IIf(val1 > val2, val1, val2)
End Function
'////GET MINIMUM OF 2
Public Function Min( _
       ByVal val1 As Long, _
       ByVal val2 As Long _
       ) As Long

    Min = IIf(val1 < val2, val1, val2)
End Function
Public Function FileExists(ByVal sFilename$) As Boolean
'checks if a file or dir exists
'returns true if it does, returns false otherwise
    Dim hFile&, Win32FindData As WIN32_FIND_DATA
    sFilename = Trim$(sFilename)
    hFile = FindFirstFile(sFilename, Win32FindData)

    If (hFile <> INVALID_HANDLE_VALUE) And (hFile <> ERROR_NO_MORE_FILES) Then
        FileExists = True
    ElseIf GetFileAttributes(sFilename) <> (-1) Then
        ' FindFirstFile will not return the root dor of a drive so we check the attributes
        ' of sFileName in case it is the root
        FileExists = True
    End If

    Call FindClose(hFile)

End Function

Public Function FolderExists(ByVal sFilename$) As Boolean
'checks if a dir exists
'returns true if it does, returns false otherwise
    sFilename = Trim$(sFilename)
    FolderExists = Fsys.FolderExists(sFilename)
End Function

Public Function GetFromINI(Section As String, Key As String, Directory As String, Default As String) As String
    Dim strBuffer As String
    strBuffer = String(750, Chr(0))
    Key$ = LCase$(Key)

    GetFromINI = Left(strBuffer, GetPrivateProfileString(Section, ByVal Key, "", strBuffer, Len(strBuffer), Directory))

    If IsBlank(GetFromINI) Then GetFromINI = Default
End Function

'////CHECK IF VARIANT HAS LENGTH OR NOT
Public Function IsBlank(Var As Variant) As Boolean
    If Len(Trim(Var)) = 0 Then
        IsBlank = True
    Else
        IsBlank = False
    End If
End Function

'////CONVERT TIME FROM SECOND TO HH:MM:SS FORMAT
Function Convert_Time_to_string(ByVal LSec As Long) As String
    Dim HH As Long, MM As Long, ss As Long
    Dim Tmp As String

    HH = LSec \ 3600  '// calCULATE HOURS
    MM = LSec \ 60 Mod 60    '//calCULATE HMINUTES
    ss = LSec Mod 60  '// calCULATE SECONDS

    If HH > 0 Then Tmp = HH & ":"
    Convert_Time_to_string = Tmp & Format$(MM, "00") & ":" & Format$(ss, "00")
End Function

'////CONVERT TIME FROM String "xx:xx" TO seconds
Public Function Convert_TextTime_to_Seconds(sTempTime As String) As Long
    Dim ArrTime() As String
    ArrTime = Split(sTempTime, ":", , vbTextCompare)

    Convert_TextTime_to_Seconds = Val(Right(sTempTime, 2)) + 60 * Val(Left(sTempTime, 2))    '+ 3600 * CLng(Left(sTemptime, Len(sTemptime) - 5))
End Function

Function NZ(Value, Optional ValueIfNull)
    If (IsMissing(ValueIfNull)) Then
        NZ = IIf(IsNull(Value), vbNullString, Value)
    Else
        NZ = IIf(IsNull(Value), ValueIfNull, Value)
    End If
End Function



Public Function UpperCase_Firstletter(sText As String) As String
'use: It replces first letter of each word by capital letter
'     eg. mahesh kumar kurmi >>> Mahesh Kumar Kurmi
    Dim sCleanStr As String, sNewString As String
    Dim sSplitField() As String
    Dim iSpaces As Integer
    sCleanStr = Trim$(sText)

    'Upper case and / or lower case the string correctly.
    sSplitField = Split(sCleanStr, " ", , vbTextCompare)
    sCleanStr = ""
    ''Debug.Print sCleanStr
    For iSpaces = 0 To UBound(sSplitField)
        If (Not iSpaces = 0 Or Not IsNumeric(sSplitField(iSpaces))) And sSplitField(iSpaces) <> "" Then
            sNewString = UCase$(Left$(sSplitField(iSpaces), 1))
            sNewString = sNewString & (Right$(sSplitField(iSpaces), Len(sSplitField(iSpaces)) - 1))
            sCleanStr = sCleanStr & sNewString & " "
        End If
    Next iSpaces
    UpperCase_Firstletter = Trim$(sCleanStr)
End Function
