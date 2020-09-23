Attribute VB_Name = "modRegisterFileType"
'===========================================================================================================
'ASSOCIATE ICON WITH FILE - MODULE CODE
'===========================================================================================================

Option Explicit
Public Declare Function GetLongPathName Lib "kernel32" Alias "GetLongPathNameA" (ByVal lpszShortPath As String, ByVal lpszLongPath As String, ByVal cchBuffer As Long) As Long


Public Declare Function GetShortPathName Lib "kernel32" Alias "GetShortPathNameA" (ByVal lpszLongPath As String, ByVal lpszShortPath As String, ByVal cchBuffer As Long) As Long
Public LstTextHeight As Long
Public LstTextWidth As Long

Public ListRange As Long                            'Range of list that will be displayed


'===========================================================================================================
'START VARIABLES TO ENQUEUE FILES FROM EXPLORER
'===========================================================================================================
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (hpvDest As Any, hpvSource As Any, ByVal cbCopy As Long)

Private Type COPYDATASTRUCT
    dwData As Long
    cbData As Long
    lpData As Long
End Type

Private Const GWL_WNDPROC = -4

Private Const WM_CLOSE = &H10
Private Const WM_COPYDATA = &H4A

Private nCopyData As COPYDATASTRUCT
Private nBuffer() As Byte
Private nOldProc As Long
'===========================================================================================================
'END VARIABLES TO ENQUEUE FILES FROM EXPLORER
'===========================================================================================================

Private Declare Sub SHChangeNotify Lib "shell32" (ByVal wEventId As Long, ByVal uFlags As Long, ByVal dwItem1 As Long, ByVal dwItem2 As Long)

Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Private Declare Function RegCreateKey Lib "advapi32.dll" Alias "RegCreateKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Private Declare Function RegDeleteKey Lib "advapi32.dll" Alias "RegDeleteKeyA" (ByVal hKey As Long, ByVal lpSubKey As String) As Long
Private Declare Function RegDeleteValue Lib "advapi32.dll" Alias "RegDeleteValueA" (ByVal hKey As Long, ByVal lpValueName As String) As Long
Private Declare Function RegOpenKey Lib "advapi32.dll" Alias "RegOpenKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Private Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long
Private Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpData As Any, ByVal cbData As Long) As Long

Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long

Private Const REG_SZ = 1

Private Const ERROR_SUCCESS = 0&

Private Const HKEY_CLASSES_ROOT = &H80000000

Private Const SHCNF_IDLIST = &H0
Private Const SHCNE_ASSOCCHANGED = &H8000000

'ASSOCIATES A FILETYPE WITH THE PROGRAM. IT USES THE COMMAND, " %1", TO LOAD FILE IN YOUR PROGRAM
Public Sub RegisterType(ByVal sExt As String, ByVal sName As String, ByVal sConType As String, ByVal sDescription As String, ByVal iResIcon As Integer)

    If Left(sExt, 1) <> "." Then sExt = "." & sExt

    Call DeleteKey(HKEY_CLASSES_ROOT, sName & "\Shell")
    Call DeleteKey(HKEY_CLASSES_ROOT, sName)

    Call CreateKey(HKEY_CLASSES_ROOT, sExt, "", sName)
    Call CreateKey(HKEY_CLASSES_ROOT, sExt, "Content Type", sConType)
    Call CreateKey(HKEY_CLASSES_ROOT, sName, "", sDescription)
    Call CreateKey(HKEY_CLASSES_ROOT, sName & "\DefaultIcon", "", App.Path & "\" & App.EXEName & ".exe," & iResIcon)
    Call CreateKey(HKEY_CLASSES_ROOT, sName & "\Shell", "", "")
    Call CreateKey(HKEY_CLASSES_ROOT, sName & "\Shell\Open", "", "")
    Call CreateKey(HKEY_CLASSES_ROOT, sName & "\Shell\Play with NeoMP3", "", "")
    Call CreateKey(HKEY_CLASSES_ROOT, sName & "\Shell\Enqueue in NeoMP3", "", "")

    Call CreateKey(HKEY_CLASSES_ROOT, sName & "\Shell\Open\Command", "", App.Path & "\" & App.EXEName & ".exe %1")
    If UCase(sExt) = ".M3U" Or UCase(sExt) = ".NPL" Or UCase(sExt) = ".PLS" Then
        Call CreateKey(HKEY_CLASSES_ROOT, sName & "\Shell\Enqueue in NeoMP3\Command", "", App.Path & "\" & App.EXEName & ".exe ADDP%1")
        Call CreateKey(HKEY_CLASSES_ROOT, sName & "\Shell\Play with NeoMP3\Command", "", App.Path & "\" & App.EXEName & ".exe RUNP%1")
    Else
        Call CreateKey(HKEY_CLASSES_ROOT, sName & "\Shell\Enqueue in NeoMP3\Command", "", App.Path & "\" & App.EXEName & ".exe ADDF%1")
        Call CreateKey(HKEY_CLASSES_ROOT, sName & "\Shell\Play with NeoMP3\Command", "", App.Path & "\" & App.EXEName & ".exe PLAY%1")
    End If
    SHChangeNotify SHCNE_ASSOCCHANGED, SHCNF_IDLIST, 0, 0
End Sub

'UNASSOCIATES A FILETYPE WITH PROGRAM. IT SIMPLY DELETES THE FILENAME KEY.
Public Sub DeleteType(ByVal sExt As String, ByVal sName As String)
    If Left(sExt, 1) <> "." Then sExt = "." & sExt
    If FileAssociated(sExt, sName) = False Then Exit Sub
    Call CreateKey(HKEY_CLASSES_ROOT, sExt, "", "")
    Call CreateKey(HKEY_CLASSES_ROOT, sExt, "Content Type", "")
    Call DeleteKey(HKEY_CLASSES_ROOT, sName & "\DefaultIcon")
    Call DeleteKey(HKEY_CLASSES_ROOT, sName & "\Shell\Open\Command")
    Call DeleteKey(HKEY_CLASSES_ROOT, sName & "\Shell\Open")
    Call DeleteKey(HKEY_CLASSES_ROOT, sName & "\Shell")
    Call DeleteKey(HKEY_CLASSES_ROOT, sName)

    SHChangeNotify SHCNE_ASSOCCHANGED, SHCNF_IDLIST, 0, 0
End Sub

'CHECKS TO SEE IF YOUR PROGRAM IS THE HANDLER FOR A CERTAIN FILETYPE
Public Function FileAssociated(ByVal sExt As String, ByVal sName As String) As Boolean
    If Left(sExt, 1) <> "." Then sExt = "." & sExt

    FileAssociated = CBool(GetString(HKEY_CLASSES_ROOT, sExt, "") = sName)
End Function

'HELPER FUNCTION THAT CHECKS TO SEE IF A FILE IS ASSOCIATED WITH YOUR PROGRAM
Private Function GetString(ByVal hKey As Long, ByVal sPath As String, ByVal sValue As String)
    Dim lResult As Long
    Dim lHandle As Long
    Dim sBuffer As String
    Dim lLenBuffer As Long
    Dim lValueType As Long
    Dim iZeroPos As Integer

    Call RegOpenKey(hKey, sPath, lHandle)

    lResult = RegQueryValueEx(lHandle, sValue, 0&, lValueType, ByVal 0&, lLenBuffer)

    If lValueType = REG_SZ Then
        sBuffer = String(lLenBuffer, " ")
        lResult = RegQueryValueEx(lHandle, sValue, 0&, 0&, ByVal sBuffer, lLenBuffer)
        If lResult = ERROR_SUCCESS Then
            iZeroPos = InStr(sBuffer, Chr$(0))
            If iZeroPos > 0 Then
                GetString = Left$(sBuffer, iZeroPos - 1)
            Else
                GetString = sBuffer
            End If
        End If
    End If
End Function

'HELPER PROCEDURE THAT CREATES ALL STRING VALUES IN THE REGISTRY
Private Sub CreateKey(ByVal hKey As Long, ByVal sPath As String, ByVal sValue As String, ByVal sData As String)
    Dim lResult As Long

    Call RegCreateKey(hKey, sPath, lResult)
    Call RegSetValueEx(lResult, sValue, 0, REG_SZ, ByVal sData, Len(sData))
    Call RegCloseKey(lResult)
End Sub

'HELPER PROCEDIRE THAT DELETES A STRING IN A REGISTRY KEY
Private Sub DeleteKey(ByVal hKey As Long, ByVal sKey As String)
    Call RegDeleteKey(hKey, sKey)
End Sub

'===========================================================================================================
'START PROCEDURES TO ENQUEUE FILES FROM EXPLORER
'===========================================================================================================
'SUBCLASSES THE FORM
Public Sub SubclabssEnqueue(ByVal hwnd As Long)
    nOldProc = GetWindowLong(hwnd, GWL_WNDPROC)
    ' Call SetWindowLong(hWnd, GWL_WNDPROC, AddressOf EnqueueProcedure)
End Sub

'UNSUBCLASSES THE FORM
Public Sub UnSubclabssEnqueue(ByVal hwnd As Long)
    Call SetWindowLong(hwnd, GWL_WNDPROC, nOldProc)
End Sub


'FIRES WHEN ANOTHER FILE IS OPENED WHEN AN INSTANCE IS ALREADY AVAILABLE
Private Function EnhqueueProcedure(ByVal hwnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    Select Case uMsg
    Case WM_COPYDATA
        Dim sCommand As String

        Call CopyMemory(nCopyData, ByVal lParam, Len(nCopyData))
        Call CopyMemory(nBuffer(1), ByVal nCopyData.lpData, nCopyData.cbData)
        sCommand = StrConv(nBuffer, vbUnicode)

        'PROCESS FILE FROM ALL OTHER INSTANCES
        'MsgBox sCommand
        frmMain.Caption = sCommand
        ' EnqueueProcedure = 0
    Case WM_CLOSE
        ' Call UnSubclassEnqueue(hWnd)
    Case Else
        ' EnqueueProcedure = CallWindowProc(nOldProc, hWnd, uMsg, wParam, lParam)
    End Select
End Function
'===========================================================================================================
'END PROCEDURES TO ENQUEUE FILES FROM EXPLORER
'===========================================================================================================

'===========================================================================================================
'HOW TO USE THIS CODE!
'===========================================================================================================
'
'Private Sub Form_Load()
'    Call EnqueueProcess(hWnd, Command$)
'
'    If Not App.PrevInstance And Command$ <> "%1" And Command$ <> "" Then
'        'ONLY USED IF THIS WAS THE FIRST INSTANCE - ALL OTHER INSTANCES GO THROUGH SUBCLASSING
'        MsgBox Command$
'    End If
'
'    Call SubclassEnqueue(hWnd)
'End Sub
'
'Private Sub cmdRegister_Click()
'    Call RegisterType(".mp3", "XamP.File", "Audio/MPEG", "XamP Media File", 0)
'End Sub
'
'Private Sub cmdUnRegister_Click()
'    Call DeleteType(".mp3", "XamP.File")
'End Sub
'
'Private Sub cmdCheckAssociation_Click()
'    MsgBox FileAssociated(".mp3", "XamP.File"), vbInformation, "Associated?"
'End Sub
'
'===========================================================================================================
'CODE MODIFIED BY MICHAEL DOMBROWSKI
'===========================================================================================================
'===========================================================================================================
'ASSOCIATE ICON WITH FILE - MODULE CODE
'===========================================================================================================

'===========================================================================================================
'START VARIABLES TO ENQUEUE FILES FROM EXPLORER
'===========================================================================================================

'===========================================================================================================
'END VARIABLES TO ENQUEUE FILES FROM EXPLORER
'===========================================================================================================


'ASSOCIATES A FILETYPE WITH YOUR PROGRAM. IT USES THE COMMAND, " %1", TO LOAD FILE IN YOUR PROGRAM


'===========================================================================================================
'=========================================================================================================

Public Function AddBackSlash(ByVal sPath As String) As String
'Returns sPath with a trailing backslash
'     if sPath does not
'already have a trailing backslash. Othe
'     rwise, returns sPath.
    sPath = Trim$(sPath)

    If Len(sPath) > 0 Then
        sPath = sPath & IIf(Right$(sPath, 1) <> "\", "\", "")
    End If
    AddBackSlash = sPath
End Function

Public Function GetLongFilename(ByVal sShortFilename As String) As String
'Returns the Long Filename associated wi
'     th sShortFilename
    Dim lRet As Long
    Dim sLongFilename As String
    'First attempt using 1024 character buff
    '     er.
    sLongFilename = String$(1024, " ")
    lRet = GetLongPathName(sShortFilename, sLongFilename, Len(sLongFilename))

    'If buffer is too small lRet contains bu
    '     ffer size needed.


    If lRet > Len(sLongFilename) Then
        'Increase buffer size...
        sLongFilename = String$(lRet + 1, " ")
        'and try again.
        lRet = GetLongPathName(sShortFilename, sLongFilename, Len(sLongFilename))
    End If

    'lRet contains the number of characters
    '     returned.


    'If lRet > 0 Then
    GetLongFilename = Left$(sLongFilename, lRet)
    'End If

End Function


Public Function GetShortFilename(ByVal sLongFilename As String) As String
'Returns the Short Filename associated w
'     ith sLongFilename
    Dim lRet As Long
    Dim sShortFilename As String
    'First attempt using 1024 character buff
    '     er.
    sShortFilename = String$(1024, " ")
    lRet = GetShortPathName(sLongFilename, sShortFilename, Len(sShortFilename))

    'If buffer is too small lRet contains bu
    '     ffer size needed.


    If lRet > Len(sShortFilename) Then
        'Increase buffer size...
        sShortFilename = String$(lRet + 1, " ")
        'and try again.
        lRet = GetShortPathName(sLongFilename, sShortFilename, Len(sShortFilename))
    End If

    'lRet contains the number of characters
    '     returned.
    If lRet > 0 Then
        GetShortFilename = Left$(sShortFilename, lRet)
    End If

End Function


Public Function RemoveBackSlash(ByVal sPath As String) As String
'Returns sPath without a trailing backsl
'     ash if sPath
'has one. Otherwise, returns sPath.

    sPath = Trim$(sPath)


    If Len(sPath) > 0 Then
        sPath = Left$(sPath, Len(sPath) - IIf(Right$(sPath, 1) = "\", 1, 0))
    End If
    RemoveBackSlash = sPath

End Function


Public Function AppPath() As String
'Returns App.Path with backslash "\"
    Dim sPath As String
    sPath = App.Path
    AppPath = sPath & IIf(Right$(sPath, 1) <> "\", "\", "")

End Function


Public Function GetFilePath(ByVal sFilename As String, Optional ByVal bAddBackslash As Boolean) As String
'Returns Path Without FileTitle
    Dim lPos As Long
    lPos = InStrRev(sFilename, "\")


    If lPos > 0 Then
        GetFilePath = Left$(sFilename, lPos - 1) _
                      & IIf(bAddBackslash, "\", "")
    Else
        GetFilePath = ""
    End If

End Function


Public Function GetFileTitle(ByVal sFilename As String) As String
'Returns FileTitle Without Path
    Dim lPos As Long
    lPos = InStrRev(sFilename, "\")


    If lPos > 0 Then
        If lPos < Len(sFilename) Then
            GetFileTitle = Mid$(sFilename, lPos + 1)
        Else
            GetFileTitle = ""
        End If
    Else
        GetFileTitle = sFilename
    End If

End Function



