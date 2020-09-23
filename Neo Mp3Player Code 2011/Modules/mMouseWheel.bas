Attribute VB_Name = "mMSubclass"
Option Explicit
'Subclassing main form for DragDrop and activate event
Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDest As Any, pSrc As Any, ByVal cbSrc As Long)

Private Const WM_DROPFILES = &H233
Private Const MF_INSERT = &H0&
Private Const MF_CHANGE = &H80&
Private Const MF_APPEND = &H100&
Private Const MF_DELETE = &H200&
Public Const MF_REMOVE = &H1000&
Private Const WM_SYSCOMMAND = &H112
Private Const WM_ACTIVATE As Long = &H6    ' to check if application/form activated or deactivated
Private Const WM_ACTIVATEAPP As Long = &H1C    'doesn't seem to respond with activation or deactivation of form
' it occurs only on minimizing so i used &H6
Private Const WHEEL_DELTA = 120


Private Const GWL_WNDPROC = (-4)

Private lpPrevWndProc As Long
Private Type COPYDATASTRUCT
    dwData As Long
    cbData As Long
    lpData As Long
End Type
Private Const WM_CLOSE = &H10
Private Const WM_COPYDATA = &H4A

Private nCopyData As COPYDATASTRUCT
Private nBuffer() As Byte
Private nOldProc As Long


Private mlPrevWndProc As Long
Private mhHookWindow As Long
Private moDropForm As Object

Public Declare Sub DragAcceptFiles Lib "Shell32.dll" _
                                   (ByVal hwnd As Long, _
                                    ByVal fAccept As Long)

Private Declare Sub DragFinish Lib "Shell32.dll" _
                               (ByVal hDrop As Long)

Private Declare Function DragQueryFile Lib "Shell32.dll" _
                                       Alias "DragQueryFileA" _
                                       (ByVal hDrop As Long, _
                                        ByVal UINT As Long, _
                                        ByVal lpStr As String, _
                                        ByVal ch As Long) As Long

'=======================================================
'conversion
'=======================================================
Public Function LoWord(DWord As Long) As Long
    If DWord And &H8000& Then    ' &H8000& = &H00008000
        LoWord = DWord Or &HFFFF0000
    Else
        LoWord = DWord And &HFFFF&
    End If
End Function

Public Function HiWord(DWord As Long) As Long
    HiWord = (DWord And &HFFFF0000) \ &H10000
End Function


Public Function WindowProc(ByVal hwnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    Dim fwKeys As Integer, zDelta As Integer
    On Error GoTo errorHandler
    ' Call CallWindowProc(lpPrevWndProc, hWnd, uMsg, wParam, lParam)
    Select Case uMsg
    Case WM_DROPFILES
        ' Set moDropForm = Main
        ' Get the list of dropped files from the OS
        If hwnd = frmMain.hwnd Then
            Call GetDropFileList(wParam, lParam, True)    'newlist as true
        ElseIf hwnd = frmPLST.hwnd Then
            Call GetDropFileList(wParam, lParam, False)    'newlist as false
        End If
    Case WM_MBUTTONDOWN
        '/* The Wheel button is down
    Case WM_MBUTTONUP
        '/* The Wheel button isn't down anymore
    Case &H214
        If hwnd = frmPLST.hwnd And Not bLoading Then frmPLST.Form_Resize1
    Case WM_MBUTTONDBLCLK
        '/* The Wheel button has been double-clicked
    Case WM_MOUSEWHEEL
        fwKeys = LoWord(wParam)
        zDelta = HiWord(wParam) / WHEEL_DELTA
        '/* Wheel rotate
        '/* abs(zDelta) ---> Ticks,Points
        '/* zDelta > 0  ---> Rotate forward
        '/* zDelta < 0  ---> Rotate backward
        If zDelta > 0 Then    ' forward
            If hwnd = frmPLST.hwnd Then
                If frmPLST.Scrollbar.Value >= 1 Then frmPLST.Scrollbar.Value = frmPLST.Scrollbar.Value - 1
            ElseIf hwnd = frmMain.hwnd Then
                frmMain.Mouse_Wheel_Moving (False)
                'frmMain.Form_KeyPress 43 'Mas volumen
                frmMain.Timer_Wait.Enabled = True
            End If

        Else    '/* backward

            If hwnd = frmPLST.hwnd Then
                If frmPLST.Scrollbar.Value < frmPLST.Scrollbar.Max Then frmPLST.Scrollbar.Value = frmPLST.Scrollbar.Value + 1  '(False)
            ElseIf hwnd = frmMain.hwnd Then
                frmMain.Mouse_Wheel_Moving (True)
                frmMain.Timer_Wait.Enabled = True
            End If
        End If

    Case &H4A    'WM_COPYDATA
        If hwnd = frmMain.hwnd Then
            Dim sCommand As String
            Call CopyMemory(nCopyData, ByVal lParam, Len(nCopyData))
            ReDim nBuffer(1 To nCopyData.cbData)
            Call CopyMemory(nBuffer(1), ByVal nCopyData.lpData, nCopyData.cbData)
            sCommand = StrConv(nBuffer, vbUnicode)
            sCommand = StripNulls(sCommand)

            'PROCESS FILE FROM ALL OTHER INSTANCES
            ProcessCommandParameter (sCommand)

        End If

    Case &H6    'WM_ACTIVATEAPP
        If hwnd = frmMain.hwnd Then
            If wParam = 0 Then
                frmMain.DeactivateMe
            Else
                frmMain.ActivateMe
            End If
        ElseIf hwnd = frmPLST.hwnd Then
            If wParam = 0 Then
                frmPLST.DeactivateMe
            Else
                frmPLST.ActivateMe
            End If
        End If

    Case WM_SYSCOMMAND
        Call Process_SystemMenu(wParam)
    End Select

    WindowProc = CallWindowProc(lpPrevWndProc, hwnd, uMsg, wParam, lParam)
    Exit Function
errorHandler:
    WindowProc = CallWindowProc(lpPrevWndProc, hwnd, uMsg, wParam, lParam)
    MsgBox err.Description & "Module: windowproc"
End Function
Public Sub SubclassEnqueue(ByVal hwnd As Long)
    nOldProc = GetWindowLong(hwnd, GWL_WNDPROC)
    Call SetWindowLong(hwnd, GWL_WNDPROC, AddressOf WindowProc)
End Sub

'UNSUBCLASSES THE FORM
Public Sub UnSubclassEnqueue(ByVal hwnd As Long)
    Call SetWindowLong(hwnd, GWL_WNDPROC, nOldProc)
End Sub

Public Sub Hook(Handle As Long)
    On Error Resume Next
    lpPrevWndProc = SetWindowLong(Handle, GWL_WNDPROC, AddressOf WindowProc)
End Sub

Public Sub Unhook(Handle As Long)
    On Error Resume Next
    Call SetWindowLong(Handle, GWL_WNDPROC, lpPrevWndProc)
End Sub

'===========================================================================================================
'SUBCLASSES THE FORM
'===========================================================================================================
'
Public Sub EnableFileDrops(oDropTarget As Form)
'Be sure we are not arleady subclassing drops
'If mhHookWindow <> 0 Then Call DisableFileDrops

'Save the handle and object reference
' of the calling window
    Set moDropForm = oDropTarget
    mhHookWindow = moDropForm.hwnd

    ' Set the subclassing window message hook
    ' telling system to pass message through the subroutine "windowproc" of main form which
    ' handles the message
    ' Tell the OS that the specified window accepts
    ' dropped files
    Call DragAcceptFiles(mhHookWindow, True)
End Sub

Public Sub DisableFileDrops()

    Dim lReturn As Long
    ' Check to be sure that there is a hook active
    If mhHookWindow = 0 Then Exit Sub
    If IsEmpty(mhHookWindow) = True Then Exit Sub
    If IsNull(mhHookWindow) = True Then Exit Sub

    ' Tell the OS that the specified window no longer
    ' accepts dropped files
    Call DragAcceptFiles(mhHookWindow, False)
    Call DragAcceptFiles(frmPLST.hwnd, False)
    Call DragAcceptFiles(frmMain.hwnd, False)

    'Clear the window handles and references
    mhHookWindow = 0
    mlPrevWndProc = 0
    Set moDropForm = Nothing

End Sub

Public Sub GetDropFileList(wParam As Long, _
                           lParam As Long, newList As Boolean)

    Dim nDropCount As Integer
    Dim nLoopCtr As Integer
    Dim lReturn As Long
    Dim hDrop As Long
    Dim sFilename As String
    Dim vFileNames As Variant

    ' Save the drop structure handle
    hDrop = wParam

    ' Allocate space for the return value
    sFilename = Space$(255)

    ' Get the number of file names dropped
    nDropCount = DragQueryFile(hDrop, -1, sFilename, 254)

    ' Allocate variant array elements to store the
    ' dropped file names
    vFileNames = Array(" ")
    ReDim vFileNames(nDropCount - 1) As String

    ' Loop to get each dropped file name and
    ' add it to the variant array
    For nLoopCtr = 0 To nDropCount - 1
        ' Allocate space for the return value
        sFilename = Space$(255)
        ' Get a dropped file name
        lReturn = DragQueryFile(hDrop, nLoopCtr, sFilename, 254)
        vFileNames(nLoopCtr) = Left$(sFilename, lReturn)
    Next nLoopCtr

    ' Release the drop structure from memory
    Call DragFinish(hDrop)

    ' Call the form method to pass the list of dropped files
    Call frmPLST.DroppedFiles(vFileNames, newList)
    'frmplst.o

End Sub




