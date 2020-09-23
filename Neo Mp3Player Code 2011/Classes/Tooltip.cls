VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Tooltip"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit


Private Declare Sub InitCommonControls Lib "COMCTL32.DLL" ()

''Windows API Functions
Private Declare Function CreateWindowEx Lib "user32" Alias "CreateWindowExA" (ByVal dwExStyle As Long, ByVal lpClassName As String, ByVal lpWindowName As String, ByVal dwStyle As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hWndParent As Long, ByVal hMenu As Long, ByVal hInstance As Long, lpParam As Any) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function DestroyWindow Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Declare Function GetClientRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Private Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
Private Declare Function GetDCEx Lib "user32" (ByVal hwnd As Long, ByVal hrgnclip As Long, ByVal fdwOptions As Long) As Long

''Windows API Constants
Private Const WM_USER = &H400
Private Const CW_USEDEFAULT = &H80000000
Private Const SWP_NOSIZE = &H1
Private Const SWP_NOACTIVATE = &H10
Private Const SWP_NOMOVE = &H2
Private Const HWND_TOPMOST = -1

''Windows API Types
Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

''Tooltip Window Constants
Private Const TTS_NOPREFIX = &H2
Private Const TTF_TRANSPARENT = &H100
Private Const TTF_CENTERTIP = &H2
Private Const TTM_ADDTOOLA = (WM_USER + 4)
Private Const TTM_ACTIVATE = WM_USER + 1
Private Const TTM_UPDATETIPTEXTA = (WM_USER + 12)
Private Const TTM_SETMAXTIPWIDTH = (WM_USER + 24)
Private Const TTM_SETTIPBKCOLOR = (WM_USER + 19)
Private Const TTM_SETTIPTEXTCOLOR = (WM_USER + 20)
Private Const TTM_SETTITLE = (WM_USER + 32)
Private Const TTS_BALLOON = &H40
Private Const TTS_ALWAYSTIP = &H1
Private Const TTF_SUBCLASS = &H10
Private Const TOOLTIPS_CLASSA = "tooltips_class32"

''Tooltip Window Types
Private Type TOOLINFO
    lSize As Long
    lFlags As Long
    lhWnd As Long
    lId As Long
    lpRect As RECT
    hInstance As Long
    lpStr As String
    lParam As Long
End Type


'local variable(s) to hold property value(s)
Private mvarBackColor As Long    'local copy
Private mvarTitle As String    'local copy
Private mvarForeColor As Long    'local copy
Private mvarParentControl As Object    'local copy
Private mvarIcon As ttIconType    'local copy
Private mvarCentered As Boolean    'local copy
Private mvarStyle As ttStyleEnum    'local copy

Public Enum ttIconType
    TTNoIcon = 0
    TTIconInfo = 1
    TTIconWarning = 2
    TTIconError = 3
End Enum

Public Enum ttStyleEnum
    TTStandard
    TTBalloon
End Enum

'private data
Private lhwndParentControl As Long
Private lhWnd As Long
Private TI As TOOLINFO
Public Property Let Style(ByVal vData As ttStyleEnum)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Style = 5
    mvarStyle = vData
End Property

Public Property Get Style() As ttStyleEnum
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: 'Debug.Print X.Style
    Style = mvarStyle
End Property

Public Property Let Centered(ByVal vData As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Centered = 5
    mvarCentered = vData
End Property

Public Property Get Centered() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: 'Debug.Print X.Centered
    Centered = mvarCentered
End Property

Public Function Create() As Boolean
    Dim lpRect As RECT
    Dim lWinStyle As Long

    If lhWnd <> 0 Then
        DestroyWindow lhWnd
    End If

    lWinStyle = TTS_ALWAYSTIP Or TTS_NOPREFIX

    ''create baloon style if desired
    If mvarStyle = TTBalloon Then lWinStyle = lWinStyle Or TTS_BALLOON

    ''the parent control has to have been set first
    If lhwndParentControl <> 0 Then
        lhWnd = CreateWindowEx(0&, TOOLTIPS_CLASSA, vbNullString, lWinStyle, CW_USEDEFAULT, _
                               CW_USEDEFAULT, CW_USEDEFAULT, CW_USEDEFAULT, lhwndParentControl, _
                               0&, App.hInstance, 0&)
        ''make our tooltip window a topmost window
        SetWindowPos lhWnd, HWND_TOPMOST, 0&, 0&, 0&, 0&, SWP_NOACTIVATE Or SWP_NOSIZE Or SWP_NOMOVE

        ''get the rect of the parent control
        GetClientRect lhwndParentControl, lpRect

        ''now set our tooltip info structure
        With TI
            ''if we want it centered, then set that flag
            If mvarCentered Then
                .lFlags = TTF_SUBCLASS Or TTF_CENTERTIP
            Else
                .lFlags = TTF_SUBCLASS
            End If

            'set the hwnd prop to our parent control's hwnd
            .lhWnd = lhwndParentControl
            .lId = 0
            .hInstance = App.hInstance
            '.lpstr = ALREADY SET
            .lpRect = lpRect
        End With

        ''add the tooltip structure
        SendMessage lhWnd, TTM_ADDTOOLA, 0&, TI

        ''if we want a title or we want an icon
        If mvarTitle <> vbNullString Or mvarIcon <> TTNoIcon Then
            SendMessage lhWnd, TTM_SETTITLE, CLng(mvarIcon), ByVal mvarTitle
        End If

        If mvarForeColor <> Empty Then
            SendMessage lhWnd, TTM_SETTIPTEXTCOLOR, mvarForeColor, 0&
        End If

        If mvarBackColor <> Empty Then
            SendMessage lhWnd, TTM_SETTIPBKCOLOR, mvarBackColor, 0&
        End If
        ShowWindow lhWnd, 4

        Dim k As Long
        ' k = GetDC()
        ' k = GetDCEx(frmMain.picNormalMode.hwnd, 0, &H1& Or &H400& Or &H2&)
        k = GetDCEx(lhWnd, 0, &H1& Or &H400& Or &H2&)
        'Debug.Print "Tooltip" & k
        'pDraw GetDC(lhWnd), 300, 300, 120, 120, 0, 0   '//-- Draw The Temp DC into the Real Window DC.
        frmMain.cImage.Render k, 0, 0, 100, 100
    End If
End Function

'DRAWS THE TEMP DC
Private Sub pDraw( _
        ByVal hDC As Long, _
        Optional ByVal xSrc As Long = 0, Optional ByVal ySrc As Long = 0, _
        Optional ByVal WidthSrc As Long = 0, Optional ByVal HeightSrc As Long = 0, _
        Optional ByVal xDst As Long = 0, Optional ByVal yDst As Long = 0 _
                                         )
'If WidthSrc <= 0 Then WidthSrc = M_ExWindow.Width
'If HeightSrc <= 0 Then HeightSrc = M_ExWindow.Height


'BitBlt hDC, xDst, yDst, 120, 120, frmMain.picNormalMode.DC, 0, 0, vbSrcCopy

End Sub
'======================================================================

'Public Property Let ParentControl(ByVal hwnd As Long)
'used when assigning an Object to the property, on the left side of a Set statement.
'Syntax: Set x.ParentControl = frmVisualisation
' Set mvarParentControl = vData
'End Property

Public Property Let ParentControl(ByVal hwnd As Long)
'used when assigning an Object to the property, on the left side of a Set statement.
'Syntax: Set x.ParentControl = frmVisualisation
    lhwndParentControl = hwnd
End Property

Public Property Get hwndParentControl() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: 'Debug.Print X.ParentControl
    hwndParentControl = lhwndParentControl
End Property

Public Property Let Icon(ByVal vData As ttIconType)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Icon = 5
    mvarIcon = vData
    If lhWnd <> 0 And mvarTitle <> Empty And mvarIcon <> TTNoIcon Then
        SendMessage lhWnd, TTM_SETTITLE, CLng(mvarIcon), ByVal mvarTitle
    End If

End Property


Public Property Get Icon() As ttIconType
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: 'Debug.Print X.Icon
    Icon = mvarIcon
End Property
Public Property Let ForeColor(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.ForeColor = 5
    mvarForeColor = vData
    If lhWnd <> 0 Then
        SendMessage lhWnd, TTM_SETTIPTEXTCOLOR, mvarForeColor, 0&
    End If

End Property

Public Property Get ForeColor() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: 'Debug.Print X.ForeColor
    ForeColor = mvarForeColor
End Property

Public Property Let Title(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Title = 5
    mvarTitle = vData
    If lhWnd <> 0 And mvarTitle <> Empty And mvarIcon <> TTNoIcon Then
        SendMessage lhWnd, TTM_SETTITLE, CLng(mvarIcon), ByVal mvarTitle
    End If
End Property

Public Property Get Title() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: 'Debug.Print X.Title
    Title = TI.lpStr
End Property

Public Property Let BackColor(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.BackColor = 5
    mvarBackColor = vData
    If lhWnd <> 0 Then
        SendMessage lhWnd, TTM_SETTIPBKCOLOR, vData, 0
    End If

End Property


Public Property Get BackColor() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: 'Debug.Print X.BackColor
    BackColor = mvarBackColor
End Property

Public Property Let TipText(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.TipText = 5
    TI.lpStr = vData
    If lhWnd <> 0 Then
        SendMessage lhWnd, TTM_UPDATETIPTEXTA, 0&, TI
    End If

End Property

Public Property Get TipText() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: 'Debug.Print X.TipText
    TipText = TI.lpStr
End Property

Private Sub Class_Terminate()
    If lhWnd <> 0 Then
        DestroyWindow lhWnd
    End If

End Sub


