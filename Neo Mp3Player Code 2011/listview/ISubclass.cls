VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ISubclass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'*************************************************************************************************
'* iSubclass - cSubclass' implementation interface class + enumerations
'*
'* Paul_Caton@hotmail.com
'* Copyright free, use and abuse as you see fit.
'*
'* v1.0 The original..................................................................... 20060322
'*************************************************************************************************

Option Explicit

'-cSublass public declarations--------------------------------------------------------------------
Public Enum eMsgWhen                                                        'When to callback
    MSG_BEFORE = 1                                                            'Callback before the original WndProc
    MSG_AFTER = 2                                                             'Callback after the original WndProc
    MSG_BEFORE_AFTER = MSG_BEFORE Or MSG_AFTER                                'Callback before and after the original WndProc
End Enum

'Windows message values
Public Enum eMsg
    ALL_MESSAGES = -1
    eWM_NULL = &H0
    eWM_CREATE = &H1
    eWM_DESTROY = &H2
    eWM_MOVE = &H3
    eWM_SIZE = &H5
    eWM_ACTIVATE = &H6
    eWM_SETFOCUS = &H7
    eWM_KILLFOCUS = &H8
    eWM_ENABLE = &HA
    eWM_SETREDRAW = &HB
    eWM_SETTEXT = &HC
    eWM_GETTEXT = &HD
    eWM_GETTEXTLENGTH = &HE
    eWM_PAINT = &HF
    eWM_CLOSE = &H10
    eWM_QUERYENDSESSION = &H11
    eWM_QUIT = &H12
    eWM_QUERYOPEN = &H13
    eWM_ERASEBKGND = &H14
    eWM_SYSCOLORCHANGE = &H15
    eWM_ENDSESSION = &H16
    eWM_SHOWWINDOW = &H18
    eWM_WININICHANGE = &H1A
    eWM_SETTINGCHANGE = &H1A
    eWM_DEVMODECHANGE = &H1B
    eWM_ACTIVATEAPP = &H1C
    eWM_FONTCHANGE = &H1D
    eWM_TIMECHANGE = &H1E
    eWM_CANCELMODE = &H1F
    eWM_SETCURSOR = &H20
    eWM_MOUSEACTIVATE = &H21
    eWM_CHILDACTIVATE = &H22
    eWM_QUEUESYNC = &H23
    eWM_GETMINMAXINFO = &H24
    eWM_PAINTICON = &H26
    eWM_ICONERASEBKGND = &H27
    eWM_NEXTDLGCTL = &H28
    eWM_SPOOLERSTATUS = &H2A
    eWM_DRAWITEM = &H2B
    eWM_MEASUREITEM = &H2C
    eWM_DELETEITEM = &H2D
    eWM_VKEYTOITEM = &H2E
    eWM_CHARTOITEM = &H2F
    eWM_SETFONT = &H30
    eWM_GETFONT = &H31
    eWM_SETHOTKEY = &H32
    eWM_GETHOTKEY = &H33
    eWM_QUERYDRAGICON = &H37
    eWM_COMPAREITEM = &H39
    eWM_GETOBJECT = &H3D
    eWM_COMPACTING = &H41
    eWM_WINDOWPOSCHANGING = &H46
    eWM_WINDOWPOSCHANGED = &H47
    eWM_POWER = &H48
    eWM_COPYDATA = &H4A
    eWM_CANCELJOURNAL = &H4B
    eWM_NOTIFY = &H4E
    eWM_INPUTLANGCHANGEREQUEST = &H50
    eWM_INPUTLANGCHANGE = &H51
    eWM_TCARD = &H52
    eWM_HELP = &H53
    eWM_USERCHANGED = &H54
    eWM_NOTIFYFORMAT = &H55
    eWM_CONTEXTMENU = &H7B
    eWM_STYLECHANGING = &H7C
    eWM_STYLECHANGED = &H7D
    eWM_DISPLAYCHANGE = &H7E
    eWM_GETICON = &H7F
    eWM_SETICON = &H80
    eWM_NCCREATE = &H81
    eWM_NCDESTROY = &H82
    eWM_NCCALCSIZE = &H83
    eWM_NCHITTEST = &H84
    eWM_NCPAINT = &H85
    eWM_NCACTIVATE = &H86
    eWM_GETDLGCODE = &H87
    eWM_SYNCPAINT = &H88
    eWM_NCMOUSEMOVE = &HA0
    eWM_NCLBUTTONDOWN = &HA1
    eWM_NCLBUTTONUP = &HA2
    eWM_NCLBUTTONDBLCLK = &HA3
    eWM_NCRBUTTONDOWN = &HA4
    eWM_NCRBUTTONUP = &HA5
    eWM_NCRBUTTONDBLCLK = &HA6
    eWM_NCMBUTTONDOWN = &HA7
    eWM_NCMBUTTONUP = &HA8
    eWM_NCMBUTTONDBLCLK = &HA9
    eWM_KEYFIRST = &H100
    eWM_KEYDOWN = &H100
    eWM_KEYUP = &H101
    eWM_CHAR = &H102
    eWM_DEADCHAR = &H103
    eWM_SYSKEYDOWN = &H104
    eWM_SYSKEYUP = &H105
    eWM_SYSCHAR = &H106
    eWM_SYSDEADCHAR = &H107
    eWM_KEYLAST = &H108
    eWM_IME_STARTCOMPOSITION = &H10D
    eWM_IME_ENDCOMPOSITION = &H10E
    eWM_IME_COMPOSITION = &H10F
    eWM_IME_KEYLAST = &H10F
    eWM_INITDIALOG = &H110
    eWM_COMMAND = &H111
    eWM_SYSCOMMAND = &H112
    eWM_TIMER = &H113
    eWM_HSCROLL = &H114
    eWM_VSCROLL = &H115
    eWM_INITMENU = &H116
    eWM_INITMENUPOPUP = &H117
    eWM_MENUSELECT = &H11F
    eWM_MENUCHAR = &H120
    eWM_ENTERIDLE = &H121
    eWM_MENURBUTTONUP = &H122
    eWM_MENUDRAG = &H123
    eWM_MENUGETOBJECT = &H124
    eWM_UNINITMENUPOPUP = &H125
    eWM_MENUCOMMAND = &H126
    eWM_CTLCOLORMSGBOX = &H132
    eWM_CTLCOLOREDIT = &H133
    eWM_CTLCOLORLISTBOX = &H134
    eWM_CTLCOLORBTN = &H135
    eWM_CTLCOLORDLG = &H136
    eWM_CTLCOLORSCROLLBAR = &H137
    eWM_CTLCOLORSTATIC = &H138
    eWM_MOUSEFIRST = &H200
    eWM_MOUSEMOVE = &H200
    eWM_LBUTTONDOWN = &H201
    eWM_LBUTTONUP = &H202
    eWM_LBUTTONDBLCLK = &H203
    eWM_RBUTTONDOWN = &H204
    eWM_RBUTTONUP = &H205
    eWM_RBUTTONDBLCLK = &H206
    eWM_MBUTTONDOWN = &H207
    eWM_MBUTTONUP = &H208
    eWM_MBUTTONDBLCLK = &H209
    eWM_MOUSEWHEEL = &H20A
    eWM_PARENTNOTIFY = &H210
    eWM_ENTERMENULOOP = &H211
    eWM_EXITMENULOOP = &H212
    eWM_NEXTMENU = &H213
    eWM_SIZING = &H214
    eWM_CAPTURECHANGED = &H215
    eWM_MOVING = &H216
    eWM_DEVICECHANGE = &H219
    eWM_MDICREATE = &H220
    eWM_MDIDESTROY = &H221
    eWM_MDIACTIVATE = &H222
    eWM_MDIRESTORE = &H223
    eWM_MDINEXT = &H224
    eWM_MDIMAXIMIZE = &H225
    eWM_MDITILE = &H226
    eWM_MDICASCADE = &H227
    eWM_MDIICONARRANGE = &H228
    eWM_MDIGETACTIVE = &H229
    eWM_MDISETMENU = &H230
    eWM_ENTERSIZEMOVE = &H231
    eWM_EXITSIZEMOVE = &H232
    eWM_DROPFILES = &H233
    eWM_MDIREFRESHMENU = &H234
    eWM_IME_SETCONTEXT = &H281
    eWM_IME_NOTIFY = &H282
    eWM_IME_CONTROL = &H283
    eWM_IME_COMPOSITIONFULL = &H284
    eWM_IME_SELECT = &H285
    eWM_IME_CHAR = &H286
    eWM_IME_REQUEST = &H288
    eWM_IME_KEYDOWN = &H290
    eWM_IME_KEYUP = &H291
    eWM_MOUSEHOVER = &H2A1
    eWM_MOUSELEAVE = &H2A3
    eWM_CUT = &H300
    eWM_COPY = &H301
    eWM_PASTE = &H302
    eWM_CLEAR = &H303
    eWM_UNDO = &H304
    eWM_RENDERFORMAT = &H305
    eWM_RENDERALLFORMATS = &H306
    eWM_DESTROYCLIPBOARD = &H307
    eWM_DRAWCLIPBOARD = &H308
    eWM_PAINTCLIPBOARD = &H309
    eWM_VSCROLLCLIPBOARD = &H30A
    eWM_SIZECLIPBOARD = &H30B
    eWM_ASKCBFORMATNAME = &H30C
    eWM_CHANGECBCHAIN = &H30D
    eWM_HSCROLLCLIPBOARD = &H30E
    eWM_QUERYNEWPALETTE = &H30F
    eWM_PALETTEISCHANGING = &H310
    eWM_PALETTECHANGED = &H311
    eWM_HOTKEY = &H312
    eWM_PRINT = &H317
    eWM_PRINTCLIENT = &H318
    eWM_THEMECHANGED = &H31A
    eWM_HANDHELDFIRST = &H358
    eWM_HANDHELDLAST = &H35F
    eWM_AFXFIRST = &H360
    eWM_AFXLAST = &H37F
    eWM_PENWINFIRST = &H380
    eWM_PENWINLAST = &H38F
    eWM_USER = &H400
    eWM_APP = &H8000
End Enum

'-Subclass callback definition--------------------------------------------------------------------
Sub WndProc(ByVal bBefore As Boolean, _
            ByRef bHandled As Boolean, _
            ByRef lReturn As Long, _
            ByVal lng_hWnd As Long, _
            ByVal uMsg As eMsg, _
            ByVal wParam As Long, _
            ByVal lParam As Long, _
            ByRef lParamUser As Long)
'*************************************************************************************************
'* bBefore    - Indicates whether the callback is before or after the original WndProc. Usually
'*              you will know unless the callback for the uMsg value is specified as
'*              MSG_BEFORE_AFTER (both before and after the original WndProc).
'* bHandled   - In a before original WndProc callback, setting bHandled to True will prevent the
'*              message being passed to the original WndProc and (if set to do so) the after
'*              original WndProc callback.
'* lReturn    - WndProc return value. Set as per the MSDN documentation for the message value,
'*              and/or, in an after the original WndProc callback, act on the return value as set
'*              by the original WndProc.
'* hWnd       - Window handle.
'* uMsg       - Message value.
'* wParam     - Message related data.
'* lParam     - Message related data.
'* lParamUser - User-defined callback parameter
'*************************************************************************************************
End Sub
