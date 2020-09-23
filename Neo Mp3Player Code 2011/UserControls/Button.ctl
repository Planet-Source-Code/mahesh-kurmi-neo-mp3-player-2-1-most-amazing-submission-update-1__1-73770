VERSION 5.00
Begin VB.UserControl Button 
   ClientHeight    =   1095
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2850
   ClipControls    =   0   'False
   DefaultCancel   =   -1  'True
   EditAtDesignTime=   -1  'True
   LockControls    =   -1  'True
   Picture         =   "Button.ctx":0000
   ScaleHeight     =   73
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   190
   ToolboxBitmap   =   "Button.ctx":0312
   Begin VB.Timer OverTimer 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   0
      Top             =   0
   End
End
Attribute VB_Name = "Button"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "PropPageWizardRun" ,"Yes"
Private Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, ByVal hDC As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long


Private Declare Function TransparentBlt Lib "msimg32" _
                                        (ByVal hdcDst As Long, ByVal nXOriginDst As Long, _
                                         ByVal nYOriginDst As Long, ByVal nWidthDst As Long, _
                                         ByVal nHeightDst As Long, ByVal hDcSrc As Long, _
                                         ByVal nXOriginSrc As Long, ByVal nYOriginSrc As Long, _
                                         ByVal nWidthSrc As Long, ByVal nHeightSrc As Long, _
                                         ByVal crTransparent As Long) As Long


Private Declare Function DrawIconEx Lib "user32" (ByVal hDC As Long, ByVal xLeft As Long, ByVal yTop As Long, ByVal hIcon As Long, ByVal cxWidth As Long, ByVal cyHeight As Long, ByVal istepIfAniCur As Long, ByVal hbrFlickerFreeDraw As Long, ByVal diFlags As Long) As Long
' DrawIconEx constants
Private Const DI_MASK = &H1
Private Const DI_IMAGE = &H2
Private Const DI_NORMAL = &H3
Private Const DI_COMPAT = &H4
Private Const DI_DEFAULTSIZE = &H8


Private Declare Function SetCapture Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function WindowFromPoint Lib "user32" (ByVal xPoint As Long, ByVal yPoint As Long) As Long
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long


Enum AlignConstants
    [AlignNone]
    [AlignTop]
    [AlignBottom]
    [AlignLeft]
    [AlignRight]
End Enum


Enum ButtonStyleConstants
    [Standard]
    [Graphical]
End Enum

Dim g_3DInc As Integer

Dim g_MouseDown As Boolean, g_MouseIn As Boolean, g_Selected As Boolean
Attribute g_MouseIn.VB_VarUserMemId = 1073938433
Attribute g_Selected.VB_VarUserMemId = 1073938433
Dim g_Button As Integer, g_Shift As Integer, g_X As Single, g_Y As Single
Attribute g_Button.VB_VarUserMemId = 1073938436
Attribute g_Shift.VB_VarUserMemId = 1073938436
Attribute g_X.VB_VarUserMemId = 1073938436
Attribute g_Y.VB_VarUserMemId = 1073938436

Const m_def_Style = 0               'Standard
Const m_def_UseMaskColor = False
Const m_def_PictureAlign = 0        'AlignNone (Center)

'Property Variables:
Dim m_Style As ButtonStyleConstants
Attribute m_Style.VB_VarUserMemId = 1073938440
Dim m_UseMaskColor As Boolean
Attribute m_UseMaskColor.VB_VarUserMemId = 1073938441
Dim m_PictureAlign As AlignConstants
Attribute m_PictureAlign.VB_VarUserMemId = 1073938442

'Dim m_PictureBack As StdPicture
Dim m_PictureNormal As StdPicture
Attribute m_PictureNormal.VB_VarUserMemId = 1073938443
Dim m_PictureDown As StdPicture
Attribute m_PictureDown.VB_VarUserMemId = 1073938444
Dim m_PictureOver As StdPicture
Attribute m_PictureOver.VB_VarUserMemId = 1073938445
Dim m_PictureDisabled As StdPicture
Attribute m_PictureDisabled.VB_VarUserMemId = 1073938446

Dim g_Light As ole_color
Attribute g_Light.VB_VarUserMemId = 1073938447
Dim g_Shadow As ole_color
Attribute g_Shadow.VB_VarUserMemId = 1073938448
Dim g_HighLight As ole_color
Attribute g_HighLight.VB_VarUserMemId = 1073938449
Dim g_DarkShadow As ole_color
Attribute g_DarkShadow.VB_VarUserMemId = 1073938450

'Event Declarations:
Event Click()
Event KeyDown(KeyCode As Integer, Shift As Integer)
Event KeyPress(KeyAscii As Integer)
Event KeyUp(KeyCode As Integer, Shift As Integer)
Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Event MouseIn(Shift As Integer)
Event MouseOut(Shift As Integer)



Private Sub UserControl_Click()
    RaiseEvent Click
End Sub

'##############################################################################
'  Init / read / write properties
'################################################################################

Private Sub UserControl_InitProperties()


    m_Style = m_def_Style
    m_UseMaskColor = m_def_UseMaskColor
    m_PictureAlign = m_def_PictureAlign


    Set m_PictureNormal = LoadPicture("")
    Set m_PictureDisabled = LoadPicture("")
    Set m_PictureDown = LoadPicture("")
    Set m_PictureOver = LoadPicture("")

    UserControl.BackColor = Ambient.BackColor

End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    m_Style = PropBag.ReadProperty("Style", m_def_Style)
    m_UseMaskColor = PropBag.ReadProperty("UseMaskColor", m_def_UseMaskColor)
    m_PictureAlign = PropBag.ReadProperty("PictureAlign", m_def_PictureAlign)

    Set UserControl.Picture = PropBag.ReadProperty("PictureBack", Nothing)
    Set m_PictureNormal = PropBag.ReadProperty("PictureNormal", Nothing)
    Set m_PictureDisabled = PropBag.ReadProperty("PictureDisabled", Nothing)
    Set m_PictureDown = PropBag.ReadProperty("PictureDown", Nothing)
    Set m_PictureOver = PropBag.ReadProperty("PictureOver", Nothing)
    Set MouseIcon = PropBag.ReadProperty("MouseIcon", Nothing)

    UserControl.BackColor = PropBag.ReadProperty("ButtonColor", &H8000000F)

    g_Selected = PropBag.ReadProperty("Selected", False)
    UserControl.MaskColor = PropBag.ReadProperty("MaskColor", &H8000000F)
    UserControl.MousePointer = PropBag.ReadProperty("MousePointer", 0)

    Refresh
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
' write property is to make initial property persistnet so that
' whenever control is kept on a form these are updated
    Call PropBag.WriteProperty("ButtonColor", UserControl.BackColor, &H8000000F)
    Call PropBag.WriteProperty("Selected", g_Selected, False)
    Call PropBag.WriteProperty("PictureAlign", m_PictureAlign, m_def_PictureAlign)
    Call PropBag.WriteProperty("MaskColor", UserControl.MaskColor, &H8000000F)
    Call PropBag.WriteProperty("MouseIcon", MouseIcon, Nothing)
    Call PropBag.WriteProperty("MousePointer", UserControl.MousePointer, 0)
    Call PropBag.WriteProperty("PictureBack", UserControl.Picture, Nothing)
    Call PropBag.WriteProperty("PictureNormal", m_PictureNormal, Nothing)
    Call PropBag.WriteProperty("PictureDisabled", m_PictureDisabled, Nothing)
    Call PropBag.WriteProperty("PictureDown", m_PictureDown, Nothing)
    Call PropBag.WriteProperty("PictureOver", m_PictureOver, Nothing)
    Call PropBag.WriteProperty("Style", m_Style, m_def_Style)
    Call PropBag.WriteProperty("UseMaskColor", m_UseMaskColor, m_def_UseMaskColor)

End Sub


'################################################################################
'  'Ambient' control
'################################################################################
Private Sub UserControl_Resize()

    Refresh

End Sub

Public Sub Refresh()

    AutoRedraw = True

    UserControl.cls

    'Draw picture

    If m_Style = Graphical Then DrawPicture

    AutoRedraw = False

End Sub


'################################################################################
'  Events
'################################################################################

Private Sub UserControl_DblClick()

    SetCapture hwnd    'Preseve hWnd on DblClick
    UserControl_MouseDown g_Button, g_Shift, g_X, g_Y

End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    g_Button = Button: g_Shift = Shift: g_X = X: g_Y = Y

    If Button <> vbRightButton Then

        g_MouseDown = True
        Refresh

    End If

    RaiseEvent MouseDown(Button, Shift, X, Y)

End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If (X >= 0 And Y >= 0) And (X < ScaleWidth And Y < ScaleHeight) Then

        If g_MouseIn = False Then

            OverTimer.Enabled = True
            g_MouseIn = True

            RaiseEvent MouseIn(Shift)

            Refresh

        End If

    End If

    RaiseEvent MouseMove(Button, Shift, X, Y)

End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

    g_MouseDown = False

    'If Button <> vbRightButton Then

    Refresh
    '  If (x >= 0 And y >= 0) And (x < ScaleWidth And y < ScaleHeight) Then RaiseEvent Click

    'End If

    RaiseEvent MouseUp(Button, Shift, X, Y)

End Sub




'################################################################################
'  Properties
'################################################################################
'get property returns value of property through itself
Public Property Get PictureAlign() As AlignConstants

    PictureAlign = m_PictureAlign

End Property

Public Property Let PictureAlign(ByVal New_PictureAlign As AlignConstants)

    m_PictureAlign = New_PictureAlign
    PropertyChanged "PictureAlign"    'it tells contril toperform all necessary tasks when property changs
    'e.g. height is changed control should be redrawn.
    Refresh

End Property

'ButtonColor ####################################################################

Public Property Get ButtonColor() As ole_color

    ButtonColor = UserControl.BackColor

End Property

Public Property Let ButtonColor(ByVal New_ButtonColor As ole_color)

    UserControl.BackColor = New_ButtonColor
    PropertyChanged "ButtonColor"

    Refresh

End Property

'Selected ########################################################################
Public Property Get Selected() As Boolean

    Selected = g_Selected

End Property

Public Property Let Selected(ByVal New_Selected As Boolean)

    g_Selected = New_Selected
    PropertyChanged "Selected"

    Refresh

End Property

'hWnd ###########################################################################
Public Property Get hwnd() As Long

    hwnd = UserControl.hwnd

End Property

'MaskColor ######################################################################
Public Property Get MaskColor() As ole_color

    MaskColor = UserControl.MaskColor

End Property

Public Property Let MaskColor(ByVal New_MaskColor As ole_color)

    UserControl.MaskColor() = New_MaskColor
    PropertyChanged "MaskColor"

    Refresh

End Property

'MousePointer & MouseIcon #######################################################
Public Property Get MousePointer() As MousePointerConstants

    MousePointer = UserControl.MousePointer

End Property

Public Property Let MousePointer(ByVal New_MousePointer As MousePointerConstants)

    UserControl.MousePointer() = New_MousePointer
    PropertyChanged "MousePointer"

End Property

Public Property Get MouseIcon() As StdPicture

    Set MouseIcon = UserControl.MouseIcon

End Property

Public Property Set MouseIcon(ByVal New_MouseIcon As StdPicture)

    Set UserControl.MouseIcon = New_MouseIcon
    PropertyChanged "MouseIcon"

End Property

'Picture, PictureNormal,PictureDisabled, PictureDown & PictureOver ############################
Public Property Get PictureBack() As StdPicture

    Set PictureBack = UserControl.Picture

End Property

Public Property Set PictureBack(ByVal New_Picture As StdPicture)

    Set UserControl.Picture = New_Picture
    PropertyChanged "PictureBack"

    Refresh
End Property


Public Property Get PictureNormal() As StdPicture

    Set PictureNormal = m_PictureNormal

End Property

Public Property Set PictureNormal(ByVal New_Picture As StdPicture)

    Set m_PictureNormal = New_Picture
    PropertyChanged "PictureNormal"

    Refresh
End Property

Public Property Get PictureDisabled() As StdPicture

    Set PictureDisabled = m_PictureDisabled

End Property

Public Property Set PictureDisabled(ByVal New_PictureDisabled As StdPicture)

    Set m_PictureDisabled = New_PictureDisabled
    PropertyChanged "PictureDisabled"

    Refresh

End Property

Public Property Get PictureDown() As StdPicture

    Set PictureDown = m_PictureDown

End Property

Public Property Set PictureDown(ByVal New_PictureDown As StdPicture)

    Set m_PictureDown = New_PictureDown
    PropertyChanged "PictureDown"

    Refresh

End Property

Public Property Get PictureOver() As StdPicture

    Set PictureOver = m_PictureOver

End Property

Public Property Set PictureOver(ByVal New_PictureOver As StdPicture)

    Set m_PictureOver = New_PictureOver
    PropertyChanged "PictureOver"

    Refresh

End Property

'Style ##########################################################################
Public Property Get Style() As ButtonStyleConstants

    Style = m_Style

End Property

Public Property Let Style(ByVal New_Style As ButtonStyleConstants)

    m_Style = New_Style
    PropertyChanged "Style"

    Refresh

End Property

'UseMaskColor ###################################################################
Public Property Get UseMaskColor() As Boolean

    UseMaskColor = m_UseMaskColor

End Property

Public Property Let UseMaskColor(ByVal New_UseMaskColor As Boolean)

    m_UseMaskColor = New_UseMaskColor
    PropertyChanged "UseMaskColor"
    Refresh

End Property

Public Sub Reset()
    Set m_PictureNormal = LoadPicture("")
    Set m_PictureDisabled = LoadPicture("")
    Set m_PictureDown = LoadPicture("")
    Set m_PictureOver = LoadPicture("")
    UserControl.MouseIcon = LoadPicture()
End Sub

'DrawPicture ####################################################################
'            1. Get picture by actual state
'            2. If no image in actual state: take normal state picture
'               If no normal state picture: exit sub
'            3. Set picture position by align mode
'            4. Readjust drawed text left/right margins
'            5. If UseMaskColor = True draw picture with standard PaintPicture Maskcolor means the color specifying transparent area in maskpicture
'               If not case:
'                  a) BMP, DIB, GIF, JPG: TransparentBlt function
'                     (StdPicture not accepted -> CreateCompatibleDC)
'                  b) ICO, CUR:           DrawIconEx function
'                     (Transp. 'ability' included in this type)
'                  c) WMF, EMF:           Standard PaintPicture function
'                     (Transp. 'ability' included in this type)
'                  d) Invalid picture

Private Sub DrawPicture()

    Set tmpPicture = New StdPicture
    Dim PosInc As Integer, PosX As Integer, PosY As Integer
    Dim W As Integer, H As Integer

    'Set tmpPicture by button state:
    If g_MouseDown Then
        'Mouse down
        Set tmpPicture = m_PictureDown  ': PosInc = 1
    ElseIf g_MouseIn And g_Selected = False Then
        'Mouse in (over)and it is not selected yet so picture should be changed
        Set tmpPicture = m_PictureOver
    ElseIf g_Selected = True Then
        'Button disabled
        Set tmpPicture = m_PictureDisabled
    Else
        'Mouse out
        Set tmpPicture = m_PictureNormal
    End If

    If tmpPicture Is Nothing Then
        If m_PictureNormal Is Nothing Then
            'No picture
            Exit Sub
        Else
            'Use default picture for actual state
            Set tmpPicture = m_PictureNormal
        End If
    End If

    If tmpPicture = 0 Then Exit Sub    'Filter if not initialized

    g_TextWithPicture = True        'We have a picture

    'Set drawed picture dimensions (cms to pixels)
    W = Int(tmpPicture.Width / 26.1)
    H = Int(tmpPicture.Height / 26.1)

    'Set drawed picture location
    Select Case m_PictureAlign

        Dim MaxPicture As Integer

    Case 0    'None (center picture)
        PosX = Int((ScaleWidth - W) / 2) + PosInc
        PosY = Int((ScaleHeight - H) / 2) + PosInc

    Case 1    'Top
        PosX = Int((ScaleWidth - W) / 2) + PosInc
        PosY = PosInc + MaxPicture + 3

    Case 2    'Bottom
        PosX = Int((ScaleWidth - W) / 2) + PosInc
        PosY = (ScaleHeight - H) + PosInc - MaxPicture - 4

    Case 3    'Left
        PosX = PosInc + MaxPicture + 3
        PosY = Int((ScaleHeight - H) / 2) + PosInc

    Case 4    'Right
        PosX = (ScaleWidth - W) + PosInc - MaxPicture - 4
        PosY = Int((ScaleHeight - H) / 2) + PosInc

    End Select


    If m_UseMaskColor Then

        Select Case tmpPicture.Type

        Case vbPicTypeBitmap        ' BMP, DIB, GIF, JPG

            hDCScreen = GetDC(0&)

            hDcSrc = CreateCompatibleDC(hDCScreen)
            SelectObject hDcSrc, tmpPicture.Handle

            '???: TransparentBlt turns to 0 nXOriginDst and nYOriginDst values
            '     If PosX or PosY < 0 -> The picture can't be centered

            TransparentBlt hDC, PosX, PosY, W, H, _
                           hDcSrc, 0, 0, W, H, MaskColor

            DeleteDC hDcSrc
            ReleaseDC 0&, hDCScreen

        Case vbPicTypeIcon          ' ICO, CUR

            DrawIconEx hDC, PosX, PosY, tmpPicture.Handle, W, H, 0, 0, DI_NORMAL Or DI_DEFAULTSIZE

        Case vbPicTypeMetafile, _
             vbPicTypeEMetafile     ' WMF, EMF

            PaintPicture tmpPicture, PosX, PosY

        Case Else                   ' Invalid picture

            err.Raise 481

        End Select

    Else

        PaintPicture tmpPicture, PosX, PosY

    End If

End Sub


'Timer ##########################################################################
'      Use of WindowFromPoint(X,Y) function
'      1. Get handle of actual absolute mouse position
'      2. If UserControl handle <> returned handle : Out of button
'         (See: Sub UserControl_MouseMove)

Private Sub OverTimer_Timer()

    Dim p As POINTAPI

    GetCursorPos p

    If hwnd <> WindowFromPoint(p.X, p.Y) Then

        OverTimer.Enabled = False

        g_MouseIn = False
        RaiseEvent MouseOut(g_Shift)

        Refresh                     'Refresh picture

        If g_MouseDown = True Then  'Resfresh state
            g_MouseDown = False
            Refresh
            g_MouseDown = True
        End If

    End If

End Sub




