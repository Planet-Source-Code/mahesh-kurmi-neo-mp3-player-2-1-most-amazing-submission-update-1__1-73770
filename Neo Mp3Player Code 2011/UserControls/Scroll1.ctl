VERSION 5.00
Begin VB.UserControl Eq_SliderCtrl 
   AutoRedraw      =   -1  'True
   ClientHeight    =   4035
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3915
   ScaleHeight     =   269
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   261
   ToolboxBitmap   =   "Scroll1.ctx":0000
   Begin VB.PictureBox picBarOver 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   720
      MousePointer    =   99  'Custom
      ScaleHeight     =   13
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   8
      TabIndex        =   4
      Top             =   0
      Visible         =   0   'False
      Width           =   120
   End
   Begin VB.Timer OverTimer 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   330
      Top             =   3510
   End
   Begin VB.PictureBox picBar 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   480
      MousePointer    =   99  'Custom
      ScaleHeight     =   13
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   8
      TabIndex        =   3
      Top             =   0
      Visible         =   0   'False
      Width           =   120
   End
   Begin VB.PictureBox picBarDown 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   0
      MousePointer    =   99  'Custom
      ScaleHeight     =   13
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   8
      TabIndex        =   2
      Top             =   0
      Visible         =   0   'False
      Width           =   120
   End
   Begin VB.PictureBox picBack1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1035
      Left            =   240
      ScaleHeight     =   69
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   8
      TabIndex        =   1
      Top             =   0
      Visible         =   0   'False
      Width           =   120
   End
   Begin VB.PictureBox picBack 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1035
      Left            =   0
      MousePointer    =   99  'Custom
      ScaleHeight     =   69
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   8
      TabIndex        =   0
      Top             =   0
      Width           =   120
   End
End
Attribute VB_Name = "Eq_SliderCtrl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Many bugs fixed
'Value error removed


Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function WindowFromPoint Lib "user32" (ByVal xPoint As Long, ByVal yPoint As Long) As Long
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long

Private Type POINTAPI
    X As Long
    Y As Long
End Type


' Declarations
Dim iY As Long
Dim bDrag As Boolean
Dim iMin As Double
Dim iMax As Double
Dim iValue As Long
'Dim iSelected As Boolean
Private bMouseOver As Boolean, bMouseDown As Boolean
Attribute bMouseDown.VB_VarUserMemId = 1073938437
Private iLargeChange As Integer
Attribute iLargeChange.VB_VarUserMemId = 1073938439


Public Enum ePos
    Verticall = 0
    Horizontall = 1
End Enum

Private Enum eImg
    Normal = 0
    Down = 1
    Over = 2
End Enum
Private bEnabled As Boolean
Attribute bEnabled.VB_VarUserMemId = 1073938440
Private ePosition As ePos
Attribute ePosition.VB_VarUserMemId = 1073938441
' Events
Event KeyUp(KeyCode As Integer, Shift As Integer)
Event Change(Value As Long)
Event KeyPress(KeyAscii As Integer)
Event KeyDown(KeyCode As Integer, Shift As Integer)
Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

'//--------------------------------------------------------------------------

Public Sub ResetPictures()
    picBack.Picture = LoadPicture()
    picBack1.Picture = LoadPicture()
    picBar.Picture = LoadPicture()
    picBarOver.Picture = LoadPicture()
    picBarDown.Picture = LoadPicture()
    picBack.MouseIcon = LoadPicture()
End Sub

Public Property Get MouseIcon() As Picture
    Set MouseIcon = picBar.MouseIcon
End Property

Public Property Set MouseIcon(ByVal New_Icon As Picture)
    Set picBack.MouseIcon = New_Icon

    PropertyChanged "MouseIcon"
End Property

Public Property Get BackColor() As ole_color
    BackColor = picBack.BackColor
End Property

Public Property Let BackColor(ByVal New_Color As ole_color)
    picBack.BackColor = New_Color
    picBack1.BackColor = New_Color

    PropertyChanged "BackColor"
End Property
Public Property Get Position() As ePos
    Position = ePosition
End Property

Public Property Let Position(ByVal NewValue As ePos)
    On Error Resume Next
    Dim W As Integer, H As Integer
    ePosition = NewValue


    If picBar.Picture <> 0 Then
        picBar.AutoSize = True
    Else
        picBar.Width = 9: picBar.Height = 9
    End If

    picBarOver.Width = picBar.Width: picBarOver.Height = picBar.Height
    picBarDown.Width = picBar.Width: picBarDown.Height = picBar.Height

    W = ScaleWidth
    H = ScaleHeight

    UserControl.Width = H * 15
    UserControl.Height = W * 15

    picBar.AutoSize = False
    picBarDown.AutoSize = False
    picBarOver.AutoSize = False

    UserControl_Resize

    PropertyChanged "Position"
End Property

Public Property Get bar() As Picture
    Set bar = picBar.Picture
End Property

Public Property Set bar(ByVal New_Bar As Picture)
    Set picBar.Picture = New_Bar

    picBar.AutoSize = True

    If picBarDown.Picture = 0 Then
        picBarDown.Picture = picBar.Picture
        picBarDown.AutoSize = True
    End If

    If picBarOver.Picture = 0 Then
        picBarOver.Picture = picBar.Picture
        picBarOver.AutoSize = True
    End If

    picBar.AutoSize = False
    picBarDown.AutoSize = False
    picBarOver.AutoSize = False


    Call DrawBar(Normal)
    PropertyChanged "Bar"
End Property

Public Property Get BarDown() As Picture
    Set BarDown = picBarDown.Picture
End Property

Public Property Set BarDown(ByVal New_Bar As Picture)
    Set picBarDown.Picture = New_Bar
    picBarDown.AutoSize = True
    picBarDown.AutoSize = False
    PropertyChanged "BarDown"
End Property

Public Property Get BarOver() As Picture
    Set BarOver = picBarOver.Picture
End Property

Public Property Set BarOver(ByVal New_Bar As Picture)
    Set picBarOver.Picture = New_Bar
    picBarOver.AutoSize = True
    picBarOver.AutoSize = False
    PropertyChanged "BarOver"
End Property


Private Sub CalcValue()
    On Error Resume Next
    If ePosition = Verticall Then
        iValue = (iY) / (picBack.Height - picBar.Height) * (iMax - iMin) + iMin
        'iValue = (iMax - iValue) - iMin
        If iMin < 0 Then iValue = -iValue Else iValue = iMax - iValue
    Else
        iValue = iY / (picBack.Width - picBar.Width) * (iMax - iMin) + iMin
    End If
End Sub


Private Sub DrawBar(ImgState As eImg, Optional CalculateX As Boolean = True)
    On Error Resume Next
    Dim intY As Integer, intX As Integer


    If CalculateX Then
        If ePosition = Verticall Then
            'If iMin < 0 Then iValue = -iValue Else iValue = iMax - iValue
            iY = (iMax - iValue) / (iMax - iMin) * (picBack.Height - picBar.Height)
            intX = 0: intY = iY
        Else
            iY = (iValue - iMin) / (iMax - iMin) * (picBack.Width - picBar.Width)
            intY = 0: intX = iY
        End If

    Else
        If ePosition = Verticall Then intX = 0: intY = iY Else intX = iY: intY = 0
    End If

    picBack.cls

    '// draw progress
    If ePosition = Verticall Then
        Call BitBlt(picBack.hDC, intX, intY, picBack1.ScaleWidth, picBack1.ScaleHeight, _
                    picBack1.hDC, intX, intY, vbSrcCopy)
    Else
        Call BitBlt(picBack.hDC, 0, 0, intX, picBack1.ScaleHeight, _
                    picBack1.hDC, 0, 0, vbSrcCopy)
    End If

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
        UserControl.Refresh
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
    UserControl.Refresh
End Sub
Public Property Get Max() As Long
    Max = iMax
End Property

Public Property Let Max(New_Max As Long)
    If iValue > New_Max Then iValue = New_Max

    iMax = New_Max
    Call DrawBar(Normal)

    PropertyChanged "Max"
End Property

Public Property Get Min() As Long
    Min = iMin
End Property

Public Property Let Min(New_Min As Long)
    If New_Min > iValue Then iValue = New_Min

    iMin = New_Min
    Call DrawBar(Normal)

    PropertyChanged "Min"
End Property

Public Property Get LargeChange() As Integer
    LargeChange = iLargeChange
End Property

Public Property Let LargeChange(New_Value As Integer)
    If New_Value >= iMax Then Exit Property

    iLargeChange = New_Value

    PropertyChanged "LargeChange"
End Property


Public Property Get PictureBack() As Picture
    Set PictureBack = picBack.Picture
End Property

Public Property Set PictureBack(ByVal New_Picture As Picture)
    Set picBack.Picture = New_Picture
    picBack.AutoSize = True
    picBack.AutoSize = False
    '    UserControl.Width = picBack.ScaleWidth * 15
    '    UserControl.Height = picBack.ScaleHeight * 15

    If picBack1.Picture = 0 Then
        picBack1.Picture = picBack.Picture
        picBack1.AutoSize = True
        picBack1.AutoSize = False
    End If

    Call DrawBar(Normal)

    PropertyChanged "PictureBack"
End Property
Public Property Get PictureProgress() As Picture
    Set PictureProgress = picBack1.Picture
End Property

Public Property Set PictureProgress(ByVal New_Picture2 As Picture)
    Set picBack1.Picture = New_Picture2
    picBack1.AutoSize = True
    picBack1.AutoSize = False

    Call DrawBar(Normal)

    PropertyChanged "PictureProgress"
End Property


Public Property Get Value() As Long
    Value = iValue
End Property

Public Property Let Value(New_Value As Long)
    If New_Value < iMin Or New_Value > iMax Then Exit Property
    'If bMouseDown = True Then Exit Property
    iValue = New_Value
    Call DrawBar(Normal, True)
    RaiseEvent Change(iValue)
    PropertyChanged "Value"
End Property
Public Property Get Enabled() As Boolean
    Enabled = bEnabled
End Property
Public Property Let Enabled(boolEnabled As Boolean)
    bEnabled = boolEnabled
    PropertyChanged "Enabled"
    UserControl.Enabled = bEnabled
    picBack.Enabled = bEnabled
End Property


Private Sub picBack_KeyDown(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyDown(KeyCode, Shift)
End Sub

Private Sub picBack_KeyPress(KeyAscii As Integer)
    RaiseEvent KeyPress(KeyAscii)
End Sub

Private Sub picBack_KeyUp(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyUp(KeyCode, Shift)
End Sub

Private Sub picBack_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        Dim k As Integer
        'k = iValue
        bDrag = True
        bMouseDown = True
        '// Verticall
        RaiseEvent MouseDown(Button, Shift, X, Y)

        If ePosition = Verticall Then
            'iy means top position of bar to be drawn

            iY = Y
            If iY > picBack.ScaleHeight - (picBar.ScaleHeight / 2) Then
                iY = picBack.ScaleHeight - (picBar.ScaleHeight)
            ElseIf iY < picBar.ScaleHeight / 2 Then
                iY = 0
            Else
                iY = iY - picBar.ScaleHeight / 2
            End If

            k = (picBack.Height - picBar.Height - iY) / (picBack.Height - picBar.Height) * (iMax - iMin) + iMin
            'CalcValue

            If iValue <> k Then
                iValue = k
                RaiseEvent Change(iValue)
            End If

            Call DrawBar(Down, True)
            ' RaiseEvent Change(iValue)
            'Else
            ' If y > iY Then '// sumar
            '   Value = Value + LargeChange
            '  Else
            '  Value = Value - LargeChange
            '  End If
            'Call CalcValue
            ' RaiseEvent Change(iValue)
            ' Call DrawBar(down)
            'End If

        Else    '// Horizontall
            If X >= iY And X <= iY + picBar.ScaleWidth And Button = 1 Then
                bDrag = True
                bMouseDown = True
                Call DrawBar(Down, False)
            Else
                'If iLargeChange = 0 Then
                iY = X
                If iY > picBack.ScaleWidth - (picBar.ScaleWidth / 2) Then iY = picBack.ScaleWidth - (picBar.ScaleWidth / 2)
                If iY < picBar.ScaleWidth / 2 Then iY = picBar.ScaleWidth / 2
                iY = iY - picBar.ScaleWidth / 2
                Call CalcValue
                If iValue <> k Then
                    RaiseEvent Change(iValue)
                End If
                Call DrawBar(Down, False)
                ' Else
                ' If x > iY Then '// sumar
                '   Value = Value + LargeChange
                ' Else
                '   Value = Value - LargeChange
                ' End If
                'End If
            End If

        End If

        ' RaiseEvent MouseDown(Button, Shift, X, Y)
    End If
End Sub

Private Sub picBack_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error Resume Next
    If bDrag Then  '// dragging
        Dim k As Integer
        If Button = 1 Then

            'k = iValue
            'iSelected = True
            '// Verticall
            If ePosition = Verticall Then
                iY = Y
                If iY > picBack.ScaleHeight - (picBar.ScaleHeight / 2) Then
                    iY = picBack.ScaleHeight - (picBar.ScaleHeight)
                ElseIf iY < picBar.ScaleHeight / 2 Then
                    iY = 0
                Else
                    iY = iY - picBar.ScaleHeight / 2
                End If

                k = (picBack.Height - picBar.Height - iY) / (picBack.Height - picBar.Height) * (iMax - iMin) + iMin
                'CalcValue

                If iValue <> k Then
                    iValue = k
                    RaiseEvent Change(iValue)
                End If

                Call DrawBar(Down, True)
                '// Horizontall
            Else
                iY = X

                If iY > picBack.Width - (picBar.Width / 2) Then iY = picBack.Width - (picBar.Width / 2)

                If iY < picBar.Width / 2 Then iY = picBar.Width / 2

                iY = iY - picBar.Width / 2

                CalcValue
                RaiseEvent Change(iValue)
                'End If
                Call DrawBar(Down, False)
            End If
            ' k = (picBack.Height - picBar.Height - iY) / (picBack.Height - picBar.Height) * (iMax - iMin) + iMin

            'If iValue <> k Then
            ' iValue = k

            'RaiseEvent Change(iValue)
        End If
    Else
        '// mouse over
        If ePosition = Verticall Then
            If bMouseOver = False Then
                bMouseOver = True
                Call DrawBar(Over, False)
                OverTimer.Enabled = True
            End If
        Else
            If bMouseOver = False Then
                bMouseOver = True
                Call DrawBar(Over, False)
                OverTimer.Enabled = True
            End If
        End If
    End If

    RaiseEvent MouseMove(Button, Shift, X, Y)

End Sub

Private Sub picBack_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
'If bDrag = False Then
'Call CalcValue
'RaiseEvent Change(iValue)
'End If
    If ePosition = Horizontall Then
        Call CalcValue
        bMouseDown = False
        Call DrawBar(Normal)
    Else
        bMouseDown = False
        Call DrawBar(Normal)
    End If
    bDrag = False

    RaiseEvent MouseUp(Button, Shift, X, Y)
End Sub


Private Sub UserControl_Initialize()
    If iMax = 0 Then iMax = 100
    Call DrawBar(Normal)
    Selected = False
    bEnabled = True
    ' lpPrevWndProc = SetWindowLong(Me.hwnd, GWL_WNDPROC, AddressOf WindowProc)
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    picBack.Picture = PropBag.ReadProperty("PictureBack", Nothing)
    picBar.MouseIcon = PropBag.ReadProperty("MouseIcon", Nothing)
    picBack1.Picture = PropBag.ReadProperty("PictureProgress", Nothing)
    picBarDown.Picture = PropBag.ReadProperty("BarDown", Nothing)
    picBarOver.Picture = PropBag.ReadProperty("BarOver", Nothing)
    picBar.Picture = PropBag.ReadProperty("Bar", Nothing)
    picBack.BackColor = PropBag.ReadProperty("BackColor", &H8000000F)
    picBack1.BackColor = picBack.BackColor
    iMin = PropBag.ReadProperty("Min", 0)
    iMax = PropBag.ReadProperty("Max", 100)
    iLargeChange = PropBag.ReadProperty("LargeChange", 0)
    iValue = PropBag.ReadProperty("Value", 0)
    Position = PropBag.ReadProperty("Position", 0)

    Call DrawBar(Normal)
End Sub

Private Sub UserControl_Resize()
    picBack.Width = UserControl.ScaleWidth
    picBack.Height = UserControl.ScaleHeight
    picBack1.Width = picBack.Width
    picBack1.Height = picBack.Height
    If ePosition = Verticall Then
        picBar.Width = UserControl.ScaleWidth
        picBarDown.Width = UserControl.ScaleWidth
        picBarOver.Width = UserControl.ScaleWidth
    Else
        picBar.Height = UserControl.ScaleHeight
        picBarDown.Height = UserControl.ScaleHeight
        picBarOver.Height = UserControl.ScaleHeight
    End If
    Call DrawBar(Normal)
End Sub


Private Sub UserControl_Terminate()
' Call SetWindowLong(Me.hwnd, GWL_WNDPROC, lpPrevWndProc)

End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("PictureBack", picBack.Picture, Nothing)
    Call PropBag.WriteProperty("MouseIcon", picBar.MouseIcon, Nothing)
    Call PropBag.WriteProperty("PictureProgress", picBack1.Picture, Nothing)
    Call PropBag.WriteProperty("Bar", picBar.Picture, Nothing)
    Call PropBag.WriteProperty("BarOver", picBarOver.Picture, Nothing)
    Call PropBag.WriteProperty("BarDown", picBarDown.Picture, Nothing)
    Call PropBag.WriteProperty("BackColor", picBack.BackColor, &H8000000F)
    Call PropBag.WriteProperty("Min", iMin, 0)
    Call PropBag.WriteProperty("Max", iMax, 100)
    Call PropBag.WriteProperty("LargeChange", iLargeChange, 0)
    Call PropBag.WriteProperty("Value", iValue, 0)
    Call PropBag.WriteProperty("Position", ePosition, 0)
End Sub


Private Sub OverTimer_Timer()

    Dim p As POINTAPI

    GetCursorPos p

    If picBack.hwnd <> WindowFromPoint(p.X, p.Y) Then

        OverTimer.Enabled = False
        bMouseOver = False
        Call DrawBar(Normal, False)

    End If

End Sub


'Public Property Get Selected() As Boolean
'selcted = pselected
'End Property

'Public Property Let Selected(ByVal vNewValue As Boolean)
'iSelected = vNewValue
'PropertyChanged "selected"
'End Property


