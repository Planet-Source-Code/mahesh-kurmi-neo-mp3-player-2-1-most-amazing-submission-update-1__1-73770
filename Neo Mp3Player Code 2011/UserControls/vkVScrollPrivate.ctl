VERSION 5.00
Begin VB.UserControl vkVScrollPrivate 
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000004&
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   EditAtDesignTime=   -1  'True
   ScaleHeight     =   3600
   ScaleWidth      =   4800
   Begin VB.PictureBox picBarOver 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1005
      Left            =   4110
      MousePointer    =   99  'Custom
      Picture         =   "vkVScrollPrivate.ctx":0000
      ScaleHeight     =   1005
      ScaleWidth      =   180
      TabIndex        =   0
      Top             =   1410
      Visible         =   0   'False
      Width           =   180
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Left            =   600
      Top             =   0
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Left            =   0
      Top             =   120
   End
End
Attribute VB_Name = "vkVScrollPrivate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' =======================================================
'
' vkUserControlsXP
' Coded by violent_ken (Alain Descotes)
'
' =======================================================
'
' Some graphical UserControls for your VB application.
'
' Copyright © 2006-2007 by Alain Descotes.
'
' vkUserControlsXP is free software; you can redistribute it and/or
' modify it under the terms of the GNU Lesser General Public
' License as published by the Free Software Foundation; either
' version 2.1 of the License, or (at your option) any later version.
'
' vkUserControlsXP is distributed in the hope that it will be useful,
' but WITHOUT ANY WARRANTY; without even the implied warranty of
' MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the GNU
' Lesser General Public License for more details.
'
' You should have received a copy of the GNU Lesser General Public
' License along with this library; if not, write to the Free Software
' Foundation, Inc., 51 Franklin Street, Fifth Floor, Boston, MA  02110-1301  USA
'
' =======================================================


Option Explicit


Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hDC As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function CreateDC Lib "gdi32" Alias "CreateDCA" (ByVal lpDriverName As String, ByVal lpDeviceName As String, ByVal lpOutput As String, lpInitData As Any) As Long
Private Declare Function CreateHatchBrush Lib "gdi32" (ByVal fnStyle As Integer, ByVal COLORREF As Long) As Long
Private Declare Function CreatePatternBrush Lib "gdi32" (ByVal hBitmap As Long) As Long
Private Declare Function CreatePen Lib "gdi32" (ByVal nPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long) As Long
Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function DrawText Lib "user32" Alias "DrawTextA" (ByVal hDC As Long, ByVal lpStr As String, ByVal nCount As Long, lpRect As RECT, ByVal wFormat As Long) As Long
Private Declare Function DrawEdge Lib "user32" (ByVal hDC As Long, qrc As RECT, ByVal edge As Long, ByVal grfFlags As Long) As Long
Private Declare Function FillRect Lib "user32" (ByVal hDC As Long, lpRect As RECT, ByVal hBrush As Long) As Long
Private Declare Function FrameRect Lib "user32" (ByVal hDC As Long, lpRect As RECT, ByVal hBrush As Long) As Long
Private Declare Function GetClientRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Private Declare Function GetSysColor Lib "user32" (ByVal nIndex As Long) As Long
Private Declare Function LineTo Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function MoveToEx Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, lpPoint As POINTAPI) As Long
Private Declare Function PatBlt Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal dwRop As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long
Private Declare Function SetBkColor Lib "gdi32" (ByVal hDC As Long, ByVal crColor As Long) As Long
Private Declare Function SetBkMode Lib "gdi32" (ByVal hDC As Long, ByVal nBkMode As Long) As Long
Private Declare Function SetPixelV Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByVal crColor As Long) As Long
Private Declare Function SetTextColor Lib "gdi32" (ByVal hDC As Long, ByVal crColor As Long) As Long

Const BF_BOTTOM = &H8
Const BF_LEFT = &H1
Const BF_RIGHT = &H4
Const BF_TOP = &H2
Const BF_RECT = (BF_LEFT Or BF_TOP Or BF_RIGHT Or BF_BOTTOM)

'=======================================================
'VARIABLES PRIVEES
'=======================================================
Private mAsm(63) As Byte    'contient le code ASM
Private OldProc As Long     'adresse de l'ancienne window proc
Private objHwnd As Long     'handle de l'objet concerné
Private ET As TRACKMOUSEEVENTTYPE   'type pour le mouse_hover et le mouse_leave
Private IsMouseIn As Boolean    'si la souris est dans le controle

Private bCol As ole_color
Private lArrowColor As ole_color
Private lFrontColor As ole_color
Private lBorderColor As ole_color
Private lSelColor As ole_color
Private lLargeChangeColor As ole_color
Private lDownColor As ole_color
Private bEnable As Boolean
Private bMouseDown As Boolean

Private bNotOk As Boolean
Private bNotOk2 As Boolean
Private lScrollHeight As Byte
Private lMin As Currency
Private lMax As Currency
Private lValue As Currency
Private lSmallChange As Currency
Private lLargeChange As Currency
Private lWheelChange As Currency
Private bEnableWheel As Boolean
Private lPos1 As Long    'position haute du curseur
Private lPos2 As Long   'position basse du curseur
Private lGrise As Long
Private lUpMoused As Long
Private lDownMoused As Long
Private lMouseInterval As Long
Private lT As Long
Private lH As Long
Private nY As Long
Private n1 As Long
Private bUnRefreshControl As Boolean
Private bHasLeftOneTime As Boolean
Private bBlockVS As Boolean
Private bBlockValue As Boolean

'=======================================================
'EVENTS
'=======================================================
Public Event Change(Value As Currency)
Public Event KeyDown(KeyCode As Integer, Shift As Integer)
Public Event KeyPress(KeyAscii As Integer)
Public Event KeyUp(KeyCode As Integer, Shift As Integer)
Public Event MouseHover()
Public Event MouseLeave()
Public Event MouseWheel(Sens As Wheel_Sens)
Public Event MouseDown(Button As MouseButtonConstants, Shift As Integer, Control As Integer, X As Long, Y As Long)
Public Event MouseUp(Button As MouseButtonConstants, Shift As Integer, Control As Integer, X As Long, Y As Long)
Public Event MouseDblClick(Button As MouseButtonConstants, Shift As Integer, Control As Integer, X As Long, Y As Long)
Public Event MouseMove(Button As MouseButtonConstants, Shift As Integer, Control As Integer, X As Long, Y As Long)
Public Event Scroll()



'=======================================================
'USERCONTROL SUBS
'=======================================================
'=======================================================
' /!\ NE PAS DEPLACER CETTE FONCTION /!\ '
'=======================================================
' Cette fonction doit rester la premiere '
' fonction "public" du module de classe  '
'=======================================================
Public Function WindowProc(ByVal hwnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Attribute WindowProc.VB_MemberFlags = "40"
    Dim iControl As Integer
    Dim iShift As Integer
    Dim z As Long
    Dim X As Long
    Dim Y As Long

    Select Case uMsg

    Case WM_LBUTTONDBLCLK
        iShift = Abs((wParam And MK_SHIFT) = MK_SHIFT)
        iControl = Abs((wParam And MK_CONTROL) = MK_CONTROL)
        X = LoWord(lParam) * Screen.TwipsPerPixelX
        Y = HiWord(lParam) * Screen.TwipsPerPixelY

        If lUpMoused = 1 Then
            lUpMoused = 2: lValue = lValue - lSmallChange
        End If
        If lDownMoused = 1 Then
            lDownMoused = 2: lValue = lValue + lSmallChange
        End If

        If bEnable Then
            If Y > 255 And Y < lT Then
                n1 = -1
                lValue = lValue - lLargeChange
            End If
            If Y > lT + lH And Y < Height - 270 Then
                n1 = 1
                lValue = lValue + lLargeChange
            End If
        End If

        Call ChangeValues
        RaiseEvent Change(lValue)

        RaiseEvent MouseDblClick(vbLeftButton, iShift, iControl, X, Y)
    Case WM_LBUTTONDOWN
        iShift = Abs((wParam And MK_SHIFT) = MK_SHIFT)
        iControl = Abs((wParam And MK_CONTROL) = MK_CONTROL)
        X = LoWord(lParam) * Screen.TwipsPerPixelX
        Y = HiWord(lParam) * Screen.TwipsPerPixelY

        If lUpMoused Then
            lValue = lValue - lSmallChange
            lUpMoused = 2: Timer1.Enabled = True: ChangeValues: RaiseEvent Change(lValue)
        End If
        If lDownMoused Then
            lValue = lValue + lSmallChange
            lDownMoused = 2: Timer1.Enabled = True: ChangeValues: RaiseEvent Change(lValue)
        End If

        If bEnable Then
            If Y > 255 And Y < lT Then
                lValue = lValue - lLargeChange
                n1 = -1: ChangeValues: RaiseEvent Change(lValue)
            ElseIf (Y > lT + lH And Y < Height - 270) Then
                lValue = lValue + lLargeChange
                n1 = 1: ChangeValues: RaiseEvent Change(lValue)
            ElseIf Y >= lT And Y <= lH + lT Then
                bMouseDown = True: Refresh
            End If
            If Y > 255 And Y < lT Then
                Timer2.Enabled = True
            End If
            If Y > lT + lH And Y < Height - 270 Then
                Timer2.Enabled = True
            End If
        End If

        RaiseEvent MouseDown(vbLeftButton, iShift, iControl, X, Y)
    Case WM_LBUTTONUP
        iShift = Abs((wParam And MK_SHIFT) = MK_SHIFT)
        iControl = Abs((wParam And MK_CONTROL) = MK_CONTROL)
        X = LoWord(lParam) * Screen.TwipsPerPixelX
        Y = HiWord(lParam) * Screen.TwipsPerPixelY
        bMouseDown = False
        n1 = 0
        If Y > 255 And Y < lT Then
            Timer2.Enabled = False
        End If
        If Y > lT + lH And Y < Height - 270 Then
            Timer2.Enabled = False
        End If

        Call Refresh
        If lUpMoused = 2 Then
            lUpMoused = 1: Refresh: Timer1.Enabled = False
        End If
        If lDownMoused = 2 Then
            lDownMoused = 1: Refresh: Timer1.Enabled = False
        End If

        RaiseEvent MouseUp(vbLeftButton, iShift, iControl, X, Y)
    Case WM_MBUTTONDBLCLK
        iShift = Abs((wParam And MK_SHIFT) = MK_SHIFT)
        iControl = Abs((wParam And MK_CONTROL) = MK_CONTROL)
        X = LoWord(lParam) * Screen.TwipsPerPixelX
        Y = HiWord(lParam) * Screen.TwipsPerPixelY

        RaiseEvent MouseDblClick(vbMiddleButton, iShift, iControl, X, Y)
    Case WM_MBUTTONDOWN
        iShift = Abs((wParam And MK_SHIFT) = MK_SHIFT)
        iControl = Abs((wParam And MK_CONTROL) = MK_CONTROL)
        X = LoWord(lParam) * Screen.TwipsPerPixelX
        Y = HiWord(lParam) * Screen.TwipsPerPixelY

        RaiseEvent MouseDown(vbMiddleButton, iShift, iControl, X, Y)
    Case WM_MBUTTONUP
        iShift = Abs((wParam And MK_SHIFT) = MK_SHIFT)
        iControl = Abs((wParam And MK_CONTROL) = MK_CONTROL)
        X = LoWord(lParam) * Screen.TwipsPerPixelX
        Y = HiWord(lParam) * Screen.TwipsPerPixelY

        RaiseEvent MouseUp(vbMiddleButton, iShift, iControl, X, Y)
    Case WM_MOUSEHOVER
        If IsMouseIn = False Then
            RaiseEvent MouseHover
            IsMouseIn = True
        End If
    Case WM_MOUSELEAVE
        RaiseEvent MouseLeave
        IsMouseIn = False
        lUpMoused = 0
        lDownMoused = 0
        If bHasLeftOneTime Then
            Call Refresh
        Else
            bHasLeftOneTime = True
        End If
    Case WM_MOUSEMOVE
        Call TrackMouseEvent(ET)

        iShift = Abs((wParam And MK_SHIFT) = MK_SHIFT)
        iControl = Abs((wParam And MK_CONTROL) = MK_CONTROL)
        X = LoWord(lParam) * Screen.TwipsPerPixelX
        Y = HiWord(lParam) * Screen.TwipsPerPixelY

        If lUpMoused Then
            If Y > 255 Then
                'on vire le cadre de sélection
                lUpMoused = 0: Refresh
            End If
        End If
        If lDownMoused Then
            If Y < Height - 270 Then
                'on vire le cadre de sélection
                lDownMoused = 0: Refresh
            End If
        End If

        If lUpMoused = 0 And Y <= 255 Then lUpMoused = 1: Refresh
        If lDownMoused = 0 And Y >= Height - 270 Then lDownMoused = 1: Refresh

        If (wParam And MK_LBUTTON) = MK_LBUTTON Then z = vbLeftButton
        If (wParam And MK_RBUTTON) = MK_RBUTTON Then z = vbRightButton
        If (wParam And MK_MBUTTON) = MK_MBUTTON Then z = vbMiddleButton

        If z = vbLeftButton Then
            'alors clic gauche enfoncé
            If nY >= lT And nY <= lH + lT Then
                'alors c'est sur le curseur O_o !

                RaiseEvent Scroll

                lT = lT + Y - nY

                If lT <= 270 Then lT = 270
                If lT >= Height - 285 - lH Then lT = Height - 285 - lH

                lValue = Int((lT - 255) * (lMax - lMin) / (Height - 510 - lH)) + lMin

                If lT = 120 Then lValue = lMin
                If lT = Height - 285 - lH Then lValue = lMax

                Call Refresh
            End If
            ' RaiseEvent Change(lValue)
        End If

        'sauvegarde la position
        nY = Y
        RaiseEvent MouseMove(z, iShift, iControl, X, Y)

    Case WM_RBUTTONDBLCLK
        iShift = Abs((wParam And MK_SHIFT) = MK_SHIFT)
        iControl = Abs((wParam And MK_CONTROL) = MK_CONTROL)
        X = LoWord(lParam) * Screen.TwipsPerPixelX
        Y = HiWord(lParam) * Screen.TwipsPerPixelY

        RaiseEvent MouseDblClick(vbRightButton, iShift, iControl, X, Y)
    Case WM_RBUTTONDOWN
        iShift = Abs((wParam And MK_SHIFT) = MK_SHIFT)
        iControl = Abs((wParam And MK_CONTROL) = MK_CONTROL)
        X = LoWord(lParam) * Screen.TwipsPerPixelX
        Y = HiWord(lParam) * Screen.TwipsPerPixelY

        RaiseEvent MouseDown(vbRightButton, iShift, iControl, X, Y)
    Case WM_RBUTTONUP
        iShift = Abs((wParam And MK_SHIFT) = MK_SHIFT)
        iControl = Abs((wParam And MK_CONTROL) = MK_CONTROL)
        X = LoWord(lParam) * Screen.TwipsPerPixelX
        Y = HiWord(lParam) * Screen.TwipsPerPixelY

        RaiseEvent MouseUp(vbRightButton, iShift, iControl, X, Y)
    Case WM_MOUSEWHEEL
        If wParam < 0 Then
            RaiseEvent MouseWheel(WHEEL_DOWN)
            If bEnableWheel Then Me.Value = Me.Value + lWheelChange
            RaiseEvent Change(lValue)
        Else
            RaiseEvent MouseWheel(WHEEL_UP)
            If bEnableWheel Then Me.Value = Me.Value - lWheelChange
            RaiseEvent Change(lValue)
        End If
    Case WM_PAINT
        bNotOk = True  'évite le clignotement lors du survol de la souris
    End Select

    'appel de la routine standard pour les autres messages
    WindowProc = CallWindowProc(OldProc, hwnd, uMsg, wParam, lParam)

End Function

Private Sub Timer1_Timer()

    If bEnable = False Then Exit Sub

    If lUpMoused = 2 Then
        'on clique sur le bouton haut
        lValue = lValue - lSmallChange
        Call ChangeValues
        RaiseEvent Change(lValue)
    ElseIf lDownMoused = 2 Then
        'bouton du bas
        lValue = lValue + lSmallChange
        Call ChangeValues
        RaiseEvent Change(lValue)
    End If
End Sub

Private Sub Timer2_Timer()

    If bEnable = False Or n1 = 0 Then Exit Sub

    If n1 = 1 Then
        'largechange en bas
        lValue = lValue + lLargeChange
        Call ChangeValues
        RaiseEvent Change(lValue)
    ElseIf n1 = -1 Then
        lValue = lValue - lLargeChange
        Call ChangeValues
        RaiseEvent Change(lValue)
    End If

    If lValue = lMax Or lValue = lMin Then Timer2.Enabled = False

End Sub

Private Sub UserControl_Initialize()
    Dim Ofs As Long
    Dim Ptr As Long

    'Recupere l'adresse de "Me.WindowProc"
    Call CopyMemory(Ptr, ByVal (ObjPtr(Me)), 4)
    Call CopyMemory(Ptr, ByVal (Ptr + 489 * 4), 4)

    'Crée la veritable fonction WindowProc (à optimiser)
    Ofs = VarPtr(mAsm(0))
    MovL Ofs, &H424448B            '8B 44 24 04          mov         eax,dword ptr [esp+4]
    MovL Ofs, &H8245C8B            '8B 5C 24 08          mov         ebx,dword ptr [esp+8]
    MovL Ofs, &HC244C8B            '8B 4C 24 0C          mov         ecx,dword ptr [esp+0Ch]
    MovL Ofs, &H1024548B           '8B 54 24 10          mov         edx,dword ptr [esp+10h]
    MovB Ofs, &H68                 '68 44 33 22 11       push        Offset RetVal
    MovL Ofs, VarPtr(mAsm(59))
    MovB Ofs, &H52                 '52                   push        edx
    MovB Ofs, &H51                 '51                   push        ecx
    MovB Ofs, &H53                 '53                   push        ebx
    MovB Ofs, &H50                 '50                   push        eax
    MovB Ofs, &H68                 '68 44 33 22 11       push        ObjPtr(Me)
    MovL Ofs, ObjPtr(Me)
    MovB Ofs, &HE8                 'E8 1E 04 00 00       call        Me.WindowProc
    MovL Ofs, Ptr - Ofs - 4
    MovB Ofs, &HA1                 'A1 20 20 40 00       mov         eax,RetVal
    MovL Ofs, VarPtr(mAsm(59))
    MovL Ofs, &H10C2               'C2 10 00             ret         10h
End Sub

Private Sub UserControl_InitProperties()
'valeurs par défaut
    bNotOk2 = True
    With Me
        .ArrowColor = vbWhite
        .BackColor = vbWhite
        .BorderColor = &HFF8080
        .Enabled = True
        .EnableWheel = True
        .FrontColor = 15782079
        .LargeChange = 10
        .Max = 100
        .Min = 0
        .ScrollHeight = 10
        .SmallChange = 1
        .Value = 50
        .WheelChange = 3
        .DownColor = 12492429
        .MouseHoverColor = vbWhite
        .MouseInterval = 100
        .LargeChangeColor = 12492429
        .UnRefreshControl = True
    End With
    bNotOk2 = False
    Call UserControl_Paint  'refresh
    Timer1.Enabled = True
    Timer2.Enabled = True
End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)

    If bEnable = False Then Exit Sub

    Select Case KeyCode
    Case vbKeyLeft, vbKeyUp
        lValue = lValue - SmallChange: ChangeValues: RaiseEvent Change(lValue)
    Case vbKeyRight, vbKeyDown
        lValue = lValue + SmallChange: ChangeValues: RaiseEvent Change(lValue)
    Case vbKeyPageUp
        lValue = lValue - LargeChange: ChangeValues: RaiseEvent Change(lValue)
    Case vbKeyPageDown
        lValue = lValue + LargeChange: ChangeValues: RaiseEvent Change(lValue)
    Case vbKeyEnd
        lValue = lMax: ChangeValues: RaiseEvent Change(lValue)
    Case vbKeyHome
        lValue = lMin: ChangeValues: RaiseEvent Change(lValue)
    End Select

    RaiseEvent KeyDown(KeyCode, Shift)
End Sub

Private Sub UserControl_KeyPress(KeyAscii As Integer)
    RaiseEvent KeyPress(KeyAscii)
End Sub

Private Sub UserControl_KeyUp(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyUp(KeyCode, Shift)
End Sub
Private Sub UserControl_Terminate()
'vire le subclassing
    If OldProc Then Call SetWindowLong(UserControl.hwnd, GWL_WNDPROC, OldProc)
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    With PropBag
        Call .WriteProperty("ArrowColor", Me.ArrowColor, vbWhite)
        Call .WriteProperty("BackColor", Me.BackColor, vbWhite)
        Call .WriteProperty("BorderColor", Me.BorderColor, &HFF8080)
        Call .WriteProperty("Enabled", Me.Enabled, True)
        Call .WriteProperty("EnableWheel", Me.EnableWheel, True)
        Call .WriteProperty("FrontColor", Me.FrontColor, 15782079)
        Call .WriteProperty("LargeChange", Me.LargeChange, 10)
        Call .WriteProperty("Max", Me.Max, 100)
        Call .WriteProperty("Min", Me.Min, 0)
        Call .WriteProperty("WheelChange", Me.WheelChange, 3)
        Call .WriteProperty("ScrollHeight", Me.ScrollHeight, 10)
        Call .WriteProperty("SmallChange", Me.SmallChange, 1)
        Call .WriteProperty("Value", Me.Value, 50)
        Call .WriteProperty("MouseHoverColor", Me.MouseHoverColor, vbWhite)
        Call .WriteProperty("DownColor", Me.DownColor, 12492429)
        Call .WriteProperty("MouseInterval", Me.MouseInterval, 100)
        Call .WriteProperty("LargeChangeColor", Me.LargeChangeColor, 12492429)
        Call .WriteProperty("UnRefreshControl", Me.UnRefreshControl, True)
    End With
End Sub
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    bNotOk2 = True
    With PropBag
        Me.ArrowColor = .ReadProperty("ArrowColor", vbWhite)
        Me.BackColor = .ReadProperty("BackColor", vbWhite)
        Me.BorderColor = .ReadProperty("BorderColor", &HFF8080)
        Me.Enabled = .ReadProperty("Enabled", True)
        Me.EnableWheel = .ReadProperty("EnableWheel", True)
        Me.FrontColor = .ReadProperty("FrontColor", 15782079)
        Me.LargeChange = .ReadProperty("LargeChange", 10)
        Me.Max = .ReadProperty("Max", 100)
        Me.Min = .ReadProperty("Min", 0)
        Me.ScrollHeight = .ReadProperty("ScrollHeight", 10)
        Me.SmallChange = .ReadProperty("SmallChange", 1)
        Me.Value = .ReadProperty("Value", 50)
        Me.WheelChange = .ReadProperty("WheelChange", 3)
        Me.MouseHoverColor = .ReadProperty("MouseHoverColor", vbWhite)
        Me.DownColor = .ReadProperty("DownColor", 12492429)
        Me.MouseInterval = .ReadProperty("MouseInterval", 100)
        Me.LargeChangeColor = .ReadProperty("LargeChangeColor", 12492429)
        Me.UnRefreshControl = .ReadProperty("UnRefreshControl", True)
    End With
    bNotOk2 = False
    'Call UserControl_Paint  'refresh
    Call Refresh

    'le bon endroit pour lancer le subclassing
    'Call LaunchKeyMouseEvents
End Sub
Private Sub UserControl_Resize()
'    If Height < 800 Then Height = 800
'  If Width < 255 Then Width = 255

'lScrollHeight représente le pourcentage de la hauteur
'calcule la hauteur du curseur
    lH = Int((Height - 510) * lScrollHeight / 100)

    Call ChangeValues
    'Call Refresh
End Sub

'=======================================================
'lance le subclassing
'=======================================================
Public Sub LaunchKeyMouseEvents()

    If Ambient.UserMode Then

        OldProc = SetWindowLong(UserControl.hwnd, GWL_WNDPROC, _
                                VarPtr(mAsm(0)))    'pas de AddressOf aujourd'hui ;)

        'prépare le terrain pour le mouse_over et mouse_leave
        With ET
            .cbSize = Len(ET)
            .hwndTrack = UserControl.hwnd
            .dwFlags = TME_LEAVE Or TME_HOVER
            .dwHoverTime = 1
        End With

        'démarre le tracking de l'entrée
        Call TrackMouseEvent(ET)

        'pas dedans par défaut
        IsMouseIn = False

    End If

End Sub



'=======================================================
'PROPERTIES
'=======================================================
Public Property Get hDC() As Long: hDC = UserControl.hDC: End Property
Public Property Get hwnd() As Long: hwnd = UserControl.hwnd: End Property
Public Property Get BackColor() As ole_color: BackColor = bCol: End Property
Public Property Let BackColor(BackColor As ole_color): bCol = BackColor: bNotOk = False: UserControl_Paint: End Property
Public Property Get BorderColor() As ole_color: BorderColor = lBorderColor: End Property
Public Property Let BorderColor(BorderColor As ole_color): lBorderColor = BorderColor: bNotOk = False: UserControl_Paint: End Property
Public Property Get ArrowColor() As ole_color: ArrowColor = lArrowColor: End Property
Public Property Let ArrowColor(ArrowColor As ole_color): lArrowColor = ArrowColor: bNotOk = False: UserControl_Paint: End Property
Public Property Get FrontColor() As ole_color: FrontColor = lFrontColor: End Property
Public Property Let FrontColor(FrontColor As ole_color): lFrontColor = FrontColor: bNotOk = False: UserControl_Paint: End Property
Public Property Get Enabled() As Boolean: Enabled = bEnable: End Property
Public Property Let Enabled(Enabled As Boolean)
    If bEnable <> Enabled Then
        bEnable = Enabled: bNotOk = False: UserControl_Paint
    End If
End Property
Public Property Get EnableWheel() As Boolean: EnableWheel = bEnableWheel: End Property
Public Property Let EnableWheel(EnableWheel As Boolean): bEnableWheel = EnableWheel: bNotOk = False: UserControl_Paint: End Property
Public Property Get ScrollHeight() As Byte: ScrollHeight = lScrollHeight: End Property
Public Property Let ScrollHeight(ScrollHeight As Byte)
    lScrollHeight = ScrollHeight
    'lScrollHeight représente le pourcentage de la hauteur
    'calcule la hauteur du curseur
    lH = Int((Height - 510) * lScrollHeight / 100)
ChangeValues:     bNotOk = False: UserControl_Paint
End Property
Public Property Get Min() As Currency: Min = lMin: End Property
Public Property Let Min(Min As Currency)
    If lMin > lValue Then Exit Property
    If lMin <> Min Then
        lMin = Min
        ChangeValues
        bNotOk = False: UserControl_Paint
    End If
End Property
Public Property Get Max() As Currency: Max = lMax: End Property
Public Property Let Max(Max As Currency)
'If lMax < lValue Then Exit Property
    If Max <> lMax Then
        lMax = Max
        ChangeValues
        bNotOk = False: UserControl_Paint
    End If
End Property
Public Property Get Value() As Currency: Value = lValue: End Property
Public Property Let Value(Value As Currency)
    If bBlockValue Then Exit Property
    If Value <> lValue Then
        If bBlockVS = False Then RaiseEvent Change(lValue)
        lValue = Value: Call ChangeValues
    End If
    bBlockVS = False
End Property
Public Property Get SmallChange() As Currency: SmallChange = lSmallChange: End Property
Public Property Let SmallChange(SmallChange As Currency): lSmallChange = SmallChange: bNotOk = False: UserControl_Paint: End Property
Public Property Get LargeChange() As Currency: LargeChange = lLargeChange: End Property
Public Property Let LargeChange(LargeChange As Currency): lLargeChange = LargeChange: bNotOk = False: UserControl_Paint: End Property
Public Property Get WheelChange() As Currency: WheelChange = lWheelChange: End Property
Public Property Let WheelChange(WheelChange As Currency): lWheelChange = WheelChange: bNotOk = False: UserControl_Paint: End Property
Public Property Get DownColor() As ole_color: DownColor = lDownColor: End Property
Public Property Let DownColor(DownColor As ole_color): lDownColor = DownColor: bNotOk = False: UserControl_Paint: End Property
Public Property Get MouseHoverColor() As ole_color: MouseHoverColor = lSelColor: End Property
Public Property Let MouseHoverColor(MouseHoverColor As ole_color): lSelColor = MouseHoverColor: bNotOk = False: UserControl_Paint: End Property
Public Property Get MouseInterval() As Long: MouseInterval = lMouseInterval: End Property
Public Property Let MouseInterval(MouseInterval As Long): lMouseInterval = MouseInterval: Timer1.Interval = lMouseInterval: Timer2.Interval = lMouseInterval: End Property
Public Property Get LargeChangeColor() As ole_color: LargeChangeColor = lLargeChangeColor: End Property
Public Property Let LargeChangeColor(LargeChangeColor As ole_color): lLargeChangeColor = LargeChangeColor: bNotOk = False: UserControl_Paint: End Property
Public Property Get UnRefreshControl() As Boolean: UnRefreshControl = bUnRefreshControl: End Property
Public Property Let UnRefreshControl(UnRefreshControl As Boolean): bUnRefreshControl = UnRefreshControl: End Property
Public Property Let BlockVS(BlockVS As Boolean): bBlockVS = BlockVS: End Property
Public Property Let BlockValue(BlockValue As Boolean): bBlockValue = BlockValue: End Property


Private Sub UserControl_Paint()

    If bNotOk Or bNotOk2 Then Exit Sub     'pas prêt à peindre

    Call Refresh    'on refresh
End Sub




'=======================================================
'PRIVATE SUBS
'=======================================================
'=======================================================
'copie un "byte"
'=======================================================
Private Sub MovB(Ofs As Long, ByVal Value As Long)
    Call CopyMemory(ByVal Ofs, Value, 1): Ofs = Ofs + 1
End Sub

'=======================================================
'copie un "long"
'=======================================================
Private Sub MovL(Ofs As Long, ByVal Value As Long)
    Call CopyMemory(ByVal Ofs, Value, 4): Ofs = Ofs + 4
End Sub

'=======================================================
'récupère la hauteur d'un caractère
'=======================================================
Private Function GetCharHeight() As Long
    Dim Res As Long
    Res = GetTabbedTextExtent(UserControl.hDC, "A", 1, 0, 0)
    GetCharHeight = (Res And &HFFFF0000) \ &H10000
End Function

'=======================================================
'change la valeur Value
'=======================================================
Private Sub ChangeValues()

    If lValue > lMax Then lValue = lMax
    If lValue < lMin Then lValue = lMin

    'calcule le Top du curseur
    If lMax <> lMin Then _
       lT = By15(Int(Abs((Height - 510 - lH) * (lValue - lMin) / (lMax - lMin))) + 255)

    If lT <= 270 Then lT = 270
    If lT >= Height - 285 - lH Then lT = Height - 285 - lH

    'refresh le controle
    bNotOk = False: Call UserControl_Paint
End Sub

'=======================================================
'MAJ du controle
'=======================================================
Public Sub Refresh()
    Dim X As Long
    Dim Y As Long
    Dim m As RECT
    If bUnRefreshControl Then Exit Sub

    '//on efface
    Call UserControl.cls

    '//convertit les couleurs
    Call OleTranslateColor(lArrowColor, 0, lArrowColor)
    Call OleTranslateColor(lFrontColor, 0, lFrontColor)
    Call OleTranslateColor(bCol, 0, bCol)
    Call OleTranslateColor(lBorderColor, 0, lBorderColor)

    '//on trace les bords haut et bas, leur bordure et la bordure générale
    'contour général
    'Line (0, 0)-(Width, Height), lBorderColor, BF

    m.Left = 0
    m.Top = 0
    m.Bottom = Height / 15
    m.Right = Width / 15
    DrawRectangle m, lBorderColor, UserControl.hDC
    'repeind intérieur en backcolor
    'Line (15, 15)-(Width - 30, Height - 30), bCol, BF

    m.Left = 1
    m.Top = 1
    m.Bottom = Height / 15 - 1
    m.Right = Width / 15 - 1
    DrawFillRectangle m, bCol, UserControl.hDC

    'zone haut
    'Line (15, 15)-(Width - 30, 255), lFrontColor, BF
    m.Left = 1
    m.Top = 1
    m.Bottom = 18
    m.Right = Width / 15 - 1
    'DrawRectangle m, lFrontColor, UserControl.hDC
    If lUpMoused <= 1 Then
        DrawGradient lLargeChangeColor, lSelColor, 1, 1, Width / 15 - 1, 9, UserControl.hDC, False
        DrawGradient lSelColor, lLargeChangeColor, 1, 9, Width / 15 - 1, 18, UserControl.hDC, False
    ElseIf lUpMoused = 2 Then
        DrawGradient lLargeChangeColor, BlendColor(lSelColor, lLargeChangeColor), 1, 1, Width / 15 - 1, 9, UserControl.hDC, False
        DrawGradient BlendColor(lSelColor, lLargeChangeColor), lLargeChangeColor, 1, 9, Width / 15 - 1, 18, UserControl.hDC, False
    End If
    DrawRectangle m, lSelColor, UserControl.hDC

    'zone bas
    ' Line (15, Height - 285)-(Width - 30, Height - 30), lFrontColor, BF
    m.Left = 1
    m.Top = Height / 15 - 18
    m.Bottom = Height / 15 - 1
    m.Right = Width / 15 - 1
    'DrawRectangle m, lFrontColor, UserControl.hDC
    If lDownMoused <= 1 Then
        DrawGradient lLargeChangeColor, lSelColor, 1, Height / 15 - 18, Width / 15 - 1, Height / 15 - 10, UserControl.hDC, False
        DrawGradient lSelColor, lLargeChangeColor, 1, Height / 15 - 10, Width / 15 - 1, Height / 15 - 2, UserControl.hDC, False
    ElseIf lDownMoused = 2 Then
        DrawGradient lLargeChangeColor, BlendColor(lSelColor, lLargeChangeColor), 1, Height / 15 - 18, Width / 15 - 1, Height / 15 - 10, UserControl.hDC, False
        DrawGradient BlendColor(lSelColor, lLargeChangeColor), lLargeChangeColor, 1, Height / 15 - 10, Width / 15 - 1, Height / 15 - 2, UserControl.hDC, False
    End If
    DrawRectangle m, lSelColor, UserControl.hDC

    'lignes de séparation zones haut et bas
    ' Line (0, 270)-(Width, 270), lBorderColor

    DrawLine 0, 18, Width / 15, 18, UserControl.hDC, lBorderColor
    'Line (0, Height - 285)-(Width, Height - 285), lBorderColor
    DrawLine 0, Height / 15 - 18, Width / 15, Height / 15 - 18, UserControl.hDC, lBorderColor


    '//trace les rectangles de sélection/pushed/rien
    'Call DrawSelRectUp
    'Call DrawSelRectDown

    '//on trace les flèches
    'si Enabled=false on met la couleur 10070188 aux flèches
    DrawArrow

    If bEnable = False Then
        'boulot terminé !
        Call UserControl.Refresh
        bNotOk = True
        Exit Sub
    End If

    '//on trace le curseur
    'Line (15, lT)-(Width - 30, lT + lH), lBorderColor ', BF  'bordure
    'Line (15, lT + 15)-(Width - 30, lT + lH - 15), lFrontColor, BF

    'Line (0, lT)-(0, lT + lH), lBorderColor ', BF  'bordure
    'Line (Width / 15, lT + 15)-(Width / 15, lT + lH), lBorderColor
    'Line (0, lT)-(0, lT + lH), lBorderColor ', BF  'bordure
    'Line (Width / 15, lT + 15)-(Width / 15, lT + lH), lBorderColor
    ' StretchBlt UserControl.hDC, 0, lT / 15, Width, lH / 15, picBarOver.hDC, 0, 0, picBarOver.ScaleWidth, picBarOver.ScaleHeight, vbSrcCopy
    m.Left = 0
    m.Right = Width / 15
    m.Top = lT / 15
    m.Bottom = (lT + lH) / 15
    ' UserControl.PaintPicture picBarOver, 1, lT + 1, Width - 1, lH - 1, 0, 0

    If bMouseDown = False Then
        DrawGradient lSelColor, lLargeChangeColor, 2, lT / 15, (Width) / 15 - 1, (2 * lH / 3 + lT) / 15, UserControl.hDC, False
        DrawGradient lLargeChangeColor, BlendColor(lSelColor, lLargeChangeColor), 2, (lT + 2 * lH / 3) / 15, (Width) / 15 - 1, (lH + lT) / 15, UserControl.hDC, False
    Else
        DrawGradient lLargeChangeColor, BlendColor(lSelColor, lLargeChangeColor), 2, lT / 15, (Width) / 15 - 1, (2 * lH / 3 + lT) / 15, UserControl.hDC, False
        DrawGradient BlendColor(lSelColor, lLargeChangeColor), lLargeChangeColor, 2, (lT + 2 * lH / 3) / 15, (Width) / 15 - 1, (lH + lT) / 15, UserControl.hDC, False
    End If

    DrawGradient lLargeChangeColor, lSelColor, 1, lT / 15, 2, (lH + lT) / 15, UserControl.hDC, True
    DrawGradient lSelColor, lLargeChangeColor, Width / 15 - 1, lT / 15, Width / 15 - 1, (lH + lT) / 15, UserControl.hDC, True
    '
    '//on trace un rectangle de LargeChange

    m.Left = 2
    m.Right = Width / 15 - 2
    m.Top = lT / 15 + 1    '(lT + lH) / 15 - 1
    m.Bottom = (lT + lH) / 15 - 1
    DrawRectangle m, lSelColor, UserControl.hDC

    m.Left = 0
    m.Right = Width / 15
    m.Top = lT / 15
    m.Bottom = (lT + lH) / 15 + 1

    DrawRectangle m, lBorderColor, UserControl.hDC
    ' Call DrawLine(0, lT / 15 - 2, Width / 15, lH, UserControl.hDC, lBorderColor)
    If n1 = -1 Then
        m.Left = 1
        m.Top = 19
        m.Bottom = lT / 15 - 1
        m.Right = Width / 15 - 1
        DrawFillRectangle m, lLargeChangeColor, UserControl.hDC
        'en haut
        ' Line (15, 285)-(Width - 30, lT - 15), lLargeChangeColor, BF
    ElseIf n1 = 1 Then
        'en bas
        m.Left = 1
        m.Top = (lT + lH) / 15 + 1
        m.Bottom = Height / 15 - 19
        m.Right = Width / 15 - 1
        DrawFillRectangle m, lLargeChangeColor, UserControl.hDC
        ' Line (15, 285)-(Width - 30, lT - 15), lLargeChangeColor, BF
    End If

    '//on refresh le control
    Call UserControl.Refresh

    bNotOk = True
End Sub

'=======================================================
'trace le rectangle de sélection de la fleche du haut
'=======================================================
Private Sub DrawSelRectUp()

'lUpMoused 1 (lignes blanches, survol) 2 (lignes foncées, pushed) 0 (rien)
    Dim m As RECT

    If bEnable = False Then Exit Sub

    Call OleTranslateColor(lSelColor, 0, lSelColor)
    Call OleTranslateColor(lDownColor, 0, lDownColor)

    Select Case lUpMoused

    Case 0
        Exit Sub

    Case 1
        'survol

        UserControl.ForeColor = lSelColor
        Line (15, 15)-(Width - 15, 15)
        Line (15, 15)-(15, 255)
        Line (15, 255)-(Width - 15, 255)
        Line (Width - 30, 240)-(Width - 30, 15)

    Case 2
        'clic
        m.Left = 1
        m.Top = 1
        m.Bottom = 18
        m.Right = Width / 15 - 2
        'Line (15, 15)-(Width - 30, 255), lDownColor, BF
        DrawRectangle m, lDownColor, UserControl.hDC
    End Select

    Call UserControl.Refresh
End Sub

'=======================================================
'trace le rectangle de sélection de la fleche du bas
'=======================================================
Private Sub DrawSelRectDown()

'lDownMoused 1 (lignes blanches, survol) 2 (lignes foncées, pushed) 0 (rien)

    If bEnable = False Then Exit Sub

    Call OleTranslateColor(lSelColor, 0, lSelColor)
    Call OleTranslateColor(lDownColor, 0, lDownColor)

    Select Case lDownMoused

    Case 0
        Exit Sub

    Case 1
        'survol

        UserControl.ForeColor = lSelColor
        Line (15, Height - 30)-(Width - 15, Height - 30)
        Line (15, Height - 30)-(15, Height - 270)
        Line (15, Height - 270)-(Width - 15, Height - 270)
        Line (Width - 30, Height - 270)-(Width - 30, Height - 30)

    Case 2
        'clic

        Line (15, Height - 270)-(Width - 30, Height - 30), lDownColor, BF
    End Select

End Sub

'=======================================================
'renvoie une valeur divisible par 15 (supérieure à l)
'=======================================================
Private Function By15(ByVal l As Currency) As Currency
    By15 = Int((l + 14) / 15) * 15
End Function

'=======================================================
'renvoie l'objet extender de ce usercontrol (pour les propertypages)
'=======================================================
Friend Property Get MyExtender() As Object
    Set MyExtender = UserControl.Extender
End Property
Friend Property Let MyExtender(MyExtender As Object)
    Set UserControl.Extender = MyExtender
End Property


'======================================================================
'DRAWS A LINE WITH A DEFINED COLOR
Private Sub DrawLine( _
        ByVal X As Long, _
        ByVal Y As Long, _
        ByVal X2 As Long, _
        ByVal Y2 As Long, _
        ByVal cHdc As Long, _
        ByVal Color As Long)

    Dim Pen1 As Long
    Dim Pen2 As Long
    Dim Pos As POINTAPI

    Pen1 = CreatePen(0, 1, GetLngColor(Color))
    Pen2 = SelectObject(cHdc, Pen1)

    MoveToEx cHdc, X, Y, Pos
    LineTo cHdc, X2, Y2

    SelectObject cHdc, Pen2
    DeleteObject Pen2
    DeleteObject Pen1

End Sub
'======================================================================

'======================================================================
'DRAWS A 2 COLOR GRADIENT AREA WITH A PREDEFINED DIRECTION
Public Sub DrawGradient(lEndColor As Long, lStartcolor As Long, ByVal X As Long, ByVal Y As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal hDC As Long, Optional bH As Boolean)
    On Error Resume Next

    ''Draw a Vertical Gradient in the current HDC
    Dim sR As Single, sG As Single, sB As Single
    Dim eR As Single, eG As Single, eB As Single
    Dim ni As Long

    lEndColor = GetLngColor(lEndColor)
    lStartcolor = GetLngColor(lStartcolor)

    sR = (lStartcolor And &HFF)
    sG = (lStartcolor \ &H100) And &HFF
    sB = (lStartcolor And &HFF0000) / &H10000
    eR = (lEndColor And &HFF)
    eG = (lEndColor \ &H100) And &HFF
    eB = (lEndColor And &HFF0000) / &H10000
    sR = (sR - eR) / IIf(bH, X2 - X, Y2 - Y)
    sG = (sG - eG) / IIf(bH, X2 - X, Y2 - Y)
    sB = (sB - eB) / IIf(bH, X2 - X, Y2 - Y)

    If bH Then
        For ni = 0 To X2 - X
            DrawLine X + ni, Y, X + ni, Y2, hDC, RGB(eR + (ni * sR), eG + (ni * sG), eB + (ni * sB))
        Next ni
    Else
        For ni = 0 To Y2 - Y
            DrawLine X, Y + ni, X2, Y + ni, hDC, RGB(eR + (ni * sR), eG + (ni * sG), eB + (ni * sB))
        Next ni
    End If
End Sub
'======================================================================



'======================================================================
'BLENDS 2 COLORS WITH A PREDEFINED ALPHA VALUE
Private Function BlendColor(ByVal oColorFrom As ole_color, ByVal oColorTo As ole_color, Optional ByVal Alpha As Long = 128) As Long
    Dim lCFrom As Long
    Dim lCTo As Long
    Dim lSrcR As Long
    Dim lSrcG As Long
    Dim lSrcB As Long
    Dim lDstR As Long
    Dim lDstG As Long
    Dim lDstB As Long

    lCFrom = GetLngColor(oColorFrom)
    lCTo = GetLngColor(oColorTo)

    lSrcR = lCFrom And &HFF
    lSrcG = (lCFrom And &HFF00&) \ &H100&
    lSrcB = (lCFrom And &HFF0000) \ &H10000
    lDstR = lCTo And &HFF
    lDstG = (lCTo And &HFF00&) \ &H100&
    lDstB = (lCTo And &HFF0000) \ &H10000

    BlendColor = RGB( _
                 ((lSrcR * Alpha) / 255) + ((lDstR * (255 - Alpha)) / 255), _
                 ((lSrcG * Alpha) / 255) + ((lDstG * (255 - Alpha)) / 255), _
                 ((lSrcB * Alpha) / 255) + ((lDstB * (255 - Alpha)) / 255) _
                 )

End Function
'======================================================================

'======================================================================
'CONVERTION FUNCTION
Private Function GetLngColor(Color As Long) As Long
    If (Color And &H80000000) Then
        GetLngColor = GetSysColor(Color And &H7FFFFFFF)
    Else
        GetLngColor = Color
    End If
End Function
'======================================================================

'======================================================================
'DRAWS A BORDER RECTANGLE AREA OF AN SPECIFIED COLOR
Private Sub DrawRectangle(ByRef bRect As RECT, ByVal Color As Long, ByVal hDC As Long)

    Dim hBrush As Long

    hBrush = CreateSolidBrush(Color)
    FrameRect hDC, bRect, hBrush
    DeleteObject hBrush

End Sub
'======================================================================


'======================================================================
'DRAWS A FILL RECTANGLE AREA OF AN SPECIFIED COLOR
Private Sub DrawFillRectangle(ByRef hRect As RECT, ByVal Color As Long, ByVal MyHdc As Long)

    Dim hBrush As Long

    hBrush = CreateSolidBrush(GetLngColor(Color))
    FillRect MyHdc, hRect, hBrush
    DeleteObject hBrush

End Sub
'======================================================================


Private Sub DrawArrow()
    Dim X As Long
    Dim lColor As Long
    Call OleTranslateColor(lArrowColor, 0, lArrowColor)
    X = (Width - 255) / 2 + 15
    'Dim color when pressed
    lColor = IIf(lUpMoused <= 1, lArrowColor, BlendColor(lArrowColor, bCol))
    If bEnable = False Then lColor = 10070188
    '///draw upper triangle
    DrawLine 7 + X / 15, 6, 8 + X / 15, 6, UserControl.hDC, lColor
    DrawLine 6 + X / 15, 7, 9 + X / 15, 7, UserControl.hDC, lColor
    DrawLine 5 + X / 15, 8, 10 + X / 15, 8, UserControl.hDC, lColor
    DrawLine 4 + X / 15, 9, 11 + X / 15, 9, UserControl.hDC, lColor

    'Dim color when pressed
    lColor = IIf(lDownMoused <= 1, lArrowColor, BlendColor(lArrowColor, bCol))
    If bEnable = False Then lColor = 10070188

    '///draw lower triangle
    DrawLine 7 + X / 15, Height / 15 - 7, 8 + X / 15, Height / 15 - 7, UserControl.hDC, lColor
    DrawLine 6 + X / 15, Height / 15 - 8, 9 + X / 15, Height / 15 - 8, UserControl.hDC, lColor
    DrawLine 5 + X / 15, Height / 15 - 9, 10 + X / 15, Height / 15 - 9, UserControl.hDC, lColor
    DrawLine 4 + X / 15, Height / 15 - 10, 11 + X / 15, Height / 15 - 10, UserControl.hDC, lColor

End Sub
