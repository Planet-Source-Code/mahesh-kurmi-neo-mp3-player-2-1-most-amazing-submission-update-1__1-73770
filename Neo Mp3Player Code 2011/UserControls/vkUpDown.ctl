VERSION 5.00
Begin VB.UserControl vkUpDown 
   AutoRedraw      =   -1  'True
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   PropertyPages   =   "vkUpDown.ctx":0000
   ScaleHeight     =   3600
   ScaleWidth      =   4800
   ToolboxBitmap   =   "vkUpDown.ctx":002B
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Left            =   1920
      Top             =   840
   End
End
Attribute VB_Name = "vkUpDown"
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
Private lDownColor As ole_color
Private bEnable As Boolean
Private lSens As Direction
Private bNotOk As Boolean
Private bNotOk2 As Boolean
Private lMin As Currency
Private lMax As Currency
Private lValue As Currency
Private lSmallChange As Currency
Private lPos1 As Long    'position haute du curseur
Private lPos2 As Long   'position basse du curseur
Private lGrise As Long
Private lUpMoused As Long
Private lDownMoused As Long
Private lMouseInterval As Long
Private bUnRefreshControl As Boolean
Private bHasLeftOneTime As Boolean


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
Attribute WindowProc.VB_Description = "Internal proc for subclassing"
    Dim iControl As Integer
    Dim iShift As Integer
    Dim z As Long
    Dim X As Long
    Dim Y As Long
    Dim z2 As Long

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

        RaiseEvent MouseDown(vbLeftButton, iShift, iControl, X, Y)
    Case WM_LBUTTONUP
        iShift = Abs((wParam And MK_SHIFT) = MK_SHIFT)
        iControl = Abs((wParam And MK_CONTROL) = MK_CONTROL)
        X = LoWord(lParam) * Screen.TwipsPerPixelX
        Y = HiWord(lParam) * Screen.TwipsPerPixelY

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

        If lSens = [Up_Down] Then
            z = Height / 2
            z2 = Y
        Else
            z = Width / 2
            z2 = X
        End If

        If lUpMoused Then
            If z2 > z Then
                'on vire le cadre de sélection
                lUpMoused = 0: Refresh
            End If
        End If
        If lDownMoused Then
            If z2 < z Then
                'on vire le cadre de sélection
                lDownMoused = 0: Refresh
            End If
        End If

        If lUpMoused = 0 And z2 <= z Then lUpMoused = 1: Refresh
        If lDownMoused = 0 And z2 >= z Then lDownMoused = 1: Refresh

        If (wParam And MK_LBUTTON) = MK_LBUTTON Then z = vbLeftButton
        If (wParam And MK_RBUTTON) = MK_RBUTTON Then z = vbRightButton
        If (wParam And MK_MBUTTON) = MK_MBUTTON Then z = vbMiddleButton

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
        Else
            RaiseEvent MouseWheel(WHEEL_UP)
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
        .BorderColor = &HFF8080
        .Enabled = True
        .FrontColor = 15782079
        .Max = 100
        .Min = 0
        .SmallChange = 1
        .Value = 50
        .DownColor = 12492429
        .MouseHoverColor = vbWhite
        .MouseInterval = 100
        .Direction = [Up_Down]
        .UnRefreshControl = False
    End With
    bNotOk2 = False
    Call UserControl_Paint  'refresh
    Timer1.Enabled = True
End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)

    If bEnable = False Then Exit Sub

    Select Case KeyCode
    Case vbKeyLeft, vbKeyUp
        lValue = lValue - SmallChange
        Call ChangeValues
        RaiseEvent Change(lValue)
    Case vbKeyRight, vbKeyDown
        lValue = lValue + SmallChange
        Call ChangeValues
        RaiseEvent Change(lValue)
    Case vbKeyEnd
        lValue = lMax
        Call ChangeValues
        RaiseEvent Change(lValue)
    Case vbKeyHome
        lValue = lMin
        Call ChangeValues
        RaiseEvent Change(lValue)
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
        Call .WriteProperty("BorderColor", Me.BorderColor, &HFF8080)
        Call .WriteProperty("Enabled", Me.Enabled, True)
        Call .WriteProperty("FrontColor", Me.FrontColor, 15782079)
        Call .WriteProperty("Max", Me.Max, 100)
        Call .WriteProperty("Min", Me.Min, 0)
        Call .WriteProperty("SmallChange", Me.SmallChange, 1)
        Call .WriteProperty("Value", Me.Value, 50)
        Call .WriteProperty("MouseHoverColor", Me.MouseHoverColor, vbWhite)
        Call .WriteProperty("DownColor", Me.DownColor, 12492429)
        Call .WriteProperty("MouseInterval", Me.MouseInterval, 100)
        Call .WriteProperty("Direction", Me.Direction, [Up_Down])
        Call .WriteProperty("UnRefreshControl", Me.UnRefreshControl, False)
    End With
End Sub
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    bNotOk2 = True
    With PropBag
        Me.ArrowColor = .ReadProperty("ArrowColor", vbWhite)
        Me.BorderColor = .ReadProperty("BorderColor", &HFF8080)
        Me.Enabled = .ReadProperty("Enabled", True)
        Me.FrontColor = .ReadProperty("FrontColor", 15782079)
        Me.Max = .ReadProperty("Max", 100)
        Me.Min = .ReadProperty("Min", 0)
        Me.SmallChange = .ReadProperty("SmallChange", 1)
        Me.Value = .ReadProperty("Value", 50)
        Me.MouseHoverColor = .ReadProperty("MouseHoverColor", vbWhite)
        Me.DownColor = .ReadProperty("DownColor", 12492429)
        Me.MouseInterval = .ReadProperty("MouseInterval", 100)
        Me.Direction = .ReadProperty("Direction", [Up_Down])
        Me.UnRefreshControl = .ReadProperty("UnRefreshControl", False)
    End With
    bNotOk2 = False
    'Call UserControl_Paint  'refresh
    Call Refresh

    'le bon endroit pour lancer le subclassing
    Call LaunchKeyMouseEvents
End Sub
Private Sub UserControl_Resize()
    If Height < 255 Then Height = 255
    If Width < 255 Then Width = 255
    Call ChangeValues
    'Call Refresh
End Sub

'=======================================================
'lance le subclassing
'=======================================================
Private Sub LaunchKeyMouseEvents()

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
Public Property Get BorderColor() As ole_color: BorderColor = lBorderColor: End Property
Public Property Let BorderColor(BorderColor As ole_color): lBorderColor = BorderColor: bNotOk = False: UserControl_Paint: End Property
Public Property Get ArrowColor() As ole_color: ArrowColor = lArrowColor: End Property
Attribute ArrowColor.VB_MemberFlags = "40"
Public Property Let ArrowColor(ArrowColor As ole_color): lArrowColor = ArrowColor: bNotOk = False: UserControl_Paint: End Property
Public Property Get FrontColor() As ole_color: FrontColor = lFrontColor: End Property
Public Property Let FrontColor(FrontColor As ole_color): lFrontColor = FrontColor: bNotOk = False: UserControl_Paint: End Property
Public Property Get Enabled() As Boolean: Enabled = bEnable: End Property
Public Property Let Enabled(Enabled As Boolean): bEnable = Enabled: bNotOk = False: UserControl_Paint: End Property
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
    If Max < lValue Then Exit Property
    If Max <> lMax Then
        lMax = Max
        ChangeValues
        bNotOk = False: UserControl_Paint
    End If
End Property
Public Property Get Value() As Currency: Value = lValue: End Property
Public Property Let Value(Value As Currency)
    If Value <> lValue Then
        RaiseEvent Change(lValue)
        lValue = Value: Call ChangeValues
    End If
End Property
Public Property Get SmallChange() As Currency: SmallChange = lSmallChange: End Property
Public Property Let SmallChange(SmallChange As Currency): lSmallChange = SmallChange: bNotOk = False: UserControl_Paint: End Property
Public Property Get DownColor() As ole_color: DownColor = lDownColor: End Property
Public Property Let DownColor(DownColor As ole_color): lDownColor = DownColor: bNotOk = False: UserControl_Paint: End Property
Public Property Get MouseHoverColor() As ole_color: MouseHoverColor = lSelColor: End Property
Public Property Let MouseHoverColor(MouseHoverColor As ole_color): lSelColor = MouseHoverColor: bNotOk = False: UserControl_Paint: End Property
Public Property Get MouseInterval() As Long: MouseInterval = lMouseInterval: End Property
Public Property Let MouseInterval(MouseInterval As Long): lMouseInterval = MouseInterval: Timer1.Interval = lMouseInterval: End Property
Public Property Get Direction() As Direction: Direction = lSens: End Property
Public Property Let Direction(Direction As Direction): lSens = Direction: bNotOk = False: UserControl_Paint: End Property
Public Property Get UnRefreshControl() As Boolean: UnRefreshControl = bUnRefreshControl: End Property
Attribute UnRefreshControl.VB_Description = "Prevent to refresh control"
Public Property Let UnRefreshControl(UnRefreshControl As Boolean): bUnRefreshControl = UnRefreshControl: End Property



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

    'refresh le controle
    bNotOk = False: Call UserControl_Paint
End Sub

'=======================================================
'MAJ du controle
'=======================================================
Public Sub Refresh()
    Dim X As Long
    Dim Y As Long

    If bUnRefreshControl Then Exit Sub

    '//on efface
    Call UserControl.cls

    '//convertit les couleurs
    Call OleTranslateColor(lArrowColor, 0, lArrowColor)
    Call OleTranslateColor(lFrontColor, 0, lFrontColor)
    Call OleTranslateColor(lBorderColor, 0, lBorderColor)


    '//on trace les bords haut et bas, leur bordure et la bordure générale
    'contour général
    Line (0, 0)-(Width, Height), lBorderColor, BF
    'couleur des flèches
    Line (15, 15)-(Width - 30, Height - 30), lFrontColor, BF
    'trace une séparation à la moitié
    If lSens = Up_Down Then
        Line (0, (Height - 15) / 2)-(Width, (Height - 15) / 2), lBorderColor
    Else
        Line ((Width - 15) / 2, 0)-((Width - 15) / 2, Height), lBorderColor
    End If


    '//trace les rectangles de sélection/pushed/rien
    Call DrawSelRectUp
    Call DrawSelRectDown


    '//on trace les flèches
    'si Enabled=false on met la couleur 10070188 aux flèches
    If bEnable Then
        UserControl.ForeColor = lArrowColor
    Else
        UserControl.ForeColor = 10070188
    End If

    If lSens = Up_Down Then
        X = (Width - 255) / 2 + 15
        Y = (Height - 255) / 4
        'flèche du haut
        Line (105 + X, 45 + Y)-(120 + X, 45 + Y)
        Line (90 + X, 60 + Y)-(135 + X, 60 + Y)
        Line (75 + X, 75 + Y)-(150 + X, 75 + Y)
        Line (60 + X, 90 + Y)-(165 + X, 90 + Y)
        'en bas maintenant
        Line (105 + X, Height - 60 - Y)-(120 + X, Height - 60 - Y)
        Line (90 + X, Height - 75 - Y)-(135 + X, Height - 75 - Y)
        Line (75 + X, Height - 90 - Y)-(150 + X, Height - 90 - Y)
        Line (60 + X, Height - 105 - Y)-(165 + X, Height - 105 - Y)
    Else
        X = (Height - 255) / 2 + 15
        Y = (Width - 255) / 4
        'flèche de gauche
        Line (45 + Y, 105 + X)-(45 + Y, 120 + X)
        Line (60 + Y, 90 + X)-(60 + Y, 135 + X)
        Line (75 + Y, 75 + X)-(75 + Y, 150 + X)
        Line (90 + Y, 60 + X)-(90 + Y, 165 + X)
        'à droite
        Line (Width - 60 - Y, 105 + X)-(Width - 60 - Y, 120 + X)
        Line (Width - 75 - Y, 90 + X)-(Width - 75 - Y, 135 + X)
        Line (Width - 90 - Y, 75 + X)-(Width - 90 - Y, 150 + X)
        Line (Width - 105 - Y, 60 + X)-(Width - 105 - Y, 165 + X)
    End If


    If bEnable = False Then
        'boulot terminé !
        Call UserControl.Refresh
        bNotOk = True
        Exit Sub
    End If


    '//on refresh le control
    Call UserControl.Refresh

    bNotOk = True
End Sub

'=======================================================
'trace le rectangle de sélection de la fleche du haut
'=======================================================
Private Sub DrawSelRectUp()
    Dim Y As Long

    'lUpMoused 1 (lignes blanches, survol) 2 (lignes foncées, pushed) 0 (rien)

    If bEnable = False Then Exit Sub

    Call OleTranslateColor(lSelColor, 0, lSelColor)
    Call OleTranslateColor(lDownColor, 0, lDownColor)

    Select Case lUpMoused

    Case 0
        Exit Sub

    Case 1
        'survol

        If lSens = Up_Down Then
            Y = (Height - 45) / 2
            UserControl.ForeColor = lSelColor
            Line (15, 15)-(Width - 15, 15)
            Line (15, 15)-(15, Y)
            Line (15, Y)-(Width - 15, Y)
            Line (Width - 30, Y)-(Width - 30, 15)
        Else
            Y = (Width - 45) / 2
            UserControl.ForeColor = lSelColor
            Line (15, 15)-(15, Height - 15)
            Line (15, 15)-(Y, 15)
            Line (Y, 15)-(Y, Height - 15)
            Line (Y, Height - 30)-(15, Height - 30)
        End If

    Case 2
        'clic

        If lSens = Up_Down Then
            Line (15, 15)-(Width - 30, (Height - 45) / 2), lDownColor, BF
        Else
            Line (15, 15)-((Width - 30) / 2, Height - 30), lDownColor, BF
        End If

    End Select

    Call UserControl.Refresh
End Sub

'=======================================================
'trace le rectangle de sélection de la fleche du bas
'=======================================================
Private Sub DrawSelRectDown()
    Dim Y As Long

    'lDownMoused 1 (lignes blanches, survol) 2 (lignes foncées, pushed) 0 (rien)

    If bEnable = False Then Exit Sub

    Call OleTranslateColor(lSelColor, 0, lSelColor)
    Call OleTranslateColor(lDownColor, 0, lDownColor)

    Select Case lDownMoused

    Case 0
        Exit Sub

    Case 1
        'survol

        If lSens = Up_Down Then
            Y = (Height - 15) / 2
            UserControl.ForeColor = lSelColor
            Line (15, Height - 30)-(Width - 15, Height - 30)
            Line (15, Height - 30)-(15, Height - Y)
            Line (15, Height - Y)-(Width - 15, Height - Y)
            Line (Width - 30, Height - Y)-(Width - 30, Height - 30)
        Else
            Y = (Width - 15) / 2
            UserControl.ForeColor = lSelColor
            Line (Width - 30, 15)-(Width - 30, Height - 15)
            Line (Width - 30, 15)-(Width - Y, 15)
            Line (Width - Y, 15)-(Width - Y, Height - 15)
            Line (Width - Y, Height - 30)-(Width - 30, Height - 30)
        End If

    Case 2
        'clic

        If lSens = Up_Down Then
            Line (15, Height - (Height - 15) / 2)-(Width - 30, Height - 30), lDownColor, BF
        Else
            Line (Width - (Width - 15) / 2, 15)-(Width - 30, Height - 30), lDownColor, BF
        End If

    End Select

End Sub

'=======================================================
'renvoie l'objet extender de ce usercontrol (pour les propertypages)
'=======================================================
Friend Property Get MyExtender() As Object
    Set MyExtender = UserControl.Extender
End Property
Friend Property Let MyExtender(MyExtender As Object)
    Set UserControl.Extender = MyExtender
End Property
