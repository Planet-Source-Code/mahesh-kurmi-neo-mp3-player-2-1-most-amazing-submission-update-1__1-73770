VERSION 5.00
Begin VB.UserControl vkFrame 
   AutoRedraw      =   -1  'True
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ControlContainer=   -1  'True
   PropertyPages   =   "vkFrame.ctx":0000
   ScaleHeight     =   3600
   ScaleWidth      =   4800
   ToolboxBitmap   =   "vkFrame.ctx":004B
   Begin VB.PictureBox pctG 
      BorderStyle     =   0  'None
      Height          =   240
      Left            =   2160
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   0
      Top             =   1080
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image PCTcolor 
      Height          =   240
      Left            =   720
      Top             =   0
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image PCTgray 
      Height          =   240
      Left            =   360
      Top             =   0
      Visible         =   0   'False
      Width           =   240
   End
End
Attribute VB_Name = "vkFrame"
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

Private lBackStyle As BackStyleConstants
Private lTextPos As AlignmentConstants
Private lTitleHeight As Long
Private lForeColor As ole_color
Private tCol1 As ole_color
Private tCol2 As ole_color
Private bCol1 As ole_color
Private bCol2 As ole_color
Private bShowTitle As Boolean
Private sCaption As String
Private bShowBackGround As Boolean
Private lTitleGradient As GradientConstants
Private lBackGradient As GradientConstants
Private bNotOk As Boolean
Private bNotOk2 As Boolean
Private bEnable As Boolean
Private lBorderColor As ole_color
Private bDisplayBorder As Boolean
Private bBreakCorner As Boolean
Private lCornerSize As Long
Private lBWidth As Long
Private pctAlign As PictureAlignment
Private bPic As Boolean
Private lOffsetX As Long
Private lOffsetY As Long
Private bGray As Boolean
Private bUnRefreshControl As Boolean

'=======================================================
'EVENTS
'=======================================================
Public Event KeyDown(KeyCode As Integer, Shift As Integer)
Attribute KeyDown.VB_Description = "Happens when a key is down"
Public Event KeyPress(KeyAscii As Integer)
Attribute KeyPress.VB_Description = "Happens when a key is pressed"
Public Event KeyUp(KeyCode As Integer, Shift As Integer)
Attribute KeyUp.VB_Description = "Happens when a key is up"
Public Event MouseHover()
Attribute MouseHover.VB_Description = "Happens when mouse enters control"
Public Event MouseLeave()
Attribute MouseLeave.VB_Description = "Happens when mouse leaves control"
Public Event MouseWheel(Sens As Wheel_Sens)
Attribute MouseWheel.VB_Description = "Happens when control gets a wheel"
Public Event MouseDown(Button As MouseButtonConstants, Shift As Integer, Control As Integer, X As Long, Y As Long)
Attribute MouseDown.VB_Description = "Happens when control gets a click"
Public Event MouseUp(Button As MouseButtonConstants, Shift As Integer, Control As Integer, X As Long, Y As Long)
Attribute MouseUp.VB_Description = "Happens when control gets a mouseup"
Public Event MouseDblClick(Button As MouseButtonConstants, Shift As Integer, Control As Integer, X As Long, Y As Long)
Attribute MouseDblClick.VB_Description = "Happens when control gets a dblclick"
Public Event MouseMove(Button As MouseButtonConstants, Shift As Integer, Control As Integer, X As Long, Y As Long)
Attribute MouseMove.VB_Description = "Happens when mouse moves on control"




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

        RaiseEvent MouseDblClick(vbLeftButton, iShift, iControl, X, Y)
    Case WM_LBUTTONDOWN
        iShift = Abs((wParam And MK_SHIFT) = MK_SHIFT)
        iControl = Abs((wParam And MK_CONTROL) = MK_CONTROL)
        X = LoWord(lParam) * Screen.TwipsPerPixelX
        Y = HiWord(lParam) * Screen.TwipsPerPixelY

        RaiseEvent MouseDown(vbLeftButton, iShift, iControl, X, Y)
    Case WM_LBUTTONUP
        iShift = Abs((wParam And MK_SHIFT) = MK_SHIFT)
        iControl = Abs((wParam And MK_CONTROL) = MK_CONTROL)
        X = LoWord(lParam) * Screen.TwipsPerPixelX
        Y = HiWord(lParam) * Screen.TwipsPerPixelY

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
    Case WM_MOUSEMOVE
        Call TrackMouseEvent(ET)

        iShift = Abs((wParam And MK_SHIFT) = MK_SHIFT)
        iControl = Abs((wParam And MK_CONTROL) = MK_CONTROL)
        X = LoWord(lParam) * Screen.TwipsPerPixelX
        Y = HiWord(lParam) * Screen.TwipsPerPixelY

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
        .BackColor1 = &HFBFBFB       '
        .BackColor2 = &HDCDCDC    '
        .BackGradient = Horizontal    '
        .BackStyle = Opaque
        .Caption = "Caption"    '
        .Font = Ambient.Font    '
        .ForeColor = vbWhite    '
        .ShowBackGround = True    '
        .ShowTitle = True    '
        .TextPosition = vbCenter    '
        .TitleColor1 = &HC85A21
        .TitleColor2 = &HE4C6B5    '
        .TitleGradient = Vertical    '
        .TitleHeight = 250    '
        .Enabled = True    '
        .BorderColor = &HFF8080    '
        .DisplayBorder = True    '
        .BreakCorner = True    '
        .BorderWidth = 1    '
        .RoundAngle = 7    '
        Set .Picture = Nothing
        .PictureAlignment = [Left Justify]
        .DisplayPicture = True
        .PictureOffsetX = 0
        .PictureOffsetY = 0
        .GrayPictureWhenDisabled = True
        .UnRefreshControl = False
    End With
    bNotOk2 = False
    Call UserControl_Paint  'refresh
End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
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
        Call .WriteProperty("UnRefreshControl", Me.UnRefreshControl, False)
        Call .WriteProperty("BackColor1", Me.BackColor1, &HFBFBFB)
        Call .WriteProperty("BackColor2", Me.BackColor2, &HDCDCDC)
        Call .WriteProperty("BackGradient", Me.BackGradient, Horizontal)
        Call .WriteProperty("BackStyle", Me.BackStyle, Opaque)
        Call .WriteProperty("Caption", Me.Caption, "Caption")
        Call .WriteProperty("Font", Me.Font, Ambient.Font)
        Call .WriteProperty("ForeColor", Me.ForeColor, vbWhite)
        Call .WriteProperty("ShowBackGround", Me.ShowBackGround, True)
        Call .WriteProperty("ShowTitle", Me.ShowTitle, True)
        Call .WriteProperty("TextPosition", Me.TextPosition, vbCenter)
        Call .WriteProperty("TitleColor1", Me.TitleColor1, &HC85A21)
        Call .WriteProperty("TitleColor2", Me.TitleColor2, &HE4C6B5)
        Call .WriteProperty("TitleGradient", Me.TitleGradient, Vertical)
        Call .WriteProperty("TitleHeight", Me.TitleHeight, 250)
        Call .WriteProperty("Enabled", Me.Enabled, True)
        Call .WriteProperty("BorderColor", Me.BorderColor, &HFF8080)
        Call .WriteProperty("DisplayBorder", Me.DisplayBorder, True)
        Call .WriteProperty("BreakCorner", Me.BreakCorner, True)
        Call .WriteProperty("Picture", Me.Picture, Nothing)
        Call .WriteProperty("RoundAngle", Me.RoundAngle, 7)
        Call .WriteProperty("BorderWidth", Me.BorderWidth, 1)
        Call .WriteProperty("PictureAlignment", Me.PictureAlignment, [Left Justify])
        Call .WriteProperty("DisplayPicture", Me.DisplayPicture, True)
        Call .WriteProperty("PictureOffsetX", Me.PictureOffsetX, 0)
        Call .WriteProperty("PictureOffsetY", Me.PictureOffsetY, 0)
        Call .WriteProperty("GrayPictureWhenDisabled", Me.GrayPictureWhenDisabled, True)
    End With
End Sub
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    bNotOk2 = True
    With PropBag
        Me.BackColor1 = .ReadProperty("BackColor1", &HFBFBFB)
        Me.BackColor2 = .ReadProperty("BackColor2", &HDCDCDC)
        Me.BackGradient = .ReadProperty("BackGradient", Horizontal)
        Me.BackStyle = .ReadProperty("BackStyle", Opaque)
        Me.Caption = .ReadProperty("Caption", "Caption")
        Set Me.Font = .ReadProperty("Font", Ambient.Font)
        Me.ForeColor = .ReadProperty("ForeColor", vbWhite)
        Me.ShowBackGround = .ReadProperty("ShowBackGround", True)
        Me.ShowTitle = .ReadProperty("ShowTitle", True)
        Me.TextPosition = .ReadProperty("TextPosition", vbCenter)
        Me.TitleColor1 = .ReadProperty("TitleColor1", &HC85A21)
        Me.TitleColor2 = .ReadProperty("TitleColor2", &HE4C6B5)
        Me.TitleGradient = .ReadProperty("TitleGradient", Vertical)
        Me.TitleHeight = .ReadProperty("TitleHeight", 250)
        Me.Enabled = .ReadProperty("Enabled", True)
        Me.BorderColor = .ReadProperty("BorderColor", &HFF8080)
        Me.DisplayBorder = .ReadProperty("DisplayBorder", True)
        Me.BreakCorner = .ReadProperty("BreakCorner", True)
        Set Me.Picture = .ReadProperty("Picture", Nothing)
        Me.BorderWidth = .ReadProperty("BorderWidth", 1)
        Me.RoundAngle = .ReadProperty("RoundAngle", 7)
        Me.PictureAlignment = .ReadProperty("PictureAlignment", [Left Justify])
        Me.DisplayPicture = .ReadProperty("DisplayPicture", True)
        Me.PictureOffsetX = .ReadProperty("PictureOffsetX", 0)
        Me.PictureOffsetY = .ReadProperty("PictureOffsetY", 0)
        Me.GrayPictureWhenDisabled = .ReadProperty("GrayPictureWhenDisabled", True)
        Me.UnRefreshControl = .ReadProperty("UnRefreshControl", False)
    End With
    bNotOk2 = False
    'Call UserControl_Paint  'refresh
    Call Refresh

    'le bon endroit pour lancer le subclassing
    Call LaunchKeyMouseEvents
End Sub
Private Sub UserControl_Resize()
    bNotOk = False
    Call UserControl_Paint  'refresh
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
Attribute hDC.VB_Description = "Get the control hDc"
Public Property Get hwnd() As Long: hwnd = UserControl.hwnd: End Property
Attribute hwnd.VB_Description = "Handle of the control"
Public Property Get BackStyle() As BackStyleConstants: BackStyle = lBackStyle: End Property
Attribute BackStyle.VB_Description = "Use a transparent control or not"
Public Property Let BackStyle(BackStyle As BackStyleConstants): lBackStyle = BackStyle: UserControl.BackStyle = BackStyle: bNotOk = False: UserControl_Paint: End Property
Public Property Get TextPosition() As AlignmentConstants: TextPosition = lTextPos: End Property
Attribute TextPosition.VB_Description = "Text position"
Public Property Let TextPosition(TextPosition As AlignmentConstants): lTextPos = TextPosition: bNotOk = False: UserControl_Paint: End Property
Public Property Get Caption() As String: Caption = sCaption: End Property
Attribute Caption.VB_Description = "Text to display"
Public Property Let Caption(Caption As String): sCaption = Caption: bNotOk = False: UserControl_Paint: bNotOk = True: End Property
Public Property Get ShowTitle() As Boolean: ShowTitle = bShowTitle: End Property
Attribute ShowTitle.VB_Description = "Display title zone or not"
Public Property Let ShowTitle(ShowTitle As Boolean): bShowTitle = ShowTitle: bNotOk = False: UserControl_Paint: End Property
Public Property Get TitleHeight() As Long: TitleHeight = lTitleHeight: End Property
Attribute TitleHeight.VB_Description = "Height of the title zone"
Public Property Let TitleHeight(TitleHeight As Long): lTitleHeight = TitleHeight: bNotOk = False: UserControl_Paint: End Property
Public Property Get ForeColor() As ole_color: ForeColor = lForeColor: End Property
Attribute ForeColor.VB_Description = "Text color"
Public Property Let ForeColor(ForeColor As ole_color): lForeColor = ForeColor: UserControl.ForeColor = ForeColor: bNotOk = False: UserControl_Paint: End Property
Public Property Get TitleColor1() As ole_color: TitleColor1 = tCol1: End Property
Attribute TitleColor1.VB_Description = "Color1 of the title gradient"
Public Property Let TitleColor1(TitleColor1 As ole_color): tCol1 = TitleColor1: bNotOk = False: UserControl_Paint: End Property
Public Property Get TitleColor2() As ole_color: TitleColor2 = tCol2: End Property
Attribute TitleColor2.VB_Description = "Color2 of the title gradient"
Public Property Let TitleColor2(TitleColor2 As ole_color): tCol2 = TitleColor2: bNotOk = False: UserControl_Paint: End Property
Public Property Get BackColor1() As ole_color: BackColor1 = bCol1: End Property
Attribute BackColor1.VB_Description = "Color1 of the back gradient"
Public Property Let BackColor1(BackColor1 As ole_color): bCol1 = BackColor1: bNotOk = False: UserControl_Paint: End Property
Public Property Get BackColor2() As ole_color: BackColor2 = bCol2: End Property
Attribute BackColor2.VB_Description = "Color2 of the back gradient"
Public Property Let BackColor2(BackColor2 As ole_color): bCol2 = BackColor2: bNotOk = False: UserControl_Paint: End Property
Public Property Get Font() As StdFont: Set Font = UserControl.Font: End Property
Attribute Font.VB_Description = "Text font"
Public Property Set Font(Font As StdFont): Set UserControl.Font = Font: bNotOk = False: UserControl_Paint: End Property
Public Property Get ShowBackGround() As Boolean: ShowBackGround = bShowBackGround: End Property
Attribute ShowBackGround.VB_Description = "Display background or not"
Public Property Let ShowBackGround(ShowBackGround As Boolean): bShowBackGround = ShowBackGround: bNotOk = False: UserControl_Paint: End Property
Public Property Get TitleGradient() As GradientConstants: TitleGradient = lTitleGradient: End Property
Attribute TitleGradient.VB_Description = "Title gradient type"
Public Property Let TitleGradient(TitleGradient As GradientConstants): lTitleGradient = TitleGradient: bNotOk = False: UserControl_Paint: End Property
Public Property Get BackGradient() As GradientConstants: BackGradient = lBackGradient: End Property
Attribute BackGradient.VB_Description = "Type of back gradient"
Public Property Let BackGradient(BackGradient As GradientConstants): lBackGradient = BackGradient: bNotOk = False: UserControl_Paint: End Property
Public Property Get Enabled() As Boolean: Enabled = bEnable: End Property
Attribute Enabled.VB_Description = "Enable control or not"
Public Property Let Enabled(Enabled As Boolean)
    bEnable = Enabled: bNotOk = False: UserControl_Paint: EnableControls
End Property
Public Property Get BorderColor() As ole_color: BorderColor = lBorderColor: End Property
Attribute BorderColor.VB_Description = "Color of the border"
Public Property Let BorderColor(BorderColor As ole_color): lBorderColor = BorderColor: bNotOk = False: UserControl_Paint: End Property
Public Property Get DisplayBorder() As Boolean: DisplayBorder = bDisplayBorder: End Property
Attribute DisplayBorder.VB_Description = "Display border or not"
Public Property Let DisplayBorder(DisplayBorder As Boolean): bDisplayBorder = DisplayBorder: bNotOk = False: UserControl_Paint: End Property
Public Property Get BreakCorner() As Boolean: BreakCorner = bBreakCorner: End Property
Attribute BreakCorner.VB_Description = "Use rounded corner or not"
Public Property Let BreakCorner(BreakCorner As Boolean): bBreakCorner = BreakCorner: bNotOk = False: UserControl_Paint: End Property
Public Property Get Picture() As Picture: Set Picture = PCTcolor.Picture: End Property
Attribute Picture.VB_Description = "Picture to display on the title zone"
Public Property Set Picture(NewPic As Picture)
    Set PCTcolor.Picture = NewPic
    Set pctG.Picture = NewPic
    If Not (NewPic Is Nothing) Then
        pctG.Width = PCTcolor.Width
        pctG.Height = PCTcolor.Height
        Call GrayScale(pctG)
    End If
    PCTgray.Picture = pctG.Image
    bNotOk = False: UserControl_Paint
End Property
Public Property Get RoundAngle() As Long: RoundAngle = lCornerSize: End Property
Attribute RoundAngle.VB_Description = "Use rounded corners or not"
Public Property Let RoundAngle(RoundAngle As Long): lCornerSize = RoundAngle: bNotOk = False: UserControl_Paint: End Property
Public Property Get BorderWidth() As Long: BorderWidth = lBWidth: End Property
Attribute BorderWidth.VB_Description = "Width of the border"
Public Property Let BorderWidth(BorderWidth As Long): lBWidth = BorderWidth: bNotOk = False: UserControl_Paint: End Property
Public Property Get PictureAlignment() As PictureAlignment: PictureAlignment = pctAlign: End Property
Attribute PictureAlignment.VB_Description = "Picture alignment"
Public Property Let PictureAlignment(PictureAlignment As PictureAlignment): pctAlign = PictureAlignment: bNotOk = False: UserControl_Paint: End Property
Public Property Get DisplayPicture() As Boolean: DisplayPicture = bPic: End Property
Attribute DisplayPicture.VB_Description = "Display picture or not"
Public Property Let DisplayPicture(DisplayPicture As Boolean): bPic = DisplayPicture: bNotOk = False: UserControl_Paint: End Property
Public Property Get PictureOffsetX() As Long: PictureOffsetX = lOffsetX: End Property
Attribute PictureOffsetX.VB_Description = "Offset (twips) of the picture"
Public Property Let PictureOffsetX(PictureOffsetX As Long): lOffsetX = PictureOffsetX: bNotOk = False: UserControl_Paint: End Property
Public Property Get PictureOffsetY() As Long: PictureOffsetY = lOffsetY: End Property
Attribute PictureOffsetY.VB_Description = "Offset (twips) of the picture"
Public Property Let PictureOffsetY(PictureOffsetY As Long): lOffsetY = PictureOffsetY: bNotOk = False: UserControl_Paint: End Property
Public Property Get GrayPictureWhenDisabled() As Boolean: GrayPictureWhenDisabled = bGray: End Property
Attribute GrayPictureWhenDisabled.VB_Description = "Use a gray (or color) picture when control is not enabled"
Public Property Let GrayPictureWhenDisabled(GrayPictureWhenDisabled As Boolean): bGray = GrayPictureWhenDisabled: bNotOk = False: UserControl_Paint: End Property
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
'applique un gradient de couleur sur un objet de gauche à droite
'il doit être "en autoredraw=true" (si c'est une form, picturebox...)
'=======================================================
Private Sub FillGradientW(LeftColor As RGB_COLOR, _
                          RightColor As RGB_COLOR, ByVal Width As Long, ByVal Height As Long, _
                          Optional ByVal Dep As Long)

    Dim rAverageColorPerSizeUnit As Double
    Dim gAverageColorPerSizeUnit As Double
    Dim bAverageColorPerSizeUnit As Double
    Dim lWidth As Long
    Dim X As Long
    Dim lHeight As Long
    Dim lSigne As Long

    With UserControl

        'récupère la largeur de l'objet
        lWidth = Width / Screen.TwipsPerPixelX
        lHeight = Height / Screen.TwipsPerPixelY

        'récupère la moyenne de couleur par unité de longueur
        rAverageColorPerSizeUnit = Abs((RightColor.R - LeftColor.R) / lWidth)
        gAverageColorPerSizeUnit = Abs((RightColor.G - LeftColor.G) / lWidth)
        bAverageColorPerSizeUnit = Abs((RightColor.B - LeftColor.B) / lWidth)

        'on change le signe (sens) au cas où
        If CLng(RGB(LeftColor.R, LeftColor.G, LeftColor.B)) <= _
           CLng(RGB(RightColor.R, RightColor.G, RightColor.B)) Then

            lSigne = 1
        Else
            lSigne = -1
        End If

        'se positionne tout à gauche de l'objet ==> balayera vers la droite
        Call MoveToEx(.hDC, 0, Dep, 0&)

        'pour chaque 'colonne' constituée par une ligne verticale, on trace une
        'ligne en récupérant la couleur correspondante
        For X = 0 To lWidth

            'change le ForeColor qui détermine la couleur de la Line
            'multiplie la largeur actuelle par la couleur par unité de longueur
            .ForeColor = RGB(LeftColor.R + X * rAverageColorPerSizeUnit * lSigne, LeftColor.G + X * _
                                                                                  gAverageColorPerSizeUnit * lSigne, LeftColor.B + X * bAverageColorPerSizeUnit * lSigne)

            'trace une ligne
            Call LineTo(.hDC, X, lHeight)

            'bouge 'd'une colonne' vers la droite
            Call MoveToEx(.hDC, X, Dep, 0&)

        Next X

        'on refresh l'objet
        Call .Refresh
    End With

End Sub

'=======================================================
'applique un gradient de couleur sur un objet de gauche à droite
'il doit être "en autoredraw=true" (si c'est une form, picturebox...)
'=======================================================
Private Sub FillGradientH(LeftColor As RGB_COLOR, _
                          RightColor As RGB_COLOR, ByVal Width As Long, ByVal Height As Long, _
                          Optional ByVal Dep As Long)

    Dim rAverageColorPerSizeUnit As Double
    Dim gAverageColorPerSizeUnit As Double
    Dim bAverageColorPerSizeUnit As Double
    Dim lHeight As Long
    Dim X As Long
    Dim lSigne As Long

    With UserControl

        'récupère la hateur de l'objet
        lHeight = Height / Screen.TwipsPerPixelY

        'récupère la moyenne de couleur par unité de longueur
        rAverageColorPerSizeUnit = Abs((RightColor.R - LeftColor.R) / lHeight)
        gAverageColorPerSizeUnit = Abs((RightColor.G - LeftColor.G) / lHeight)
        bAverageColorPerSizeUnit = Abs((RightColor.B - LeftColor.B) / lHeight)

        'on change le signe (sens) au cas où
        If CLng(RGB(LeftColor.R, LeftColor.G, LeftColor.B)) <= _
           CLng(RGB(RightColor.R, RightColor.G, RightColor.B)) Then

            lSigne = 1
        Else
            lSigne = -1
        End If

        'se positionne tout à gauche de l'objet ==> balayera vers le bas
        Call MoveToEx(.hDC, 0, Dep, 0&)

        'pour chaque 'colonne' constituée par une ligne verticale, on trace une
        'ligne en récupérant la couleur correspondante
        For X = Dep To lHeight

            'change le ForeColor qui détermine la couleur de la Line
            'multiplie la largeur actuelle par la couleur par unité de longueur
            .ForeColor = RGB(LeftColor.R + X * rAverageColorPerSizeUnit * lSigne, LeftColor.G + X * _
                                                                                  gAverageColorPerSizeUnit * lSigne, LeftColor.B + X * bAverageColorPerSizeUnit * lSigne)

            'trace une ligne
            Call LineTo(.hDC, Width, X)

            'bouge 'd'une colonne' vers la droite
            Call MoveToEx(.hDC, 0, X, 0&)

        Next X

        'on refresh l'objet
        Call .Refresh
    End With

End Sub

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
'convertit une couleur en long vers RGB
'=======================================================
Private Sub ToRGB(ByVal Color As Long, ByRef RGB As RGB_COLOR)
    With RGB
        .R = Color And &HFF&
        .G = (Color And &HFF00&) \ &H100&
        .B = Color \ &H10000
    End With
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
'PUBLIC SUB
'=======================================================

'=======================================================
'enable ou pas les controles qui sont dans le frame
'=======================================================
Private Sub EnableControls()
    Dim ctr As Control

    On Error Resume Next

    For Each ctr In UserControl.ContainedControls
        ctr.Enabled = bEnable
    Next ctr

End Sub

'=======================================================
'on dessine tout
'=======================================================
Public Sub Refresh()
Attribute Refresh.VB_Description = "Refresh control"
    Dim X As Long
    Dim RGB1 As RGB_COLOR
    Dim RGB2 As RGB_COLOR
    Dim Dep As Long
    Dim R As RECT
    Dim hBrush As Long
    Dim Rec As Long
    Dim hRgn As Long
    Dim W As Long
    Dim H As Long

    If bUnRefreshControl Then Exit Sub

    '//on efface et on vire le maskpicture
    Call UserControl.cls
    UserControl.Picture = Nothing
    UserControl.MaskPicture = Nothing

    '//on convertir les différentes couleurs si couleurs système
    Call OleTranslateColor(lBorderColor, 0, lBorderColor)
    Call OleTranslateColor(lForeColor, 0, lForeColor)
    Call OleTranslateColor(tCol1, 0, tCol1)
    Call OleTranslateColor(tCol2, 0, tCol2)
    Call OleTranslateColor(bCol1, 0, bCol1)
    Call OleTranslateColor(bCol2, 0, bCol2)


    '//on commence par créer le Title
    If bShowTitle Then

        '        If lBackStyle = Transparent And bBreakCorner = True Then
        '            'alors c'est transparent et avec arrondi
        '            'alors on créé la zone perso dès maintenant
        '            'pour ne pas pouvoir tracer le titre dans les coins
        '            'haut-gauche et haut-droite
        '
        '            'on définit une zone rectangulaire à bords arrondi
        '            hRgn = CreateRoundRectRgn(0, 0, scalewidth / 15, _
                     '                scaleheight / 15, lCornerSize, lCornerSize)
        '
        '            'on défini la zone rectangulaire arrondi comme nouvelle fenêtre
        '            Call SetWindowRgn(UserControl.hWnd, hRgn, True)
        '
        '            'on détruit le brush et la zone
        '            Call DeleteObject(hRgn)
        '        End If


        'récupère les 3 composantes des deux couleurs
        Call ToRGB(tCol1, RGB1)
        Call ToRGB(tCol2, RGB2)

        If lTitleGradient = None Then
            'pas de gradient
            'on dessine alors un rectangle
            Line (0, 0)-(ScaleWidth, lTitleHeight), _
                 tCol1, BF
        ElseIf lTitleGradient = Horizontal Then
            'gradient horizontal
            Call FillGradientH(RGB1, RGB2, ScaleWidth, lTitleHeight)
        Else
            'gradient vertical
            Call FillGradientW(RGB1, RGB2, ScaleWidth, lTitleHeight)
        End If
    End If

    '//créé un rectangle
    Call SetRect(R, 0, (lTitleHeight / Screen.TwipsPerPixelY - GetCharHeight) / 2, Width / Screen.TwipsPerPixelX, lTitleHeight / Screen.TwipsPerPixelY)

    'on affiche le caption
    UserControl.ForeColor = lForeColor
    If lTextPos = vbCenter Then
        'au centre
        Call DrawText(UserControl.hDC, sCaption, Len(sCaption), R, DT_CENTER)
    ElseIf lTextPos = vbRightJustify Then
        'à droite
        Call DrawText(UserControl.hDC, sCaption, Len(sCaption), R, DT_RIGHT)
    Else
        'à gauche
        Call DrawText(UserControl.hDC, sCaption, Len(sCaption), R, DT_LEFT)
    End If

    If lBackStyle = TRANSPARENT Then
        'alors on créé une image dans le maskpicture

        'on dessine donc maintenant le bord
        If bDisplayBorder Then
            If bBreakCorner = False Then
                'alors c'est un rectangle

                'on défini un brush
                hBrush = CreateSolidBrush(lBorderColor)

                'on définit une zone rectangulaire à bords arrondi
                hRgn = CreateRectRgn(0, 0, ScaleWidth / Screen.TwipsPerPixelX, _
                                     ScaleHeight / Screen.TwipsPerPixelY)

                'on dessine le contour
                Call FrameRgn(UserControl.hDC, hRgn, hBrush, lBWidth, lBWidth)

                'on détruit le brush et la zone
                Call DeleteObject(hBrush)
                Call DeleteObject(hRgn)

            Else
                'alors c'est un arrondi

                'on défini un brush
                hBrush = CreateSolidBrush(lBorderColor)

                'on définit une zone rectangulaire à bords arrondi
                hRgn = CreateRoundRectRgn(0, 0, ScaleWidth / Screen.TwipsPerPixelX, _
                                          ScaleHeight / Screen.TwipsPerPixelY, lCornerSize, lCornerSize)

                'on dessine le contour
                Call FrameRgn(UserControl.hDC, hRgn, hBrush, lBWidth, lBWidth)

                'on défini la zone rectangulaire arrondi comme nouvelle fenêtre
                Call SetWindowRgn(UserControl.hwnd, hRgn, True)

                'on détruit le brush et la zone
                Call DeleteObject(hBrush)
                Call DeleteObject(hRgn)

            End If
        End If

        UserControl.MaskPicture = UserControl.Image
    End If

    '//créé le gradient du reste du contrôle
    If lBackStyle = Opaque Then

        UserControl.BackStyle = 1

        If bShowBackGround Then

            'récupère les 3 composantes des deux couleurs
            Call ToRGB(bCol1, RGB1)
            Call ToRGB(bCol2, RGB2)

            If bShowTitle Then
                Dep = lTitleHeight / Screen.TwipsPerPixelY    'commence le gradient de pas tout en haut
            Else
                Dep = 0    'commence de tout en haut le gradient
            End If

            If lBackGradient = None Then
                'pas de gradient
                'on dessine alors un rectangle
                Line (15, lTitleHeight + 1)-(ScaleWidth - 30, _
                                             ScaleHeight - 30), bCol1, BF
            ElseIf lBackGradient = Horizontal Then
                'gradient horizontal
                Call FillGradientH(RGB1, RGB2, ScaleWidth, _
                                   ScaleHeight, Dep)
            Else
                'gradient vertical
                Call FillGradientW(RGB1, RGB2, ScaleWidth, _
                                   ScaleHeight, Dep)
            End If
        End If
        '    Else
        '        'on récupère le maskpicture et on met le controle en transparent
        '
        '        UserControl.BackStyle = 0
        '        UserControl.Picture = UserControl.MaskPicture
        '
    End If



    '//on va se tracer la bitmap maintenant ^^
    If bPic Then
        'on montre l'image présente

        'on choisit celle à afficher (Grise ou pas)
        If bEnable Or Not (bGray) And bShowTitle And PCTcolor.Picture Then
            With PCTcolor
                'on move IMG en fonction du choix de Aligment
                'on centre l'image dans la barre de titre
                If pctAlign = [Left Justify] Then
                    'à gauche

                    W = (lTitleHeight - .Height) / 2 + lOffsetY
                    H = 30 + lCornerSize * 2 + lOffsetX
                    If W < 0 Then W = 0
                    If H < 0 Then H = 0

                    .Top = W
                    .Left = H    '2 pxls + arrondi

                Else
                    'à droite

                    W = (lTitleHeight - .Height) / 2 + lOffsetY
                    H = Width - 30 - .Width - lCornerSize * 2 - lOffsetX
                    If H < 0 Then H = 0
                    If W < 0 Then W = 0

                    .Top = W
                    .Left = H    '2 pxls + arrondi

                End If

                .Visible = True
                PCTgray.Visible = False
            End With
        ElseIf bShowTitle And PCTcolor.Picture Then
            With PCTgray
                'on move IMG en fonction du choix de Aligment
                'on centre l'image dans la barre de titre
                If pctAlign = [Left Justify] Then
                    'à gauche

                    W = (lTitleHeight - .Height) / 2 + lOffsetY
                    H = 30 + lCornerSize * 2 + lOffsetX
                    If W < 0 Then W = 0
                    If H < 0 Then H = 0

                    .Top = W
                    .Left = H    '2 pxls + arrondi

                Else
                    'à droite

                    W = (lTitleHeight - .Height) / 2 + lOffsetY
                    H = Width - 30 - .Width - lCornerSize * 2 - lOffsetX
                    If H < 0 Then H = 0
                    If W < 0 Then W = 0

                    .Top = W
                    .Left = H    '2 pxls + arrondi

                End If

                .Visible = True
                PCTcolor.Visible = False
            End With
        End If
    Else
        'on masque l'image
        PCTgray.Visible = False
        PCTcolor.Visible = False
    End If


    '//on trace la bordure si opaque
    If lBackStyle = Opaque And bDisplayBorder Then

        If bBreakCorner = False Then
            'alors c'est un rectangle

            'on défini un brush
            hBrush = CreateSolidBrush(lBorderColor)

            'on définit une zone rectangulaire à bords arrondi
            hRgn = CreateRectRgn(0, 0, ScaleWidth / Screen.TwipsPerPixelX, _
                                 ScaleHeight / Screen.TwipsPerPixelY)

            'on dessine le contour
            Call FrameRgn(UserControl.hDC, hRgn, hBrush, lBWidth, lBWidth)

            'on détruit le brush et la zone
            Call DeleteObject(hBrush)
            Call DeleteObject(hRgn)

        Else
            'alors c'est un arrondi

            'on défini un brush
            hBrush = CreateSolidBrush(lBorderColor)

            'on définit une zone rectangulaire à bords arrondi
            hRgn = CreateRoundRectRgn(0, 0, ScaleWidth / Screen.TwipsPerPixelX, _
                                      ScaleHeight / Screen.TwipsPerPixelY, lCornerSize, lCornerSize)

            'on dessine le contour
            Call FrameRgn(UserControl.hDC, hRgn, hBrush, lBWidth, lBWidth)

            'on défini la zone rectangulaire arrondi comme nouvelle fenêtre
            Call SetWindowRgn(UserControl.hwnd, hRgn, True)

            'on détruit le brush et la zone
            Call DeleteObject(hBrush)
            Call DeleteObject(hRgn)

        End If

    End If

    bNotOk = True

    '//on délocke le controle --> a évité les clignotements
    ' Call LockWindowUpdate(0)

    'permet de refresh la bordure
    Call UserControl.Refresh
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
