VERSION 5.00
Begin VB.Form FrmVisualisation 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   ClientHeight    =   3840
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3885
   ControlBox      =   0   'False
   DrawWidth       =   2
   Icon            =   "visualisation.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   256
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   259
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin MMPlayerXProject.ScrollText ScrollText1 
      Height          =   90
      Left            =   960
      TabIndex        =   3
      ToolTipText     =   "mahesh"
      Top             =   3120
      Width           =   1725
      _extentx        =   3043
      _extenty        =   159
      aligntext       =   2
      scrollvelocity  =   150
      captiontext     =   "---MAHESH MP3 PLAYER---"
      scroll          =   -1  'True
      scrolltype      =   1
   End
   Begin VB.PictureBox ScopeBuff 
      AutoRedraw      =   -1  'True
      BackColor       =   &H000080FF&
      BorderStyle     =   0  'None
      ForeColor       =   &H00C0E0FF&
      Height          =   1170
      Left            =   900
      ScaleHeight     =   78
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   132
      TabIndex        =   1
      Top             =   1170
      Visible         =   0   'False
      Width           =   1980
   End
   Begin MMPlayerXProject.Button btnExit 
      Height          =   120
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   225
      _extentx        =   397
      _extenty        =   212
      style           =   1
      buttoncolor     =   12632256
      mousepointer    =   99
   End
   Begin VB.PictureBox PicSpectrum 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      FillColor       =   &H00808000&
      FillStyle       =   0  'Solid
      ForeColor       =   &H000000FF&
      Height          =   2940
      Left            =   360
      ScaleHeight     =   196
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   206
      TabIndex        =   0
      Top             =   300
      Width           =   3090
   End
End
Attribute VB_Name = "FrmVisualisation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'////////////////////////


Private Declare Function SetPixel Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByVal crColor As Long) As Long

Private NewX As Single, NewY As Single
Attribute NewY.VB_VarUserMemId = 1073938432
Private Col As Long
Attribute Col.VB_VarUserMemId = 1073938434

Private tp_wave_r As Single, tp_wave_g As Single, tp_wave_b As Single
Attribute tp_wave_r.VB_VarUserMemId = 1073938435
Attribute tp_wave_g.VB_VarUserMemId = 1073938435
Attribute tp_wave_b.VB_VarUserMemId = 1073938435
Private Angle As Single, Step As Integer, Temp As Single
Attribute Angle.VB_VarUserMemId = 1073938439
Attribute Step.VB_VarUserMemId = 1073938439
Attribute Temp.VB_VarUserMemId = 1073938439


Private Declare Function LineTo Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function MoveToEx Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, lpPoint As POINTAPI) As Long
Private Type POINTAPI
    X As Long
    Y As Long
End Type

Private R1 As Single, R2 As Single
Attribute R1.VB_VarUserMemId = 1073938442
Attribute R2.VB_VarUserMemId = 1073938442
Private g1 As Single, g2 As Single
Attribute g1.VB_VarUserMemId = 1073938444
Attribute g2.VB_VarUserMemId = 1073938444
Private b1 As Single, b2 As Single
Attribute b1.VB_VarUserMemId = 1073938446
Attribute b2.VB_VarUserMemId = 1073938446

Private Declare Function CreatePen Lib "gdi32" (ByVal nPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long

Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type
'////////////////////////////
Dim lSpectrumColor As Long
Dim lShadowcolor As Long
Dim InFormDrag As Boolean
Attribute InFormDrag.VB_VarUserMemId = 1073938452
Dim cWindows As New cWindowSkin
Attribute cWindows.VB_VarUserMemId = 1073938453
Dim cAjustarDesk As New clsDockingHandler
Attribute cAjustarDesk.VB_VarUserMemId = 1073938454
Dim exitRightOffset As Integer
Attribute exitRightOffset.VB_VarUserMemId = 1073938457
Dim iScrolltopOffset As Integer
Attribute iScrolltopOffset.VB_VarUserMemId = 1073938458
Dim iScrollRightOffset As Integer
Attribute iScrollRightOffset.VB_VarUserMemId = 1073938459
Dim iScrollleftOffset As Integer
Attribute iScrollleftOffset.VB_VarUserMemId = 1073938460
Dim Space As Integer
Attribute Space.VB_VarUserMemId = 1073938462
Dim Barwidth As Integer
Attribute Barwidth.VB_VarUserMemId = 1073938463
Dim Yspace As Integer
Attribute Yspace.VB_VarUserMemId = 1073938464
Dim Ywidth As Integer
Attribute Ywidth.VB_VarUserMemId = 1073938465
Dim nSample As Integer
Attribute nSample.VB_VarUserMemId = 1073938466

Dim Bandwidth As Integer
Attribute Bandwidth.VB_VarUserMemId = 1073938467

Private Sub btnExit_Click()
    boolVisShow = False
    'boolVisLoaded = False
    frmPopUp.mnuShowVis.Checked = False
    Call CheckMenu(12, frmPopUp.mnuShowVis.Checked)
    Me.Hide
    'Unload Me
End Sub

'32 bit - Add 4 to counter
'24 bit - Add 3 to counter
Private Sub Form_Load()
    Const flag As Long = SWP_NOMOVE Or SWP_NOSIZE
    Dim strRes
    strRes = Read_INI("Configuration", "VX", 6225, , True)
    If IsNumeric(strRes) = False Then strRes = 0
    FrmVisualisation.Left = CInt(strRes)

    strRes = Read_INI("Configuration", "VY", 5815, , True)
    If IsNumeric(strRes) = False Then strRes = 0
    FrmVisualisation.Top = CInt(strRes)

    strRes = Read_INI("Configuration", "VW", 5250, , True)
    FrmVisualisation.Width = CInt(strRes)

    strRes = Read_INI("Configuration", "VH", 2760, , True)
    FrmVisualisation.Height = CInt(strRes)
    LoadSkin
    boolVisLoaded = True
    If OpcionesMusic.SiempreTop = True Then
        SetWindowPos FrmVisualisation.hwnd, HWND_TOPMOST, 0, 0, 0, 0, flag
    Else
        SetWindowPos FrmVisualisation.hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, flag
    End If


    On Error GoTo No_PI
    Dim T As TextStream
    If Fsys.FileExists(App.Path & "\Sodas.list") Then
        Set T = Fsys.OpenTextFile(App.Path & "\Sodas.list", ForReading, True)
        While Not T.AtEndOfStream
            modPi.importXPIList T.readLine
        Wend
        T.Close
        Set T = Nothing
    End If

    setupVisualization
    'Set oPlugIn = CreateObject(GetSetting(App.EXEName, "Visualization", "Object"))

    Exit Sub
No_PI:
    'Debug.Print "Failed to load plugin"

End Sub
Public Sub setupVisualization()
    ScopeBuff.ScaleWidth = PicSpectrum.Width
    ScopeBuff.ScaleHeight = PicSpectrum.Height

    ScopeBuff.BackColor = PicSpectrum.BackColor
    Dim mmm As String
    mmm = GetSetting(App.EXEName, "Visualization", "Object")
    On Error GoTo No_PI
    Set oPlugIn = CreateObject(GetSetting(App.EXEName, "Visualization", "Object"))
    frmPopUp.mnuConfig.Enabled = CBool(GetSetting(App.EXEName, "Visualization", "Config"))
    'tmrVisUpdate.Enabled = True
    'picVisual.ToolTipText = GetSetting(App.EXEName, "Visualization", "Name")
    boolOpluginLoaded = True
    Exit Sub

No_PI:
    'Debug.Print "Falied to load plugin" + err.Description
    boolOpluginLoaded = False
    err.Clear
    'frmDummy.mnuLoadPi_Click
    DoStop
    'setupVisualization
End Sub
Public Sub LoadSkin()
'On Error Resume Next

    Me.Visible = False
    Set cWindows.FormularioPadre = Me
    Set cAjustarDesk.ParentForm = Me
    cWindows.ColorInvisible = Read_INI("NormalMode", "ColorTrans", RGB(255, 0, 255), True)
    cWindows.MinimoAlto = Read_INI("Configuration", "MinHeight", 10, False, False, tAppConfig.AppConfig & "Skins\" & tAppConfig.Skin & "\VIS_STUDIO\config.ini")
    cWindows.MinimoAncho = Read_INI("Configuration", "MinWidth", 10, False, False, tAppConfig.AppConfig & "Skins\" & tAppConfig.Skin & "\VIS_STUDIO\config.ini")
    'iButtonsLeft = Read_INI("Configuration", "ButtonsLeft", 5, False, False, tAppConfig.AppConfig & "Skins\" & tAppConfig.Skin & "\PLAY_LIST\config.ini")
    'iButtonsTop = Read_INI("Configuration", "ButtonsTop", 10, False, False, tAppConfig.AppConfig & "Skins\" & tAppConfig.Skin & "\PLAY_LIST\config.ini")
    cWindows.CargarSkin tAppConfig.AppConfig & "Skins\" & tAppConfig.Skin & "\VIS_STUDIO\", True
    Dim k
    k = Read_Config_Button(btnExit, "Configuration", "exitButton", "0,0,10,10", tAppConfig.AppConfig & "Skins\" & tAppConfig.Skin & "\VIS_STUDIO\config.ini")
    PicSpectrum.BackColor = cRead_INI("Configuration", "SpectrumColor", RGB(0, 0, 0), tAppConfig.AppConfig & "Skins\" & tAppConfig.Skin & "\VIS_STUDIO\config.ini")
    Set btnExit.PictureNormal = LoadPicture(tAppConfig.AppConfig & "Skins\" & tAppConfig.Skin & "\VIS_STUDIO\Exitnormal.bmp")
    Set btnExit.PictureOver = LoadPicture(tAppConfig.AppConfig & "Skins\" & tAppConfig.Skin & "\VIS_STUDIO\Exitover.bmp")
    Set btnExit.PictureDown = LoadPicture(tAppConfig.AppConfig & "Skins\" & tAppConfig.Skin & "\VIS_STUDIO\ExitDown.bmp")
    'Set btnExit.PictureNormal = LoadPicture(tAppConfig.AppConfig & "Skins\" & tAppConfig.Skin & "\DSP\Exitnormal.bmp")
    exitRightOffset = Read_INI("Configuration", "exitRightOffset", 25, False, False, tAppConfig.AppConfig & "Skins\" & tAppConfig.Skin & "\VIS_STUDIO\config.ini")
    btnExit.Left = Me.ScaleWidth - exitRightOffset

    PicSpectrum.Move cWindows.AreaLeft, cWindows.AreaTop, cWindows.AreaWidth, cWindows.AreaHeight
    ScopeBuff.Height = PicSpectrum.Height
    ScopeBuff.Width = PicSpectrum.Width

    visDRW_BARWIDTH = CInt(0.82 * PicSpectrum.ScaleWidth / FFT_BANDS)    ' - 2 * FFT_BANDS
    'DRW_BARSPACE = visDRW_BARWIDTH / 5#
    iScrolltopOffset = Read_INI("Scroll", "TopOffset", 10, False, False, tAppConfig.AppConfig & "Skins\" & tAppConfig.Skin & "\VIS_STUDIO\config.ini")
    iScrollRightOffset = Read_INI("Scroll", "RightOffset", 10, False, False, tAppConfig.AppConfig & "Skins\" & tAppConfig.Skin & "\VIS_STUDIO\config.ini")
    iScrollleftOffset = Read_INI("Scroll", "LeftOffset", 10, False, False, tAppConfig.AppConfig & "Skins\" & tAppConfig.Skin & "\VIS_STUDIO\config.ini")
    Set ScrollText1.PictureText = frmMain.ScrollText(5).PictureText
    ScrollText1.Left = iScrollleftOffset
    ScrollText1.Top = Me.ScaleHeight - iScrolltopOffset
    ScrollText1.Width = Me.ScaleWidth - iScrollRightOffset - iScrollleftOffset
    If sTextScroll <> "" Then ScrollText1.CaptionText = sTextScroll
    Dim i, j, m
    Dim dummyPt As POINTAPI
    Dim ppen
    Space = 1
    Barwidth = 2
    Ywidth = 1
    Yspace = 1
    Dim hDC
    'Exit Sub
    Exit Sub
    PicSpectrum.Picture = LoadPicture()

    hDC = PicSpectrum.hDC
    ppen = CreatePen(0, 1, RGB(10, 58, 78))
    SelectObject hDC, ppen

    For m = 0 To (1224 / (Barwidth + Space)) + 1
        For i = 0 To (800 / (Ywidth + Yspace)) + 1
            For j = 0 To Ywidth - 1
                MoveToEx hDC, m * (Barwidth + Space), i * (Ywidth + Yspace) + j, dummyPt
                LineTo hDC, m * (Barwidth + Space) + Barwidth, i * (Ywidth + Yspace) + j
            Next
        Next
    Next
    PicSpectrum.Picture = PicSpectrum.Image
    SelectObject hDC, -1
    DeleteObject ppen
End Sub


Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    cWindows.Formulario_Down X, Y
    cAjustarDesk.StartDockDrag X * Screen.TwipsPerPixelX, _
                               Y * Screen.TwipsPerPixelY
    InFormDrag = True

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    cWindows.Formulario_MouseMove Button, X, Y
    If InFormDrag And cWindows.Ajustando = False Then
        ' Continue window draggin'
        cAjustarDesk.UpdateDockDrag X * Screen.TwipsPerPixelX, _
                                    Y * Screen.TwipsPerPixelY
        'bHookForm = False
        Exit Sub
    End If

    If cWindows.Ajustando = True Then
        If cGetInputState() <> 0 Then DoEvents

        PicSpectrum.Move cWindows.AreaLeft, cWindows.AreaTop, cWindows.AreaWidth, cWindows.AreaHeight
        ScopeBuff.Width = PicSpectrum.Width
        ScopeBuff.Height = PicSpectrum.Height
        btnExit.Left = Me.ScaleWidth - exitRightOffset
        ScrollText1.Top = Me.ScaleHeight - iScrolltopOffset
        ScrollText1.Width = Me.ScaleWidth - iScrollRightOffset - iScrollleftOffset
    End If

End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    cWindows.Formulario_MouseUp X, Y

    InFormDrag = False

End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim ArchivoINI As String
    Dim INIcheck As String
    Dim Fnum As Integer
    ArchivoINI = tAppConfig.AppPath & App.EXEName & ".ini"

    '// Check attributes
    INIcheck = Dir(ArchivoINI, vbNormal + vbSystem + vbHidden + vbReadOnly + vbArchive)

    '// If you are not doing so...
    If INIcheck = "" Then
        Fnum = FreeFile  '// random number to assign to the file
        Open ArchivoINI For Output As Fnum
        Close
        'SetAttr ArchivoINI, vbHidden + vbSystem
    End If
    Write_INI "Configuration", "VX", FrmVisualisation.Left, ArchivoINI
    Write_INI "Configuration", "VY", FrmVisualisation.Top, ArchivoINI

    Write_INI "Configuration", "VW", FrmVisualisation.Width, ArchivoINI
    Write_INI "Configuration", "VH", FrmVisualisation.Height, ArchivoINI
    boolVisShow = False
    boolVisLoaded = False

End Sub

Private Sub PicSpectrum_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then frmPopUp.PopupMenu frmPopUp.mnuVis
    If Button = 1 Then
        ReleaseCapture

        Call SendMessage(Me.hwnd, &HA1, 2, 0&)
    End If
End Sub


Private Sub ScrollText1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Screen.MousePointer = vbDefault
End Sub

Public Sub drawVis(hDC As Long, DrawData() As Integer, Height As Single, Width As Single)
    Static X As Integer, Y As Single
    Dim dummyPt As POINTAPI
    Dim ppen As Long

    Dim i As Integer
    Static PeakData(0 To 1023) As Single
    Static PeakDataExTop(0 To 1023) As Single


    nSample = (Width / (Barwidth + Space))
    Bandwidth = 1024# / nSample
    For X = 0 To 1023
        If PeakData(X) < (DrawData(X)) Then PeakData(X) = DrawData(X)
        If PeakDataExTop(X) < PeakData(X) Then PeakDataExTop(X) = PeakData(X) * 1.2
    Next X
    ''' for avge FFT values but didi't work fine
    'For X = 0 To nSample - 1
    '  Dim k As Single
    '  k = 0
    ' For i = 0 To bandwidth - 1
    '   k = k + PeakData(X * 1023# / nSample + i)
    '  Next
    '  PeakData(X) = k / bandwidth
    'Next

    For X = 0 To nSample

        Y = Sqr(Abs(PeakData(X)))
        R1 = R1 + 0.35 * (0.6 * Sin(0.98 * Y) + 0.4 * Sin(1.047 * Y))
        g1 = g1 + 0.35 * (0.6 * Sin(0.835 * Y) + 0.4 * Sin(1.081 * Y))
        b1 = b1 + 0.35 * (0.6 * Sin(0.814 * Y) + 0.4 * Sin(1.011 * Y))
        Y = Sqr(Abs(PeakDataExTop(X)))
        'Y = Y / 2

        NewX = Barwidth * X + Space * X
        NewY = (Height - Y) / 2

        ppen = CreatePen(0, 1, RGB(240, 240, 255))
        SelectObject hDC, ppen
        For i = 0 To Ywidth - 1
            MoveToEx hDC, NewX, NewY - i, dummyPt
            LineTo hDC, NewX + Barwidth, NewY - i
        Next

        SelectObject hDC, -1
        DeleteObject ppen


        'Col = RGB(Abs(r1 * 0.25), Abs(g1 * 0.25), Abs(b1 * 0.25))
        NewY = (Height + Y) / 2

        ppen = CreatePen(0, 1, RGB(40, 190, 200))
        SelectObject hDC, ppen
        MoveToEx hDC, NewX, NewY, dummyPt
        LineTo hDC, NewX + Barwidth, NewY
        SelectObject hDC, -1
        DeleteObject ppen

        PeakDataExTop(X) = PeakDataExTop(X) - PeakDataExTop(X) * 0.05
    Next X
    '------------------
    MoveToEx hDC, 0, Height / 2, dummyPt

    For X = 0 To nSample

        Y = Sqr(Abs(PeakData(X)))
        'Y = Y / 2
        R2 = R2 + 0.35 * (0.6 * Sin(0.98 * Y) + 0.4 * Sin(1.047 * Y))
        g2 = g2 + 0.35 * (0.6 * Sin(0.835 * Y) + 0.4 * Sin(1.081 * Y))
        b2 = b2 + 0.35 * (0.6 * Sin(0.814 * Y) + 0.4 * Sin(1.011 * Y))

        NewX = Barwidth * X + Space * X
        NewY = (Height - Y) / 2

        Col = RGB(Abs(R2) Mod 255, Abs(g2) Mod 255, Abs(b2) Mod 255)
        '//////////////////////////////
        'ppen = CreatePen(0, 1, RGB(123, 169, 220))
        ppen = CreatePen(0, 1, RGB(40, 255, 255))

        SelectObject hDC, ppen

        For i = 0 To Abs(NewY - Height / 2) / (Ywidth + Yspace)
        '    PicSpectrum.ForeColor = GetGradColor(maxX(1, (Abs(NewY - Height / 2) / 5)), i, vbRed, vbGreen, vbBlue)
            Dim j
            For j = 0 To Ywidth - 1
                MoveToEx hDC, NewX, Height / 2 - (Ywidth + Yspace) * i - j, dummyPt
                LineTo hDC, NewX + Barwidth, Height / 2 - (Ywidth + Yspace) * i - j
            Next
        Next

        SelectObject hDC, -1
        DeleteObject ppen

        '///////////////////////////////////////////
        ' mirror
        'Col = RGB(Abs(r2 / 2) Mod 128, Abs(g2 / 2) Mod 128, Abs(b2 / 2) Mod 128)
        Col = RGB(37, 77, 122)

        ppen = CreatePen(0, 1, Col)
        SelectObject hDC, ppen
        For i = 0 To Abs(NewY - Height / 2) / (Ywidth + Yspace)
          '  PicSpectrum.ForeColor = GetGradColor(maxX(1, (Abs(NewY - Height / 2) / 5)), i, vbRed, vbGreen, vbBlue)
            For j = 0 To Ywidth - 1
                MoveToEx hDC, NewX, Height / 2 + (Ywidth + Yspace) * i - j, dummyPt
                LineTo hDC, NewX + Barwidth, Height / 2 + (Ywidth + Yspace) * i - j
            Next
        Next
        SelectObject hDC, -1
        DeleteObject ppen

        PeakData(X) = PeakData(X) - PeakData(X) * 0.1
        'PeakData(x) = PeakData(x) - 800
    Next X

End Sub


Public Sub drawVis22(hDC As Long, DrawData() As Integer, Height As Single, Width As Single)
    Static X As Integer, Y As Single
    Dim dummyPt As POINTAPI
    Dim ppen As Long
    Static WaveData(0 To 1023) As Integer
    Static WaveDataEx(0 To 1023) As Integer
    Y = Sqr(Abs(WaveData(0)))
    NewX = Sin(X * 3.14 / 180) * (Y Mod Width) + (Width / 2)
    NewY = Cos(X * 3.14 / 180) * (Y Mod Height) + (Height / 2)

    MoveToEx hDC, Width / 2, Height / 2, dummyPt

    For X = 41 To 440    'UBound(DrawData) / 5
        If WaveData(X) < DrawData(X) Then WaveData(X) = DrawData(X)
        If WaveDataEx(X) < WaveData(X) Then WaveDataEx(X) = WaveData(X)
        Y = Sqr(Abs(WaveData(X)))
        If Abs(Y - Sqr(Abs(WaveData(X - 1)))) < 2 Then Y = Sqr(Abs(WaveData(X - 1)))

        tp_wave_r = tp_wave_r + 0.35 * (0.6 * Sin(0.98 * Y) + 0.4 * Sin(1.047 * Y))
        tp_wave_g = tp_wave_g + 0.35 * (0.6 * Sin(0.835 * Y) + 0.4 * Sin(1.081 * Y))
        tp_wave_b = tp_wave_b + 0.35 * (0.6 * Sin(0.814 * Y) + 0.4 * Sin(1.011 * Y))


        If Y < 0 Then Y = 1

        If Step < Y Then Step = Y


        NewX = Sin((X - 40) * 3.14159 / 200) * (Y Mod Width) + (Width / 2)
        NewY = Cos((X - 40) * 3.14159 / 200) * (Y Mod Height) + (Height / 2)



        Col = RGB(Abs(tp_wave_r) Mod 255, Abs(tp_wave_g) Mod 255, Abs(tp_wave_b) Mod 255)
        ppen = CreatePen(0, 1, Col)
        SelectObject hDC, ppen
        LineTo hDC, NewX, NewY
        SelectObject hDC, -1
        DeleteObject ppen

        Y = Sqr(Abs(WaveDataEx(X))) * 1.2
        If Y < 0 Then Y = 1

        If Step < Y Then Step = Y

        NewX = Sin((X - 40) * 3.14159 / 200) * (Y Mod Width) + (Width / 2)
        NewY = Cos((X - 40) * 3.14159 / 200) * (Y Mod Height) + (Height / 2)

        Col = RGB(((Abs(tp_wave_r) Mod 255)), ((Abs(tp_wave_g) Mod 255)), ((Abs(tp_wave_b) Mod 255)))

        SetPixel hDC, NewX + 1, NewY, Col
        SetPixel hDC, NewX, NewY + 1, Col
        SetPixel hDC, NewX, NewY - 1, Col
        SetPixel hDC, NewX - 1, NewY, Col

        WaveData(X) = WaveData(X) - WaveData(X) * 0.1
        WaveDataEx(X) = WaveDataEx(X) - WaveDataEx(X) * 0.04

    Next X


    Angle = Angle + (0.5)
    If Angle > 360 Then Angle = 0


End Sub
