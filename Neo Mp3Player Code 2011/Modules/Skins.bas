Attribute VB_Name = "Skins"

'+-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-+
'|   API PARA to DETERMINE COLOR of IMAGE                                 |
'+-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-+
'Used to transfer bit block of color data form source to dest treating crTransparent colour in source as transparent
'Public Declare Function TransparentBlt Lib "Msimg32.dll" (ByVal hdcDest As Long, ByVal nXOriginDest As Integer, ByVal nYOriginDest As Integer, ByVal nWidthDest As Integer, ByVal nHeightDest As Integer, ByVal hDCSrc As Long, ByVal nXOriginSrc As Integer, ByVal nYOriginSrc As Integer, ByVal nWidthSrc As Integer, ByVal nHeightSrc As Integer, ByVal crTransparent As Long) As Boolean

Public iTitleHeight As Integer
'get border width
Public iBorder As Integer
'+-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-+
'|   PROPORTIONAL API WALLPAPER                       |
'+-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-+
Public mainActive As Boolean
Public RoundedNess As Integer
Public plst_widthSpacing As Long
Public min_plstWidth As Long
Public max_plstWidth As Long
Public iVolScrollChange As Long
Public iPosScrollChange As Long
Public iSpectrum_refreshRate As Integer
Public eq_position As Integer
Public Bar_gap As Integer
Public CurrentTrack_Forecolor As Long
Public List_Backcolor As Long    'Color of normal text
Public SelectedText_Backcolor As Long    'Color of normal text
Public NormalText_Forecolor As Long    'Color of normal text
Public drawShade As Boolean
Public PeakHeight As Integer
Public bLoadingSkin As Boolean
'Public plst_RightSpacing As Long
'Public plst_TopSpacing As Long
'Public plst_BottomSpacing As Long


Public COLOR_START As Long         ' = vbGreen
Public COLOR_MIDDLE As Long    '= &H22A8E1
Public COLOR_END As Long         '= vbRed

Public Peak_Color As Long

' Verymuch same as BitBlt but is used to rapidly shrink or stretch Bitmap while copying into destination
Public Declare Function StretchBlt Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long
Public Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Public Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long
Public Declare Function SetBrushOrgEx Lib "gdi32" (ByVal hDC As Long, ByVal nXOrg As Long, ByVal nYOrg As Long, lppt As POINTAPI) As Long
Public Declare Function SetStretchBltMode Lib "gdi32" (ByVal hDC As Long, ByVal nStretchMode As Long) As Long
Public Declare Function UnrealizeObject Lib "gdi32" (ByVal hObject As Long) As Long
' BitBlt means(BitBlockTransfer) is used to transfer a big amount of bitmapped datafrom one DC(Display context) to another
' ROP is raster operation that is applied to transfer
' ROP values
' SRCCOPY copy directly
' SRCAND bits of source and dest are anded together
' SRCPAINT  logically Ors the pixels of source and destination
Public Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Public Declare Function GetPrivateProfileInt Lib "kernel32" _
                                             Alias "GetPrivateProfileIntA" _
                                             (ByVal sSectionName As String, _
                                              ByVal sKeyName As String, _
                                              ByVal lDefault As Long, _
                                              ByVal sFilename As String) As Long
Public Const STRETCH_HALFTONE As Long = &H4&



'+-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-+
'| APIS EFFECT OF OUTLINE FOR FORM                                         |
'| USED FOR IMAGE PROCESSING                                                |
'+-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-+
'returns colour of pixel in for of RGB i.e long value
Public Declare Function GetPixel Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long) As Long
'sets window region acc to region given in parameter(handle of region) REdraw is setb true so that window is redrawn by the system after setting region
Public Declare Function SetWindowRgn Lib "user32" (ByVal hwnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long
' creates region and returns handle in form of their handles)and stores them in new region returning
Public Declare Function CreateRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
' Combines source region 1 & 2 (gets parameter in form of their handles)and stores them in new region returning
' its handes as hdestregion
' combines region in various modes
' RGN_AND A(and)B intersection
' RGN_OR A+B
' RGN_COPY  creates a copy of region formed by source 1
' RGN_DIFF A-A(and)B
' RGN_XOR  A+B-A(and)B
Public Declare Function CombineRgn Lib "gdi32" (ByVal hDestRgn As Long, ByVal hSrcRgn1 As Long, ByVal hSrcRgn2 As Long, ByVal nCombineMode As Long) As Long
Public Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function SetPixel Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByVal crColor As Long) As Long
' get regiondata asevs pointer to region given as 1rst parameter in third parameter(out)no of bytes required POINTER MAY BE LONG OR BYTE
Public Declare Function GetRegionDataByte Lib "gdi32" Alias "GetRegionData" (ByVal hRgn As Long, ByVal dwCount As Long, lpRgnData As Byte) As Long
Public Declare Function GetRegionDataLong Lib "gdi32" Alias "GetRegionData" (ByVal hRgn As Long, ByVal dwCount As Long, lpRgnData As Long) As Long
' parameters are in,in,in ......returns handle to region if succeds else NULL
Public Declare Function ExtCreateRegionByte Lib "gdi32" Alias "ExtCreateRegion" (lpXform As Long, ByVal nCount As Long, lpRgnData As Byte) As Long
' moves a region by specefied offset x for hor. offset & y for vertical
Public Declare Function OffsetRgn Lib "gdi32" (ByVal hRgn As Long, ByVal X As Long, ByVal Y As Long) As Long

Type RegionDataType
    RegionData() As Byte
    DataLength As Long
End Type


Public Const RGN_OR = 2
Public Const WM_NCLBUTTONDOWN = &HA1
Public Const HTCAPTION = 2

Public EdgeRegions(2) As RegionDataType
Public Width_With_Eq As Long
Public Width_Without_Eq As Long

Private tConfigSlider(5) As ptSlider
Private iAlbumsShow As Integer, iAlbumsCols As Integer, iAlbumsRows As Integer
Attribute iAlbumsCols.VB_VarUserMemId = 1073741851
Attribute iAlbumsRows.VB_VarUserMemId = 1073741851

'+-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-+


Public Sub Read_Config_Skin()
'// how to read the settings of skin color and position of
'// Buttons
    Dim i As Integer, arryS(5) As String, arry() As String
    Dim s As String
    On Error Resume Next
    s = tAppConfig.AppConfig & "Skins\" & tAppConfig.Skin & "\"

    With frmMain

        .picNormalMode.cls
        .picNormalMode.Picture = LoadPicture()
        '.picNormalMode.Width = 5085
        '.picNormalMode.Height = 2415
        .picNormalMode.Picture = LoadPicture(s & "main.bmp")
        .picNormalMode.AutoSize = True
        .picNormalMode.Refresh

        .picWithoutEq.cls
        '.picWithoutEq.Width = 5085
        '.picWithoutEq.Height = 2415
        .picWithoutEq.Picture = LoadPicture(s & "picWithoutEq.bmp")
        .picWithoutEq.AutoSize = True



        .picMiniMode.cls
        .picMiniMode.Picture = LoadPicture()
        '.picMiniMode.Width = 4110
        '.picMiniMode.Height = 430
        .picMiniMode.Picture = LoadPicture(s & "minimode.bmp")
        .picNormalMode.AutoSize = True
        .picMiniMode.Refresh
        '//image state representation

        ColorBackGraph = Read_INI("NormalMode", "ColorBackGraph", RGB(0, 0, 0), True)
        Menu_foreColor = Read_INI("NormalMode", "InActiveMenuforeColor", &HFF0000, True)
        Menu_AciveColor = Read_INI("NormalMode", "ActiveMenuforeColor", &HCC00, True)
        Menu_BackColor = Read_INI("NormalMode", "MenuBackColor", &H22A8E1, True)
        Menu_GradStColor = Read_INI("NormalMode", "MenuGradStartColor", &HCC00, True)
        Menu_GradEndColor = Read_INI("NormalMode", "MenuGradendColor", &HCC00, True)
        Menu_SepBackColor = Read_INI("NormalMode", "MenuSepBackColor", &HCC00, True)
        Peak_Color = Read_INI("NormalMode", "Peak_Color", 0, True)

        CurrentTrack_Forecolor = Read_INI("NormalMode", "CurrentTrack_Forecolor", RGB(70, 100, 150), True)
        List_Backcolor = Read_INI("NormalMode", "List_Backcolor", vbBlack, True)
        SelectedText_Backcolor = Read_INI("NormalMode", "SelectedText_Backcolor", RGB(17, 26, 39), True)
        NormalText_Forecolor = Read_INI("NormalMode", "NormalText_Forecolor", vbWhite, True)

        drawShade = CBool(Read_INI("Normalmode", "DrawShade", "TRUE"))

        '// Colors for spectrum display
        COLOR_START = Read_INI("NormalMode", "COLOR_START  ", &HCC00, True)
        COLOR_MIDDLE = Read_INI("NormalMode", "COLOR_MIDDLE", &H22A8E1, True)
        COLOR_END = Read_INI("NormalMode", "COLOR_END ", &H122BC, True)

        'Public COLOR_START         As Long ' = vbGreen
        'Public COLOR_MIDDLE        As Long '= &H22A8E1
        'Public COLOR_END           As Long '= vbRed


        '// Previous
        Read_Config_Button .Button(0), "NormalMode", "Previous", "12,137,23,18"
        '// Play
        Read_Config_Button .Button(1), "NormalMode", "Play", "37,137,23,18"
        '// Pause
        Read_Config_Button .Button(2), "NormalMode", "Pause", "62,137,23,18"
        '// Stop
        Read_Config_Button .Button(3), "NormalMode", "Stop", "87,137,23,18"
        '// Next
        Read_Config_Button .Button(4), "NormalMode", "Next", "112,137,23,18"
        '// Intro
        Read_Config_Button .Button(5), "NormalMode", "Intro", "21,62,15,13"
        '// Mute
        Read_Config_Button .Button(6), "NormalMode", "Mute", "49,62,15,13"
        '// Repeat
        Read_Config_Button .Button(7), "NormalMode", "Repeat", "77,62,15,13"
        '// Randomize
        Read_Config_Button .Button(8), "NormalMode", "Randomize", "105,62,15,13"
        '// Previous Album
        Read_Config_Button .Button(9), "NormalMode", "PreviousAlbum", "195,1,30,12"
        '// Front
        Read_Config_Button .Button(10), "NormalMode", "Front", "227,1,21,12"
        '// Next Album
        Read_Config_Button .Button(11), "NormalMode", "NextAlbum", "250,1,30,12"
        '// Menu button
        Read_Config_Button .Button(12), "NormalMode", "Menu", "1,1,10,10"
        '// Minimize
        Read_Config_Button .Button(13), "NormalMode", "Minimize", "305,1,10,10"
        '// Minimode
        Read_Config_Button .Button(14), "NormalMode", "MiniMode", "316,1,10,10"
        '// Close
        Read_Config_Button .Button(15), "NormalMode", "Close", "327,1,10,10"
        '// eq_on
        Read_Config_Button .Button(16), "NormalMode", "EQ_ON", "1223,1360,31,13"
        '// eq_auto
        Read_Config_Button .Button(17), "NormalMode", "EQ_AUTO", "1332,70,50,13"
        '//eq_presets
        Read_Config_Button .Button(18), "NormalMode", "EQ_PRESET", "32,90,51,13"

        Read_Config_Button .Button(19), "NormalMode", "ECHO_ON", "4,30,31,9"
        '// eq_auto
        Read_Config_Button .Button(20), "NormalMode", "REC_EQ", "32,30,42,9"
        '//eq_presets
        Read_Config_Button .Button(21), "NormalMode", "STEREO_ON", "60,30,51,9"
        Read_Config_Button .Button(22), "NormalMode", "DEFAULT_ON", "60,30,53,9"
        Read_Config_Button .Button(23), "NormalMode", "Record", "1130,1330,53,9"
        Read_Config_Button .Button(24), "NormalMode", "EQ_IN_OUT", "1362,1337,23,18"
        Read_Config_Button .Button(25), "NormalMode", "OPEN", "60,-130,53,9"
        Read_Config_Button .Button(26), "NormalMode", "PLAYLIST", "60,-130,53,9"

        '
        '
        '
        ' For i = 0 To 7 Step 1
        Read_Config_Button frmPLST.plstButton(0), "NormalMode", "ADD_FILES", "60,30,53,9"
        ' Next
        Read_Config_Button frmPLST.plstButton(1), "NormalMode", "ADD_DIR", "60,30,53,9"
        Read_Config_Button frmPLST.plstButton(2), "NormalMode", "DELETE_FILES", "60,30,53,9"
        Read_Config_Button frmPLST.plstButton(3), "NormalMode", "SORT_FILES", "60,30,53,9"
        Read_Config_Button frmPLST.plstButton(4), "NormalMode", "LOAD_PLST", "60,30,53,9"
        Read_Config_Button frmPLST.plstButton(5), "NormalMode", "PLST_MIN", "60,30,53,9"
        Read_Config_Button frmPLST.plstButton(6), "NormalMode", "PLST_MODE", "60,30,53,9"
        Read_Config_Button frmPLST.plstButton(7), "NormalMode", "PLST_CLOSE", "60,30,53,9"
        Read_Config_Button frmPLST.PicActivate, "NormalMode", "PLSTPicActivate", "160,-230,53,9"


        '// PosBar
        Read_Config_Button .slider(0), "NormalMode", "PosBar", "1,120,144,10,V"
        '// volume bar
        Read_Config_Button .slider(1), "NormalMode", "VolBar", "148,25,10,121,H"
        '// time
        Read_Config_Button .slider(5), "NormalMode", "BalanceBar", "1119,45,74,9,H"


        Read_Config_Button .ScrollText(0), "NormalMode", "Time", "3,89,32,6"
        '// track title normal mode
        Read_Config_Button .ScrollText(1), "NormalMode", "TrackTitle", "1,110,144,6"
        'with formmain
        .ScrollText(1).BackColor = Read_INI("NormalMode", "TTBackColor", RGB(0, 0, 0), True)

        '// Bit Rate
        Read_Config_Button .ScrollText(2), "NormalMode", "BitRate", "38,80,15,6"
        '// Frequencia
        Read_Config_Button .ScrollText(3), "NormalMode", "Freq", "38,90,10,6"
        '  frmMain.ScrollText(1).Width = GetPrivateProfileInt("NormalMode", "scrolltext_width", 180, tAppConfig.AppConfig & "Skins\" & tAppConfig.Skin & "\Skin.ini")

        '// albums


        '// spectrum
        Read_Config_Button .PicSpectrum, "NormalMode", "Spectrum", "55,79,89,28"
        Read_Config_Button .PicSpectrumMini, "MiniMode", "Spectrum", "55,19,29,28"

        '// spectrum bars
        tSpectrum.bDrawBars = CBool(Read_INI("NormalMode", "DrawBars", True))

        tSpectrum.iBars = CInt(Read_INI("NormalMode", "Bars", 15))
        If tSpectrum.iBars > 50 Then tSpectrum.iBars = 50
        If tSpectrum.iBars <= 0 Then tSpectrum.iBars = 1

        tSpectrum.iSpacio = CInt(Read_INI("NormalMode", "SpaceBar", 2))
        If tSpectrum.iSpacio > 5 Then tSpectrum.iSpacio = 5
        If tSpectrum.iSpacio < 0 Then tSpectrum.iSpacio = 0

        tSpectrum.lBackColorBar = CLng(Read_INI("NormalMode", "BackColorBar", RGB(255, 255, 255), True))
        tSpectrum.lLineColorBar = CLng(Read_INI("NormalMode", "LineColorBar", RGB(255, 255, 255), True))

        '// spectrum peaks
        tSpectrum.bDrawPeaks = CBool(Read_INI("NormalMode", "DrawPeaks", True))

        tSpectrum.lBackColorPeak = CLng(Read_INI("NormalMode", "BackColorPeak", RGB(255, 255, 255), True))

        tSpectrum.iPeakHeight = CInt(Read_INI("NormalMode", "PeakHeight", 1))
        If tSpectrum.iPeakHeight > 3 Then tSpectrum.iPeakHeight = 3
        If tSpectrum.iPeakHeight <= 0 Then tSpectrum.iPeakHeight = 1

        tSpectrum.iPeakGravity = CInt(Read_INI("NormalMode", "PeakGravity", 1))
        If tSpectrum.iPeakGravity > 10 Then tSpectrum.iPeakGravity = 10
        If tSpectrum.iPeakGravity <= 0 Then tSpectrum.iPeakGravity = 1

        '// spectrum scope
        tSpectrum.iLinesScope = CInt(Read_INI("NormalMode", "LinesScope", 30))
        If tSpectrum.iLinesScope > 50 Then tSpectrum.iLinesScope = 50
        If tSpectrum.iLinesScope <= 0 Then tSpectrum.iLinesScope = 10

        tSpectrum.lBackColorScope = CLng(Read_INI("NormalMode", "BackColorScope", RGB(255, 255, 255), True))

        If tSpectrum.bDrawBars = False And tSpectrum.bDrawPeaks = False Then tSpectrum.bDrawBars = True

        arryS(0) = "PosSlider": arryS(1) = "VolSlider"
        arryS(2) = "ListSlider"
        arryS(3) = "PosSlider": arryS(4) = "VolSlider"
        arryS(5) = "BalanceSlider"
        eq_position = GetPrivateProfileInt("NormalMode", "Eq_position", 1, tAppConfig.AppConfig & "Skins\" & tAppConfig.Skin & "\Skin.ini")
        Width_With_Eq = GetPrivateProfileInt("NormalMode", "Widthwith_Eq", 6111, tAppConfig.AppConfig & "Skins\" & tAppConfig.Skin & "\Skin.ini")
        'CInt(Read_INI("NormalMode", "Widthwith_Eq", "3775", tAppConfig.AppConfig & "Skins\" & tAppConfig.Skin & "\Skin.ini"))
        Width_Without_Eq = GetPrivateProfileInt("NormalMode", "Widthwithout_Eq", Width_With_Eq, tAppConfig.AppConfig & "Skins\" & tAppConfig.Skin & "\Skin.ini")

        eq_position = GetPrivateProfileInt("NormalMode", "Eq_position", 1, tAppConfig.AppConfig & "Skins\" & tAppConfig.Skin & "\Skin.ini")

        Read_Config_Button .PicEqAmp, "NormalMode", "EQ_CURVE", "570,2280,1935,315"
        Read_Config_Button .PicActivate, "NormalMode", "PicActivate", "500,-200,1935,315"


        'Width_Without_Eq = CInt(Read_INI("NormalMode", "Widthwithout_Eq", "3775", tAppConfig.AppConfig & "Skins\" & tAppConfig.Skin & "\Skin.ini"))

        '// sliders config normal Mode
        For i = 0 To 2
            s = Read_INI("NormalMode", arryS(i), "10,10")    'no need to give path of ini file
            arry = Split(s, ",", , vbTextCompare)
            If UBound(arry) = 1 Then
                tConfigSlider(i).Width = arry(0)
                tConfigSlider(i).Height = arry(1)
            Else
                tConfigSlider(i).Width = 10
                tConfigSlider(i).Height = 10
            End If
        Next i



        '// sliders config minimode
        For i = 3 To 5
            If i <> 5 Then
                s = Read_INI("MiniMode", arryS(i), "10,10")
            Else
                s = Read_INI("NormalMode", arryS(i), "10,10")
            End If
            arry = Split(s, ",", , vbTextCompare)
            If UBound(arry) = 1 Then
                tConfigSlider(i).Width = arry(0)
                tConfigSlider(i).Height = arry(1)
            Else
                tConfigSlider(i).Width = 10
                tConfigSlider(i).Height = 10
            End If
        Next i


        '=====================================================================
        '  MINI MODE
        '=====================================================================

        '// Previous
        Read_Config_Button .ButtonMini(0), "MiniMode", "Previous", "172,1,10,10"
        '// Play
        Read_Config_Button .ButtonMini(1), "MiniMode", "Play", "183,1,10,10"
        '// Pause
        Read_Config_Button .ButtonMini(2), "MiniMode", "Pause", "194,1,10,10"
        '// Stop
        Read_Config_Button .ButtonMini(3), "MiniMode", "Stop", "205,1,10,10"
        '// Next
        Read_Config_Button .ButtonMini(4), "MiniMode", "Next", "216,1,10,10"
        '// Menu button
        Read_Config_Button .ButtonMini(5), "MiniMode", "menu", "1,1,10,10"
        '// Minimize
        Read_Config_Button .ButtonMini(6), "MiniMode", "Minimize", "239,1,10,10"
        '// Minimode
        Read_Config_Button .ButtonMini(7), "MiniMode", "NormalMode", "250,1,10,10"
        '// Close
        Read_Config_Button .ButtonMini(8), "MiniMode", "Close", "261,1,10,10"
        '//playlist show/hide
        Read_Config_Button .ButtonMini(9), "MiniMode", "Playlist", "1261,1,10,10"
        '// record
        Read_Config_Button .ButtonMini(10), "MiniMode", "Record", "1261,1,10,10"
        '//mute
        Read_Config_Button .ButtonMini(11), "MiniMode", "mute", "1261,1,10,10"
        '//repeat
        Read_Config_Button .ButtonMini(12), "MiniMode", "repeat", "1261,1,10,10"
        '//random
        Read_Config_Button .ButtonMini(13), "MiniMode", "random", "1261,1,10,10"

        Read_Config_Button .ButtonMini(14), "MiniMode", "open", "1261,1,10,10"


        '// time
        Read_Config_Button .ScrollText(4), "MiniMode", "Time", "13,3,25,6"
        '// track title normal mode
        Read_Config_Button .ScrollText(5), "MiniMode", "TrackTitle", "43,3,128,6"
        .ScrollText(5).BackColor = Read_INI("MiniMode", "TTBackColor", RGB(0, 0, 0), True)
        '// PosBar
        Read_Config_Button .slider(3), "MiniMode", "PosBar", "41,13,97,6,V"
        '// volume bar
        Read_Config_Button .slider(4), "MiniMode", "VolBar", "147,13,58,6,V"

        'frmMain.EqBackpic.Left = GetPrivateProfileInt("MiniMode", "Eqbackpicleft", -4000, s + "Skin.ini")
        'frmMain.EqBackpic.Top = GetPrivateProfileInt("MiniMode", "Eqbackpictop", 4000, s + "Skin.ini")




        If bolLyricsShow = True Then

        End If
    End With

    '//////////////////// PLAYLIST SKIN

End Sub

Public Function MakeRegion(picSkin As PictureBox) As Long

'show an hourglass pointer
'frmmain.MousePointer = vbArrowHourglass
' SkinImage = frmmain.picnormalmode.Picture
'get titlebar height
    Dim s As String
    'frmMain.picNormalMode.AutoSize = True

    'S = tAppConfig.AppConfig & "Skins\" & tAppConfig.Skin & "\"

    'frmMain.picNormalMode.Picture = LoadPicture(S & "main.bmp")
    ' frmMain.picNormalMode.ScaleHeight = frmMain.picNormalMode.ScaleHeight + (iBorder) / Screen.TwipsPerPixelY
    iBorder = (frmMain.Width - frmMain.ScaleWidth)
    iTitleHeight = (frmMain.Height - frmMain.ScaleHeight)

    ' frmMain.WindowState = vbNormal
    'frmMain.Width = frmMain.picNormalMode.ScaleWidth * Screen.TwipsPerPixelX + iBorder
    'add space for the border and titlebar
    'frmMain.Height = frmMain.picNormalMode.ScaleHeight * Screen.TwipsPerPixelY + iBorder + iTitleHeight
    'move to the center of the screen
    'frmMain.Move (Screen.Width - frmMain.Width) / 2, (Screen.Height - frmMain.Height) / 2

    'cover the form with the splash screen
    'frmSplash.Move frmmain.Left - 100, frmmain.Top - 100, frmmain.Width + 200, frmmain.Height + 200

    'convert twips to pixels
    'iBorder = iBorder / Screen.TwipsPerPixelX
    'iTitleHeight = iTitleHeight / Screen.TwipsPerPixelY

    'now allow all forms to resize before continuing
    'DoEvents

    'vars: final region, temporary region
    Dim iRegion As Long, iTempRgn As Long
    'create an empty region to start with
    iRegion = CreateRectRgn(0, 0, 0, 0)

    'vars: current row, col, transparent colour,
    'device context storage, width of image, start of
    'solid block
    Dim Y As Long, X As Long, iTransparent As Long
    Dim ihDC As Long, iwidth As Long, XStart As Long
    'save the device context handle to avoid extra calls
    'to the property of the picturebox (to speed up the code)
    ihDC = picSkin.hDC
    'calculate the width in pixels
    iwidth = picSkin.ScaleWidth    '+iborder'/ Screen.TwipsPerPixelY
    'get the colour of the top-left pixel: will be used
    'as transparent colour
    'iTransparent = RGB(255, 0, 255)
    iTransparent = Read_INI("NormalMode", "ColorTrans", RGB(255, 0, 255), True)

    'scan every row
    For Y = 0 To picSkin.ScaleHeight - 1

        Do While X < iwidth
            'skip all "transparent" pixels in this row
            Do While X < iwidth And GetPixel(ihDC, X, Y) = iTransparent
                X = X + 1
            Loop

            'if we did not go through the entire row, there
            'are some pixels that need to be shown
            If X < iwidth Then
                'save the start of the non-transparent block
                XStart = X
                'now find the end of the block of
                'non-transparent pixels. Stop if the edge
                'is reached as well.
                Do While X < iwidth And GetPixel(ihDC, X, Y) <> iTransparent
                    X = X + 1
                Loop

                'create a region of the same size as the
                'non-transparent block (1px high)
                iTempRgn = CreateRectRgn(XStart + iBorder / Screen.TwipsPerPixelX - 1, Y + iTitleHeight / Screen.TwipsPerPixelY - 1, X + iBorder / Screen.TwipsPerPixelX - 1, Y + iTitleHeight / Screen.TwipsPerPixelX)
                'now combine it with the final region
                'using OR (2)
                CombineRgn iRegion, iRegion, iTempRgn, 2
                'remove the GDI object used by the temporary
                'region to conserve memory
                DeleteObject iTempRgn
            End If
        Loop
        'reset the x counter
        X = 0
    Next Y

    'set the window region, and repaint (True)
    'SetWindowRgn frmMain.hWnd, iRegion, True
    MakeRegion = iRegion
    'frmMain.WindowState = vbNormal
    'frmMain.picNormalMode.Move 0, 0 ', frmmain.picnormalmode.Width, frmmain.picnormalmode.Height

    'size the form so that the border is cut off
    ' frmMain.Height = frmMain.Height - iBorder * Screen.TwipsPerPixelY  '+ Screen.TwipsPerPixelY
    'frmMain.Width = frmMain.Width - iBorder * Screen.TwipsPerPixelX * 2

    'restore normal mouse pointer
    'frmMain.MousePointer = vbDefault
    'frmMain.picNormalMode.Left = (iBorder + 1) * Screen.TwipsPerPixelX / 4
    'frmMain.picNormalMode.Top = (iBorder + 1) * Screen.TwipsPerPixelY / 4
End Function

'+-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-+
'+-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-+

Public Sub Load_Buttons_Skin()
'// procedimiento para cargar todos los controles, ponerlos en su lugar

    Dim srcX As Integer, srcY As Integer, srcWidth As Integer, srcHeight As Integer
    Dim i As Integer, j As Integer, k As Integer
    Dim pTemp(18) As StdPicture, pImage As StdPicture
    Dim s As String
    Dim lColorTran As Long
    iBorder = (frmMain.Width - frmMain.ScaleWidth)
    iTitleHeight = (frmMain.Height - frmMain.ScaleHeight)
    On Error Resume Next
    With frmMain
        s = tAppConfig.AppConfig & "Skins\" & tAppConfig.Skin & "\"
        frmMain.PicActivate.Picture = LoadPicture(s & "picActivate.bmp")



        Res = Read_Config_Button(frmMain.PicInfo, "NormalMode", "picinfo", "11125,29,13,11", s + "Skin.ini")
        frmMain.PicConfigInfo.Picture = LoadPicture(s & "picConfiginfo.bmp")

        '// time font
        Set .ScrollText(0).PictureText = LoadPicture(s & "num_font.bmp")
        '// track title
        Set .ScrollText(1).PictureText = LoadPicture(s & "song_font.bmp")
        '// bitrate text
        Set .ScrollText(2).PictureText = LoadPicture(s & "songinfo_font.bmp")
        '// frecuencia text
        Set .ScrollText(3).PictureText = LoadPicture(s & "songinfo_font.bmp")
        '// time minimode
        Set .ScrollText(4).PictureText = LoadPicture(s & "num_minimode_font.bmp")
        '// track title minimode
        Set .ScrollText(5).PictureText = LoadPicture(s & "song_minimode_font.bmp")


        With frmPLST
            If frmPLST.picList.Width < min_plstWidth Then frmPLST.picList.Width = min_plstWidth
            .PicBottomRight.Picture = LoadPicture(s & "PicBottomRight.bmp")
            .PicBottomLeft.Picture = LoadPicture(s & "PicBottomleft.bmp")
            .Picright.Picture = LoadPicture(s & "PicRight.bmp")
            .PicActivate.Picture = LoadPicture(s & "PLSTpicActivate.bmp")
            .PicTemp.Picture = LoadPicture(s & "\PLAY_LIST\hsegment_top_header_deactivate.bmp")
            .Form_Load1
        End With

        Set pTemp(0) = LoadPicture(s & "player_buttons.bmp")
        Set pTemp(1) = LoadPicture(s & "options_buttons.bmp")
        Set pTemp(2) = LoadPicture(s & "albums_buttons.bmp")
        Set pTemp(3) = LoadPicture(s & "titlebar_buttons.bmp")
        Set pTemp(4) = LoadPicture(s & "posbar_slider.bmp")
        Set pTemp(5) = LoadPicture(s & "volbar_slider.bmp")
        Set pTemp(6) = LoadPicture(s & "listbar_slider.bmp")

        '// minimode pictures
        Set pTemp(7) = LoadPicture(s & "posbar_minimode_slider.bmp")
        Set pTemp(8) = LoadPicture(s & "volbar_minimode_slider.bmp")


        Set pTemp(9) = LoadPicture(s & "albums_picture.bmp")

        '// minimode
        Set pTemp(10) = LoadPicture(s & "player_minimode_buttons.bmp")
        Set pTemp(11) = LoadPicture(s & "titlebar_minimode_buttons.bmp")
        Set pTemp(12) = LoadPicture(s & "Eq_buttons.bmp")
        Set pTemp(13) = LoadPicture(s & "Rec_buttons.bmp")
        Set pTemp(14) = LoadPicture(s & "RecButton.bmp")
        Set pTemp(15) = LoadPicture(s & "PLST_Buttons.bmp")
        Set pTemp(16) = LoadPicture(s & "MINIMODE_OPTIONButtons.bmp")
        Set pTemp(17) = LoadPicture(s & "Balance_slider.bmp")

        lColorTran = Read_INI("NormalMode", "ColorTrans", RGB(255, 0, 255), True)

        .PicTemp.BackColor = &H808080

        For i = 0 To 26
            .Button(i).Reset
            .Button(i).MaskColor = lColorTran
            srcWidth = .Button(i).Width
            srcHeight = .Button(i).Height
            'If i >= 16 Then
            '  srcWidth = .Button(i).Width
            '  srcHeight = .Button(i).Height
            ' End If
            .PicTemp.Width = srcWidth
            .PicTemp.Height = srcHeight

            '// copy picture back
            .PicTemp.Picture = LoadPicture()
            .PicTemp.PaintPicture .picNormalMode.Image, 0, 0, srcWidth, srcHeight, .Button(i).Left, .Button(i).Top, srcWidth, srcHeight
            .PicTemp.Picture = .PicTemp.Image
            Set .Button(i).PictureBack = .PicTemp.Picture

            If i = 0 Then    '// play buttons from i=0 to i=4
                srcX = 0
                Set pImage = pTemp(0)    '"player_buttons.bmp"
            ElseIf i = 5 Then    '// options buttons from i=5 to 8
                srcX = 0
                Set pImage = pTemp(1)    '"option_buttons.bmp"
            ElseIf i = 9 Then    '// albums buttons from i=9 to 11
                srcX = 0
                Set pImage = pTemp(2)    '"album_buttons.bmp"
            ElseIf i = 12 Then    '// titlebar buttons from i=12 to 15
                srcX = 0
                Set pImage = pTemp(3)    '"titlebar_buttons.bmp"
            ElseIf i = 16 Then
                srcX = 0
                Set pImage = pTemp(12)    'Eq_buttons
            ElseIf i = 19 Then
                srcX = 0
                Set pImage = pTemp(13)    'rec_buttons
            ElseIf i = 23 Then
                srcX = 0
                Set pImage = pTemp(14)    'rec_buttons
            End If


            For j = 0 To 3    'for button their are 4 states each belonging to j=0(normal),1(focus mouse over it),2(picture down),3(picture disabled)
                srcY = srcHeight * j    'button with various states are one below other in player_buttons.bmp
                .PicTemp.Picture = LoadPicture()
                .PicTemp.PaintPicture pImage, 0, 0, srcWidth, srcHeight, srcX, srcY, srcWidth, srcHeight
                .PicTemp.Picture = .PicTemp.Image

                If j = 0 Then    ' stores pictures in button control picture property
                    Set .Button(i).PictureNormal = .PicTemp.Picture
                ElseIf j = 1 Then
                    Set .Button(i).PictureOver = .PicTemp.Picture
                ElseIf j = 2 Then
                    Set .Button(i).PictureDown = .PicTemp.Picture
                Else
                    Set .Button(i).PictureDisabled = .PicTemp.Picture
                End If
            Next j

            srcX = srcX + srcWidth    ' goto next button horizontally next in player_buttons.bmp in skin file

            '  DoEvents 'commented at 10:30 15/3/09
        Next i





        For i = 0 To 7 Step 1

            frmPLST.plstButton(i).Reset
            frmPLST.plstButton(i).MaskColor = lColorTran
            srcWidth = frmPLST.plstButton(i).Width
            srcHeight = frmPLST.plstButton(i).Height
            .PicTemp.Width = srcWidth
            .PicTemp.Height = srcHeight

            '// copy picture back
            .PicTemp.Picture = LoadPicture()
            .PicTemp.PaintPicture .picNormalMode.Image, 0, 0, srcWidth, srcHeight, .Button(i).Left, .Button(i).Top, srcWidth, srcHeight
            .PicTemp.Picture = .PicTemp.Image
            Set frmPLST.plstButton(i).PictureBack = .PicTemp.Picture
            If i = 0 Then
                srcX = 0
                Set pImage = pTemp(15)    'rec_
            End If

            For j = 0 To 3    'for button there are 4 states each belonging to j=0(normal),1(focus mouse over it),2(picture down),3(picture disabled)
                srcY = srcHeight * j    'button with various states are one below other in player_buttons.bmp
                .PicTemp.Picture = LoadPicture()
                .PicTemp.PaintPicture pImage, 0, 0, srcWidth, srcHeight, srcX, srcY, srcWidth, srcHeight
                .PicTemp.Picture = .PicTemp.Image

                If j = 0 Then    ' stores pictures in button control picture property
                    Set frmPLST.plstButton(i).PictureNormal = .PicTemp.Picture
                ElseIf j = 1 Then
                    Set frmPLST.plstButton(i).PictureOver = .PicTemp.Picture
                ElseIf j = 2 Then
                    Set frmPLST.plstButton(i).PictureDown = .PicTemp.Picture
                Else
                    Set frmPLST.plstButton(i).PictureDisabled = .PicTemp.Picture
                End If
            Next j

            srcX = srcX + srcWidth    ' goto next button horizontally next in player_buttons.bmp in skin file
            DoEvents    'commented at 10:30 15/3/09
        Next i



        frmMain.picNormalMode.Refresh

        '// Sliders pos - vol - list  ---- and minimode
        '// code for sliders picture capture
        For i = 0 To 5
            .PicTemp.BackColor = &H808080
            .slider(i).ResetPictures

            srcWidth = .slider(i).Width
            srcHeight = .slider(i).Height

            .PicTemp.Width = srcWidth
            .PicTemp.Height = srcHeight
            srcX = 0
            srcY = 0
            '// picture back
            .PicTemp.Picture = LoadPicture()
            If i = 5 Then
                .PicTemp.PaintPicture pTemp(17), 0, 0, srcWidth, srcHeight, srcX, srcY, srcWidth, srcHeight
            Else
                .PicTemp.PaintPicture pTemp(i + 4), 0, 0, srcWidth, srcHeight, srcX, srcY, srcWidth, srcHeight
            End If
            .PicTemp.Picture = .PicTemp.Image
            Set .slider(i).PictureBack = .PicTemp.Picture

            srcY = srcHeight
            '// picture progress
            .PicTemp.Picture = LoadPicture()
            If i = 5 Then
                .PicTemp.PaintPicture pTemp(17), 0, 0, srcWidth, srcHeight, srcX, srcY, srcWidth, srcHeight
            Else
                .PicTemp.PaintPicture pTemp(i + 4), 0, 0, srcWidth, srcHeight, srcX, srcY, srcWidth, srcHeight
            End If
            .PicTemp.Picture = .PicTemp.Image
            Set .slider(i).PictureProgress = .PicTemp.Picture

            .PicTemp.BackColor = &HC0&

            '// .Sliders
            srcX = srcWidth
            srcWidth = tConfigSlider(i).Width
            srcHeight = tConfigSlider(i).Height

            .PicTemp.Width = srcWidth
            .PicTemp.Height = srcHeight

            For j = 0 To 2
                srcY = srcHeight * j
                .PicTemp.Picture = LoadPicture()
                If i = 5 Then
                    .PicTemp.PaintPicture pTemp(17), 0, 0, srcWidth, srcHeight, srcX, srcY, srcWidth, srcHeight
                Else
                    .PicTemp.PaintPicture pTemp(i + 4), 0, 0, srcWidth, srcHeight, srcX, srcY, srcWidth, srcHeight
                End If
                .PicTemp.Picture = .PicTemp.Image

                If j = 0 Then
                    Set .slider(i).bar = .PicTemp.Picture
                ElseIf j = 1 Then
                    Set .slider(i).BarOver = .PicTemp.Picture
                Else
                    Set .slider(i).BarDown = .PicTemp.Picture
                End If
            Next j

            DoEvents    'commented at 10:30 15/3/09
        Next i

        '//==============================================================================
        '// Mini mode buttons
        '//==============================================================================

        .PicTemp.BackColor = &H808080

        For i = 0 To 14
            .ButtonMini(i).Reset
            .ButtonMini(i).MaskColor = lColorTran
            srcWidth = .ButtonMini(i).Width
            srcHeight = .ButtonMini(i).Height

            .PicTemp.Width = srcWidth
            .PicTemp.Height = srcHeight

            '// copy picture back
            .PicTemp.Picture = LoadPicture()
            .PicTemp.PaintPicture .picMiniMode.Image, 0, 0, srcWidth, srcHeight, .ButtonMini(i).Left, .ButtonMini(i).Top, srcWidth, srcHeight
            .PicTemp.Picture = .PicTemp.Image
            Set .ButtonMini(i).PictureBack = .PicTemp.Picture

            If i = 0 Then    '// play buttons
                srcX = 0
                Set pImage = pTemp(10)
            ElseIf i = 5 Then    '// options buttons
                srcX = 0
                Set pImage = pTemp(11)
            ElseIf i = 9 Then    '// options buttons
                srcX = 0
                Set pImage = pTemp(16)
            End If

            For j = 0 To 3
                srcY = srcHeight * j
                .PicTemp.Picture = LoadPicture()
                .PicTemp.PaintPicture pImage, 0, 0, srcWidth, srcHeight, srcX, srcY, srcWidth, srcHeight
                .PicTemp.Picture = .PicTemp.Image

                If j = 0 Then
                    Set .ButtonMini(i).PictureNormal = .PicTemp.Picture
                ElseIf j = 1 Then
                    Set .ButtonMini(i).PictureOver = .PicTemp.Picture
                ElseIf j = 2 Then
                    Set .ButtonMini(i).PictureDown = .PicTemp.Picture
                Else
                    Set .ButtonMini(i).PictureDisabled = .PicTemp.Picture
                End If
            Next j

            srcX = srcX + srcWidth

            DoEvents    'commented at 10:30 15/3/09
        Next i



        Dim kHeight, kTop, kLeft, kSpacing, kWidth As Integer

        frmMain.EqBackpic.Picture = LoadPicture(s + "EqBackpic.bmp")
        Res = Read_Config_Button(frmMain.EqBackpic, "NormalMode", "Eqbackpic", "120,10000,3000,3000", s + "Skin.ini")
        kWidth = GetPrivateProfileInt("NormalMode", "EqSliderwidth", 120, s + "Skin.ini")
        kHeight = GetPrivateProfileInt("NormalMode", "EqSliderHeight", 1000, s + "Skin.ini")
        kTop = GetPrivateProfileInt("NormalMode", "EqSliderTop", 800, s + "Skin.ini")
        kLeft = GetPrivateProfileInt("NormalMode", "EqSliderLeft", 100, s + "Skin.ini")
        kSpacing = GetPrivateProfileInt("NormalMode", "EqSliderSpacing", 100, s + "Skin.ini")

        For i = 0 To 9 Step 1
            frmMain.Eq_SliderCtrl(i).Visible = False
            DoEvents    'commented at 10:30 15/3/09
            frmMain.Eq_SliderCtrl(i).Move kLeft + kSpacing * i, kTop, kWidth, kHeight
            Set frmMain.Eq_SliderCtrl(i).PictureBack = LoadPicture(s + "Eqslider back1.bmp")
            Set frmMain.Eq_SliderCtrl(i).PictureProgress = LoadPicture(s + "Eqslider back.bmp")
            Set frmMain.Eq_SliderCtrl(i).bar = LoadPicture(s + "Eqslider bar.bmp")
            Set frmMain.Eq_SliderCtrl(i).BarOver = LoadPicture(s + "Eqslider bar over.bmp")
            Set frmMain.Eq_SliderCtrl(i).BarDown = LoadPicture(s + "Eqslider bar down.bmp")
            frmMain.Eq_SliderCtrl(i).Visible = True
        Next

        frmMain.balance.Visible = True
        frmMain.DrawEQ

        Res = Read_Config_Button(frmMain.PicShuffle, "NormalMode", "picShuffle", "177,47,31,8", s + "Skin.ini")
        Res = Read_Config_Button(frmMain.PicRepeat, "NormalMode", "picRepeat", "177,58,31,8", s + "Skin.ini")
        Res = Read_Config_Button(frmMain.PicCrossfade, "NormalMode", "picCrossfade", "177,69,31,8", s + "Skin.ini")
        Res = Read_Config_Button(frmMain.PicaTop, "NormalMode", "picaTop", "177,80,31,8", s + "Skin.ini")

        Res = Read_Config_Button(frmMain.balance, "NormalMode", "Balance", "120,1000,3000,3000", s + "Skin.ini")
        frmMain.Player_Repeat (PlayerLoop)
        frmMain.Button(8).Selected = Not frmMain.Button(8).Selected
        frmMain.Button_Click (8)    'again toggles button 8.selected, now picture is displayed
        Atop = Not Atop
        frmPopUp.Atop_Click

        Set frmMain.balance.PictureBack = LoadPicture(s + "Balance_back.bmp")
        Set frmMain.balance.PictureProgress = LoadPicture(s + "Balance_back1.bmp")
        'frmMain.BALANCE.PictureBack = LoadPicture(s + "balance back.bmp")
        Set frmMain.balance.bar = LoadPicture(s + "Balance_bar_normal.bmp")
        Set frmMain.balance.BarOver = LoadPicture(s + "Balance_bar_over.bmp")
        Set frmMain.balance.BarDown = LoadPicture(s + "Balance_bar_down.bmp")

        Set pTemp(0) = LoadPicture()
        Set pTemp(1) = LoadPicture()
        Set pTemp(2) = LoadPicture()
        Set pTemp(3) = LoadPicture()
        Set pTemp(4) = LoadPicture()
        Set pTemp(5) = LoadPicture()
        Set pTemp(6) = LoadPicture()
        Set pTemp(7) = LoadPicture()
        Set pTemp(8) = LoadPicture()
        Set pTemp(9) = LoadPicture()
        Set pTemp(10) = LoadPicture()
        Set pTemp(11) = LoadPicture()
        Set pImage = LoadPicture()

        .PicTemp = LoadPicture()
    End With

End Sub


Public Sub Load_Cursors()
    On Error Resume Next
    Dim sPath As String, sCursor As String
    Dim i As Integer

    sPath = tAppConfig.AppConfig & "Skins\" & tAppConfig.Skin & "\"

    With frmMain
        '=======================================================================
        ' NORMAL MODE
        '=======================================================================
        '// cursor principal
        sCursor = ""
        If Dir(sPath & "curMain.cur") <> "" Then sCursor = sPath & "curMain.cur"
        .picNormalMode.MouseIcon = LoadPicture(sCursor)

        '// cursor for buttons normalmode
        sCursor = ""
        If Dir(sPath & "curButtons.cur") <> "" Then sCursor = sPath & "curButtons.cur"

        For i = 0 To 26
            '//mouse icon for normal mode
            Set .Button(i).MouseIcon = LoadPicture(sCursor)
        Next i

        For i = 0 To 7
            Set frmPLST.plstButton(i).MouseIcon = LoadPicture(sCursor)
        Next i

        '// cCursor for albums
        sCursor = ""
        '  If Dir(sPath & "curAlbums.cur") <> "" Then sCursor = sPath & "curAlbums.cur"
        ' Set .btnAlbum(1).MouseIcon = LoadPicture(sCursor)

        '// cursor posbar
        sCursor = ""
        If Dir(sPath & "curposbar.cur") <> "" Then sCursor = sPath & "curposbar.cur"
        Set .slider(0).MouseIcon = LoadPicture(sCursor)

        '// cursor vol bar
        sCursor = ""
        If Dir(sPath & "curvolbar.cur") <> "" Then sCursor = sPath & "curvolbar.cur"
        Set .slider(1).MouseIcon = LoadPicture(sCursor)

        For i = 0 To 9
            Set .Eq_SliderCtrl(i).MouseIcon = LoadPicture(sCursor)
        Next

        '// Cursor to the play list
        sCursor = ""
        If Dir(sPath & "curListRep.cur") <> "" Then sCursor = sPath & "curlistrep.cur"
        Set frmPLST.picList.MouseIcon = LoadPicture(sCursor)

        '// Cursor to slider of playlist
        sCursor = ""
        If Dir(sPath & "curListbar.cur") <> "" Then sCursor = sPath & "curlistbar.cur"
        Set frmPLST.picBack.MouseIcon = LoadPicture(sCursor)


        '=============================================================================
        ' MINI MODE
        '=============================================================================
        '// cursor minimode
        sCursor = ""
        If Dir(sPath & "curMiniMode.cur") <> "" Then sCursor = sPath & "curMiniMode.cur"
        .picMiniMode.MouseIcon = LoadPicture(sCursor)

        '// cursor buttons to mini mode
        sCursor = ""
        If Dir(sPath & "curButtons_minimode.cur") <> "" Then sCursor = sPath & "curButtons_minimode.cur"

        For i = 0 To 8
            '// minimascara
            Set .ButtonMini(i).MouseIcon = LoadPicture(sCursor)
        Next i

        '// cursor posbar
        sCursor = ""
        If Dir(sPath & "curposbar_minimode.cur") <> "" Then sCursor = sPath & "curposbar_minimode.cur"
        Set .slider(3).MouseIcon = LoadPicture(sCursor)

        '// cursor vol bar
        sCursor = ""
        If Dir(sPath & "curvolbar_minimode.cur") <> "" Then sCursor = sPath & "curvolbar_minimode.cur"
        Set .slider(4).MouseIcon = LoadPicture(sCursor)

    End With
End Sub

'+-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-+
'+-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-+

Public Sub Change_Skin(SkinName As String)
'On Error Resume Next

    tAppConfig.Skin = SkinName

    '---------------------------------------------------------------------------------------
    '// read the configuration of the skin of the button positions
    If bLoading = True Then Read_Config_Skin

    '----------------------------------------------------------------------------------------
    '// placing the buttons if you have parts that can be transparent
    Load_Buttons_Skin

    '----------------------------------------------------------------------------------------
    '// position the buttons if you have parts that can be transparent
    Load_Cursors

    '---------------------------------------------------------------------------------------
    '// upload albums


    If boolVisLoaded = True Then FrmVisualisation.LoadSkin
    If boolOptionsLoaded Then frmOptions.LoadSkin
    If boolTagsLoaded Then frmTags.LoadSkin
    If boolDspLoaded = True Then frmDSP.LoadSkin
    frmMain.Image_State_Rep

    With frmMain
        .PicSpectrum.PaintPicture .picNormalMode.Image, 0, 0, .PicSpectrum.ScaleWidth, .PicSpectrum.ScaleHeight, .PicSpectrum.Left, .PicSpectrum.Top, .PicSpectrum.ScaleWidth, .PicSpectrum.ScaleHeight
        .PicSpectrum.Picture = .PicSpectrum.Image
        .PicSpectrumMini.PaintPicture .picMiniMode.Image, 0, 0, .PicSpectrumMini.ScaleWidth, .PicSpectrumMini.ScaleHeight, .PicSpectrumMini.Left, .PicSpectrumMini.Top, .PicSpectrumMini.ScaleWidth, .PicSpectrumMini.ScaleHeight
        .PicSpectrumMini.Picture = .PicSpectrumMini.Image
    End With

End Sub

Public Sub Change_Mask(MiniMask As Boolean, bNormal As Boolean)
'bnormal is used if only mask is changed while the skin remains the same
'On Error Resume Next

    bSlider = False

    Dim FormLeft As Long, FormTop As Long
    Dim newRegion As Long


    If MiniMask = True Then
        bMiniMask = True
        iBorder = (frmMain.Width - frmMain.ScaleWidth)
        frmMain.EqBackpic.Visible = False

        iTitleHeight = (frmMain.Height - frmMain.ScaleHeight)
        frmMain.Width = frmMain.picNormalMode.ScaleWidth * Screen.TwipsPerPixelX + iBorder
        'add space for the border and titlebar
        frmMain.Height = frmMain.picNormalMode.ScaleHeight * Screen.TwipsPerPixelY + iBorder + iTitleHeight
        frmMain.picNormalMode.Visible = False
        frmMain.picMiniMode.Visible = True
        frmMain.Width = frmMain.picMiniMode.Width + iBorder
        frmMain.Height = frmMain.picMiniMode.Height + iBorder + iTitleHeight


        ' The API call requires the address of the region data,
        ' so we pass the first cell in the array. VB passes arrays
        ' ByRef, so here's our address.

        newRegion = ExtCreateRegionByte(ByVal 0&, EdgeRegions(1).DataLength, EdgeRegions(1).RegionData(0))
        SetWindowRgn frmMain.hwnd, newRegion, True
        DeleteObject newRegion



        frmMain.ScrollText(5).CaptionText = sTextScroll
        frmMain.ScrollText(5).ToolTipText = sTextScroll
        If boolVisShow = True Then FrmVisualisation.ScrollText1.CaptionText = sTextScroll

        '// posbar
        'frmMain.slider(3).max = frmMain.slider(0).max
        frmMain.slider(3).Max = CInt(Stream_GetDuration(lCurrentChannel))

        'frmMain.slider(3).value = frmMain.slider(0).value
        '// volbar
        'frmMain.slider(4).Value = frmMain.VolumeNActuaL
        frmMain.ScrollText(5).ToolTipText = sTextScroll


        PeakHeight = GetPrivateProfileInt("MiniMode", "Peakheight", 1, tAppConfig.AppConfig & "Skins\" & tAppConfig.Skin & "\Skin.ini")
        Bar_gap = GetPrivateProfileInt("MiniMode", "Bar_gap", 2, tAppConfig.AppConfig & "Skins\" & tAppConfig.Skin & "\Skin.ini")
        FFT_BANDS = GetPrivateProfileInt("MiniMode", "FFT_BANDS", 22, tAppConfig.AppConfig & "Skins\" & tAppConfig.Skin & "\Skin.ini")
        DRW_BARWIDTH = GetPrivateProfileInt("MiniMode", "DRW_BARWIDTH", 2, tAppConfig.AppConfig & "Skins\" & tAppConfig.Skin & "\Skin.ini")
        DRW_BARSPACE = GetPrivateProfileInt("MiniMode", "DRW_BARSPACE", 1, tAppConfig.AppConfig & "Skins\" & tAppConfig.Skin & "\Skin.ini")
        DRW_PEAKFALL = GetPrivateProfileInt("MiniMode", "DRW_PEAKFALL", 1, tAppConfig.AppConfig & "Skins\" & tAppConfig.Skin & "\Skin.ini")
        COLOR_1 = Read_INI("MiniMode", "Color1", &HCC00, True)
        COLOR_2 = Read_INI("MiniMode", "Color2", &H22A8E1, True)
        COLOR_3 = Read_INI("MiniMode", "Color3", &H122BC, True)

        If bNormal = True Then    'adjust form position so that form is not felt to be displaced
            FormLeft = frmMain.Left + (frmMain.Button(14).Left * Screen.TwipsPerPixelX)
            FormLeft = FormLeft - (frmMain.ButtonMini(7).Left * Screen.TwipsPerPixelX) + (frmMain.ButtonMini(7).Width * Screen.TwipsPerPixelX)
            frmMain.Left = FormLeft
            FormTop = frmMain.Top + (frmMain.Button(14).Top * Screen.TwipsPerPixelY)
            FormTop = FormTop - (frmMain.ButtonMini(7).Top * Screen.TwipsPerPixelY)
            frmMain.Top = FormTop
        End If

    Else

        frmMain.Width = frmMain.picNormalMode.Width + iBorder
        frmMain.Height = frmMain.picNormalMode.Height + iTitleHeight + iBorder

        frmMain.EQ_in_out    'Sets region
        frmMain.picMiniMode.Visible = False
        frmMain.picNormalMode.Visible = True
        frmMain.picNormalMode.Refresh
        frmMain.EqBackpic.Visible = True

        bMiniMask = False    '//very important do not delete

        If boolVisShow = True Then FrmVisualisation.ScrollText1.CaptionText = sTextScroll

        '// posbar
        'frmMain.slider(0).max = frmMain.slider(3).max
        frmMain.slider(0).Max = CInt(Stream_GetDuration(lCurrentChannel))

        'frmMain.slider(4).Value = frmMain.VolumeNActuaL
        '// volbar
        frmMain.slider(1).Value = frmMain.VolumeNActuaL

        frmMain.ScrollText(1).CaptionText = sTextScroll

        PeakHeight = GetPrivateProfileInt("NormalMode", "Peakheight", 1, tAppConfig.AppConfig & "Skins\" & tAppConfig.Skin & "\Skin.ini")
        Bar_gap = GetPrivateProfileInt("NormalMode", "Bar_gap", 2, tAppConfig.AppConfig & "Skins\" & tAppConfig.Skin & "\Skin.ini")
        FFT_BANDS = GetPrivateProfileInt("NormalMode", "FFT_BANDS", 22, tAppConfig.AppConfig & "Skins\" & tAppConfig.Skin & "\Skin.ini")
        DRW_BARWIDTH = GetPrivateProfileInt("NormalMode", "DRW_BARWIDTH", 5, tAppConfig.AppConfig & "Skins\" & tAppConfig.Skin & "\Skin.ini")
        DRW_BARSPACE = GetPrivateProfileInt("NormalMode", "DRW_BARSPACE", 1, tAppConfig.AppConfig & "Skins\" & tAppConfig.Skin & "\Skin.ini")
        DRW_PEAKFALL = GetPrivateProfileInt("NormalMode", "DRW_PEAKFALL", 1, tAppConfig.AppConfig & "Skins\" & tAppConfig.Skin & "\Skin.ini")
        COLOR_1 = Read_INI("NormalMode", "Color1", &HCC00, True)
        COLOR_2 = Read_INI("NormalMode", "Color2", &H22A8E1, True)
        COLOR_3 = Read_INI("NormalMode", "Color3", &H122BC, True)

        If bNormal = True Then
            FormLeft = frmMain.Left + (frmMain.ButtonMini(7).Left * Screen.TwipsPerPixelX)
            FormLeft = FormLeft - (frmMain.Button(14).Left * Screen.TwipsPerPixelX) - (frmMain.Button(14).Width * Screen.TwipsPerPixelX)
            frmMain.Left = FormLeft
            FormTop = frmMain.Top + (frmMain.ButtonMini(7).Top * Screen.TwipsPerPixelY)
            FormTop = FormTop - (frmMain.Button(14).Top * Screen.TwipsPerPixelY)
            frmMain.Top = FormTop
        End If
    End If



End Sub

'+-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-+
'+-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-+

'// procedimento para hacer calkular la maskara normal y la mini
Sub Form_Mini_Normal()
    Dim WinRegion As Long
    Dim Ret As Long

    '//-----------------------------------------------------------------
    '//  MASKARA NORMAL
    '//-----------------------------------------------------------------

    frmMain.picMiniMode.Move 0, 0
    frmMain.picNormalMode.Move 0, 0
    'frmmain.
    '// loaded from file
    If bLoadRegionFile = True Then
        If LoadRegions(EdgeRegions()) = True Then
            Exit Sub
        End If
    End If

    '// First create the region for the bitmap
    WinRegion = MakeRegion(frmMain.picNormalMode)
    '// Get the size needed for the region data buffer
    EdgeRegions(0).DataLength = GetRegionDataLong(WinRegion, 0&, ByVal 0&)

    If EdgeRegions(0).DataLength <> 0 Then
        ' Actually get the data into the buffer - a byte array
        ' of the proper size.
        ' You need 32 bytes more, because the API call attaches
        ' a 32-byte structure called RGNDATAHEADER before the
        ' data itself
        ReDim EdgeRegions(0).RegionData(EdgeRegions(0).DataLength + 32)

        Ret = GetRegionDataByte(WinRegion, EdgeRegions(0).DataLength, EdgeRegions(0).RegionData(0))

    End If

    '//-----------------------------------------------------------------
    '//  MINI MASCARA
    '//-----------------------------------------------------------------
    WinRegion = MakeRegion(frmMain.picMiniMode)
    EdgeRegions(1).DataLength = GetRegionDataLong(WinRegion, 0&, ByVal 0&)

    If EdgeRegions(1).DataLength <> 0 Then
        ReDim EdgeRegions(1).RegionData(EdgeRegions(1).DataLength + 32)
        Ret = GetRegionDataByte(WinRegion, EdgeRegions(1).DataLength, EdgeRegions(1).RegionData(0))
    End If

    WinRegion = MakeRegion(frmMain.picWithoutEq)
    EdgeRegions(2).DataLength = GetRegionDataLong(WinRegion, 0&, ByVal 0&)

    If EdgeRegions(2).DataLength <> 0 Then
        ReDim EdgeRegions(2).RegionData(EdgeRegions(2).DataLength + 32)
        Ret = GetRegionDataByte(WinRegion, EdgeRegions(2).DataLength, EdgeRegions(2).RegionData(0))
    End If

    SaveRegions EdgeRegions()
    DeleteObject WinRegion

End Sub

'=================================================================================
Public Sub SaveRegions(EdgeRegions() As RegionDataType)
    On Error GoTo hell
    Dim i As Long
    Dim FileName As String
    FileName = tAppConfig.AppConfig & "Skins\" & tAppConfig.Skin & "\regions.dat"
    Open FileName For Binary As #1

    For i = 0 To 2
        Put 1, , EdgeRegions(i).DataLength
        Put 1, , EdgeRegions(i).RegionData
    Next

    Close
    Exit Sub
hell:
End Sub

'=================================================================================
Public Function LoadRegions(EdgeRegions() As RegionDataType) As Boolean
' On Error GoTo hell
    Dim i As Long
    Dim FileName As String

    FileName = tAppConfig.AppConfig & "Skins\" & tAppConfig.Skin & "\regions.dat"

    If Dir(FileName) = "" Then Exit Function

    Open FileName For Binary As #1

    For i = 0 To 2
        Get 1, , EdgeRegions(i).DataLength
        ReDim EdgeRegions(i).RegionData(EdgeRegions(i).DataLength + 32)
        Get 1, , EdgeRegions(i).RegionData
    Next

    Close

    LoadRegions = True
    Exit Function
hell:
    Close
    MsgBox err.Description
End Function

'+-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-+
'+-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-+

Public Sub Load_Skins_Menu(SelMenu As String)
'// Procedure to load the skins available that are all the folders
'// EXE in the path of most \ Mp3 Player \ Skins \ and load on the menu frmpopup
'// parameters
'// [SelMenu] -> Menu which is going to be selected

    Dim miNombre As String
    Dim i As Integer

    MiRuta = tAppConfig.AppConfig & "skins\"
    i = 0
    miNombre = Dir(MiRuta, vbDirectory)   ' Retrieves the first entry.
    'frmPopUp.picMenu.Picture = LoadPicture(S & "PicMenu.bmp")
    If miNombre = "" Then
        For i = 0 To frmPopUp.mnuSkinsAdd.Count
            frmPopUp.mnuSkinsAdd(i).Caption = ""
            frmPopUp.mnuSkinsAdd(i).Visible = False
        Next i
        tAppConfig.Skin = "\No skin selected\"
        Exit Sub
    End If

    '/* to see if there are images in the directory
    frmPopUp.fileBmps.Pattern = "*.bmp"
    'i = -1
    Do While miNombre <> ""
        If miNombre <> "." And miNombre <> ".." Then
            ' Performs a bitwise comparison to make sure "miNombre" is a directory.
            If (GetAttr(MiRuta & miNombre) And vbDirectory) = vbDirectory Then
                frmPopUp.fileBmps.Path = MiRuta & miNombre
                If frmPopUp.fileBmps.ListCount > 0 Then
                    i = i + 1

                    If i <> 25 And i >= frmPopUp.mnuSkinsAdd.Count Then Load frmPopUp.mnuSkinsAdd(i)  '// cargar los menus dinamikamente

                    frmPopUp.mnuSkinsAdd(i).Caption = " " & miNombre
                    frmPopUp.mnuSkinsAdd(i).Checked = False
                    frmPopUp.mnuSkinsAdd(i).Visible = True
                    If LCase(miNombre) = LCase(SelMenu) Then frmPopUp.mnuSkinsAdd(i).Checked = True

                End If
            End If
        End If
        miNombre = Dir
    Loop

    If OpcionesMusic.SysMenu = True Then Call InitAppend_Sys_menu(frmMain.hwnd)

End Sub



'+--------------------------------------------------------------------------------------+
'|  CREATE THE IMAGE OF WALLPAPER AS SPECIFIED OPTIONS                    |
'+--------------------------------------------------------------------------------------+

Public Sub CreatePic(picSource As PictureBox, picDestination As PictureBox)
'// Procedure to create  stretch image with the highest possible quality
    Dim hBrush As Long
    Dim hDummyBrush As Long
    Dim lOrigMode As Long
    Dim uBrushOrigPt As POINTAPI
    Dim lWidth As Long
    Dim lHeight As Long
    Dim lLeft As Integer
    Dim lTop As Integer
    picDestination.AutoRedraw = True
    picDestination.cls
    lWidth = picDestination.Width
    lHeight = picDestination.Height
    lLeft = 0
    lTop = 0
    'Set picEdit's stretch mode to halftone (this may cause misalignment of the brush)
    lOrigMode = SetStretchBltMode(picDestination.hDC, STRETCH_HALFTONE)

    'Realign the brush...
    'Get picEdit's brush by selecting a dummy brush into the DC
    hDummyBrush = CreateSolidBrush(0)
    hBrush = SelectObject(picDestination.hDC, hDummyBrush)
    'Reset the brush (This will force windows to realign it when it's put back)
    UnrealizeObject hBrush
    'Set picEdit's brush alignment coordinates to the left-top of the bitmap
    SetBrushOrgEx picDestination.hDC, lLeft, lTop, uBrushOrigPt
    'Now put the original brush back into the DC at the new alignment
    hDummyBrush = SelectObject(picDestination.hDC, hBrush)

    'Stretch the bitmap
    StretchBlt picDestination.hDC, lLeft, lTop, lWidth, lHeight, _
               picSource.hDC, 0, 0, picSource.Width, picSource.Height, vbSrcCopy

    'Set the stretch mode back to it's original mode
    SetStretchBltMode picDestination.hDC, lOrigMode

    'Reset the original alignment of the brush...
    'Get picEdit's brush by selecting the dummy brush back into the DC
    hBrush = SelectObject(picDestination.hDC, hDummyBrush)
    'Reset the brush (This will force windows to realign it when it's put back)
    UnrealizeObject hBrush
    'Set the brush alignment back to the original coordinates
    SetBrushOrgEx picDestination.hDC, uBrushOrigPt.X, uBrushOrigPt.Y, uBrushOrigPt
    'Now put the original brush back into picEdit's DC at the original alignment
    hDummyBrush = SelectObject(picDestination.hDC, hBrush)
    'Get rid of the dummy brush
    DeleteObject hDummyBrush
End Sub

'+-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-+
'+-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-+


'+--------------------------------------------------------------------------------------+
'|   CREATE THE IMAGE OF WALLPAPER AND PUT ON THE DESKTOP                         |
'+--------------------------------------------------------------------------------------+



'+-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-+
'+-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-+





