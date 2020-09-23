Attribute VB_Name = "mConfig"
Option Explicit
'+-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-+
'| INFORMATION FOR PROCESSOR MEMORY                                                    |
'+-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-+
Public Type MEMORYSTATUS
    dwLength As Long
    dwMemoryLoad As Long
    dwTotalPhys As Long
    dwAvailPhys As Long
    dwTotalPageFile As Long
    dwAvailPageFile As Long
    dwTotalVirtual As Long
    dwAvailVirtual As Long
End Type

Public sInitialDir As String

Declare Sub GlobalMemoryStatus Lib "kernel32" (lpBuffer As MEMORYSTATUS)

Public Declare Function SetActiveWindow Lib "user32" (ByVal hwnd As Long) As Long

'+-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-+
'| OPERATING SYSTEM VERSION INFORMATION                                                  |
'+-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-+

Private Type OSVERSIONINFO
    OSVSize As Long
    dwVerMajor As Long
    dwVerMinor As Long
    dwBuildNumber As Long
    PlatformID As Long
    szCSDVersion As String * 128
End Type
Private Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" _
                                      (lpVersionInformation As OSVERSIONINFO) As Long


'+-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-+
'| EXECUTE APPLICATIONS WITH GIVEN PARAMETERS                                       |
'+-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-+

'+-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-+
'| FORM DRAG                                                           |
'+-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-+

Declare Sub ReleaseCapture Lib "user32" ()

'+-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-+
'| APIS FOR ALWAYS PUTTING ON TOP THE FORM                                      |
'+-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-+

Public Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Public Const SWP_FRAMECHANGED = &H20
Public Const SWP_NOMOVE = &H2
Public Const SWP_NOSIZE = &H1
Public Const SWP_NOZORDER = &H4
Public Const SWP_SHOWWINDOW = &H40
Public Const SWP_NOOWNERZORDER = &H200      '  No usar el orden Z del propietario
Public Const HWND_TOP = 0
Public Const HWND_TOPMOST = -1
Public Const HWND_NOTOPMOST = -2

'+-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-+
'| TEXT FOR MOVING PICTURES
'+-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-+
Public Declare Function DrawText Lib "user32" Alias "DrawTextA" (ByVal hDC As Long, ByVal lpStr As String, ByVal nCount As Long, lpRect As RECT, ByVal wFormat As Long) As Long

Const DT_BOTTOM As Long = &H8
Public Const DT_CALCRECT As Long = &H400
Const DT_EXPANDTABS As Long = &H40
Const DT_EXTERNALLEADING As Long = &H200
Const DT_LEFT As Long = &H0
Const DT_NOCLIP As Long = &H100
Const DT_NOPREFIX As Long = &H800
Const DT_RIGHT As Long = &H2
Public Const DT_SINGLELINE As Long = &H20
Const DT_TABSTOP As Long = &H80
Const DT_TOP As Long = &H0
Const DT_VCENTER As Long = &H4
Public Const DT_WORDBREAK As Long = &H10

'+-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-+
'|    Declaration for Layered Windows (Windows 2000 and above)                 |
'|    API'S TO MAKE THE FORM TRANSPARENT                                            |
'+-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-+


Public Const WS_EX_LAYERED As Long = &H80000
Public Const LWA_ALPHA As Long = &H2
Public Const GWL_EXSTYLE = (-20)
Public Const RDW_INVALIDATE = &H1
Public Const RDW_ERASE = &H4
Public Const RDW_ALLCHILDREN = &H80
Public Const RDW_FRAME = &H400

'
Public Declare Function SetLayeredWindowAttributes Lib "user32" (ByVal hwnd As Long, ByVal crKey As Long, ByVal bAlpha As Long, ByVal dwFlags As Long) As Long
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Public Declare Function RedrawWindow2 Lib "user32" Alias "RedrawWindow" (ByVal hwnd As Long, ByVal lprcUpdate As Long, ByVal hrgnUpdate As Long, ByVal fuRedraw As Long) As Long


'+-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-+
'| APIS TO READ THE SETTINGS FILE. INI OR OTHER
'+-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-+
Declare Function GetPrivateProfileString Lib "kernel32" _
                                         Alias "GetPrivateProfileStringA" (ByVal lpApplicationName _
                                                                           As String, lpKeyName As Any, ByVal lpDefault As String, _
                                                                           ByVal lpRetunedString As String, ByVal nSize As Long, _
                                                                           ByVal lpFileName As String) As Long

Declare Function WritePrivateProfileString Lib "kernel32" _
                                           Alias "WritePrivateProfileStringA" (ByVal lpApplicationName _
                                                                               As String, ByVal lpKeyName As Any, ByVal lpString As Any, _
                                                                               ByVal lplFileName As String) As Long

Public Enum Sel_Option
    PathExe = 0
    PathSkin = 1
End Enum

Public Function Read_INI(Section As String, Value As String, Default As Variant, Optional IsColor As Boolean = False, Optional ConfigurationMusic As Boolean = False, Optional FilePath As String) As Variant
'// Funcion para leer configuraciones del INI
'// Parametros
'// [Section] -> sECTION IN INI FILE [Configuration]
'// [Value] -> vALUE OF KEY IN INI FILE
'// [Default] -> Return value if the value is missiong

    Dim ColorArr As Variant
    Dim str As String
    ' if ini file of music modefication is to be retrieved then ConfigurationMusic = True
    'filepath is taken automatically
    If ConfigurationMusic = True Then
        str = String(255, Chr(0))
        ' FUNCTION BELOW 2nd para. is a function that changes the value of 'str'giving it a value of referred keyname in section in ini file and returns its length
        ' if str is "No_TA" i.e. not available i.e. key does 't has any value
        str = Left(str, GetPrivateProfileString(Section, ByVal Value, "NO_TA", str, Len(str), tAppConfig.AppPath & App.EXEName & ".ini"))
        If str = "NO_TA" Then    '
            ' if no key value is found it returns default as specified in function call
            Read_INI = Trim(Default)
        Else
            Read_INI = Trim(str)
        End If
        Exit Function    'exit work is over
    End If
    ' if specific ini file is to be retrieved then filepath is given
    If Trim(FilePath) <> "" Then
        str = String(255, Chr(0))
        str = Left(str, GetPrivateProfileString(Section, ByVal Value, "NO_TA", str, Len(str), FilePath))
        If str = "NO_TA" Then    '
            Read_INI = Trim(Default)
        Else
            Read_INI = Trim(str)
        End If
        Exit Function    'exit work is over
    End If
    ' if ini file of color i.e. colour of specific item stoered in skin ini is to be retrieved then color = True
    If IsColor = True Then    ' is a color
        str = String(255, Chr(0))
        str = Left(str, GetPrivateProfileString(Section, ByVal Value, "NO_TA", str, Len(str), tAppConfig.AppConfig & "skins\" & tAppConfig.Skin & "\" & "Skin.ini"))

        If str = "NO_TA" Then    ' si no encuentra la clave
            Read_INI = Default
        Else
            ColorArr = Split(str, ",")
            If UBound(ColorArr) <> 2 Then    ' if the wrong key
                Read_INI = Default
            Else
                ' function returns rgb type long since is variant
                Read_INI = RGB(ColorArr(0), ColorArr(1), ColorArr(2))
            End If
        End If
    Else    '(color=false)
        str = String(255, Chr(0))
        str = Left(str, GetPrivateProfileString(Section, ByVal Value, "NO_TA", str, Len(str), tAppConfig.AppConfig & "skins\" & tAppConfig.Skin & "\" & "Skin.ini"))
        If str = "NO_TA" Then    ' if it does't find the key
            Read_INI = Trim(Default)
        Else
            Read_INI = Trim(str)
        End If
    End If
End Function

Public Function Read_Config_Button(Objeto As Object, Section As String, Value As String, Default As Variant, Optional iniPath As String = "") As Boolean
    On Error Resume Next

    Dim str As String
    Dim arry() As String
    Dim Path As String
    If iniPath = "" Then
        Path = tAppConfig.AppConfig & "Skins\" & tAppConfig.Skin _
               & "\Skin.ini"
    Else
        Path = iniPath
    End If
    str = String(255, Chr(0))
    str = Left(str, GetPrivateProfileString(Section, ByVal Value, "NO_TA", _
                                            str, Len(str), Path))

    If str = "NO_TA" Then str = Default
    'to split numbers of form 12,23,34,2 separated by comma in numbers that are stored in array
    arry = Split(str, ",")

    '// Slider Pos or Vol if button is for slider or scrollbar x,y,w,h,pos are there in ini file
    '// check for 5th array element i.e. pos
    If UBound(arry) = 4 Then
        If UCase(arry(4)) = "V" Then    '// Vertical position
            Objeto.Position = 0
        Else    '// Horizontal
            Objeto.Position = 1
        End If
    End If

    '// Button normal
    If UBound(arry) = 4 Or UBound(arry) = 3 Then
        'Read_INI = Str
        'set the dimensions of buttons acc to the size given in skin ini file
        'objecto is the object being passed in the function eg. formmain .button(0) which is a picture box
        Objeto.Left = arry(0)
        Objeto.Top = arry(1)
        Objeto.Width = arry(2)
        Objeto.Height = arry(3)
    End If


End Function


Public Function Write_INI(Section As String, KeyName As String, KeyValue As Variant, FilePath As String) As Boolean
    Dim Ret As Long
    Ret = WritePrivateProfileString(Section, KeyName, CStr(KeyValue), FilePath)
    If Ret = 0 Then
        Write_INI = True
    Else
        Write_INI = False
    End If
End Function

Sub Load_Settings_INI(bolNormal As Boolean)
    Dim strRes As Variant
    Dim arryFormat() As String
    Dim i As Integer
    Dim strKeyQuery As Variant
    On Error Resume Next
    strKeyQuery = vbNullString
    ' t for type
    tAppConfig.AppPath = App.Path
    'check for directory or drive
    If Right(tAppConfig.AppPath, 1) <> "\" Then tAppConfig.AppPath = tAppConfig.AppPath & "\"
    ' to search fromm application config ini file the path of config file
    tAppConfig.AppConfig = Read_INI("Configuration", "AppConfiguration", tAppConfig.AppPath & "MMp3Player\", , True)

    If Right(tAppConfig.AppConfig, 1) <> "\" Then tAppConfig.AppConfig = tAppConfig.AppConfig & "\"
    If Dir(tAppConfig.AppConfig, vbDirectory) = "" Then tAppConfig.AppConfig = tAppConfig.AppPath & "MMp3Player\"

    '// Multiples instancias
    strRes = Read_INI("Configuration", "MulInstances", 0, , True)
    If CBool(strRes) = True Then OpcionesMusic.Instancias = True

    '// Mostrar Splash Screen
    strRes = Read_INI("Configuration", "SplashScreen", 1, , True)
    If CBool(strRes) = True Then
        'frmSplash.lblSplash(0).Caption = " Loading configuration..."
        frmSplash.Show
        OpcionesMusic.Splash = True
    End If

    '// check for system menu
    strRes = Read_INI("Configuration", "SysMenu", 1, , True)
    If CBool(strRes) = True Then OpcionesMusic.SysMenu = True

    '// regions from file
    strRes = Read_INI("Configuration", "LoadRegionFile", 0, , True)
    If CBool(strRes) = True Then bLoadRegionFile = True

    '// Name of Skin
    strRes = Read_INI("Configuration", "Skin", "", , True)

    If Trim(strRes) = "" Or Dir(tAppConfig.AppConfig & "Skins\" & strRes, vbDirectory) = "" Then
        Load_Skins_Menu LCase(tAppConfig.Skin)
        Read_Config_Skin
        Change_Skin Trim(frmPopUp.mnuSkinsAdd(1).Caption)
        frmPopUp.mnuSkinsAdd(1).Checked = True
        Form_Mini_Normal
    Else
        Read_Config_Skin
        Change_Skin Trim(strRes)    '// Change skin, position of controls
        Form_Mini_Normal    '// if you set the form irregular zones
        Load_Skins_Menu LCase(tAppConfig.Skin)    '// load the Skins menu and select the current
    End If

    '// update in comnfig file Estado de la maskara mini - normal
    strRes = Read_INI("Configuration", "Mini", 0, , True)
    If CBool(strRes) = True Then bMiniMask = True
    If bolSplashScreen = False Then
        If bMiniMask = True Then
            Change_Mask True, False    ' 2nd false for positioning of form
        Else
            Change_Mask False, False
        End If
    End If

    '// Moving Forms
    strRes = Read_INI("Configuration", "MX", 3530, , True)
    If IsNumeric(strRes) = False Then strRes = 0
    frmMain.Left = CInt(strRes)


    strRes = Read_INI("Configuration", "MY", 2065, , True)
    If IsNumeric(strRes) = False Then strRes = 0
    frmMain.Top = CInt(strRes)

    'Eq in/out
    strRes = Read_INI("Configuration", "EqDisplay", True, , True)
    If CBool(strRes) And bMiniMask = False Then Call frmMain.Button_Click(24)


    'EqFreeMove Button
    strRes = Read_INI("Configuration", "EqFreeMove", False, , True)
    bEQFreeMove = CBool(strRes)
    frmMain.Button(17).Selected = bEQFreeMove

    Dim s, T As String
    Dim j As Integer
    Dim sValue
    i = 0
    'DoEvents
    'DoEvents
    frmMain.ListEq.Clear
    Do
        T = ""
        s = ""
        s = Read_INI("equalizer_" & i, "name", "", , , tAppConfig.AppConfig & "Settings\Equalizer.eql")
        For j = 0 To 9
            sValue = Read_INI("equalizer_" & i, "eq" & j, 0, , , tAppConfig.AppConfig & "Settings\Equalizer.eql")
            If IsNumeric(sValue) Then T = T + str(sValue) + ","
        Next j

        Call frmMain.ListEq.AddItem(s + "." + T)
        i = i + 1
    Loop While s <> ""



    '// Mostrar biblioteca multemedia
    strRes = Read_INI("Configuration", "MediaLibraryShow", 0, , True)
    boolMediaLibraryShow = CBool(strRes)

    frmPopUp.mnuLibrary.Checked = boolMediaLibraryShow
    '// si no esta seleccionado el splash screen mostrar los form ahora



    strRes = Read_INI("Configuration", "PlstDisplay", 1, , True)
    frmMain.Button(26).Selected = CBool(strRes)



    ' DoEvents

    strRes = Read_INI("Configuration", "VisDisplay", 1, , True)
    boolVisShow = CBool(strRes)
    If boolVisShow = True Then
        boolVisLoaded = True:
        frmPopUp.mnuShowVis.Checked = True
        Call CheckMenu(12, frmPopUp.mnuShowVis.Checked)
        Load FrmVisualisation
        FrmVisualisation.Visible = True
    End If


    If CInt(strRes) < 0 Or CInt(strRes) > 3 Or IsNumeric(strRes) = False Then strRes = 0

    '//Put correct values if the file changed
    If strRes = 0 Then
        OpcionesMusic.NoAlteraR = True
    ElseIf strRes = 1 Then
        OpcionesMusic.Mosaico = True
    ElseIf strRes = 2 Then
        OpcionesMusic.Centrar = True
    Else
        OpcionesMusic.Expander = True
    End If

    '// Visualizacion
    strRes = Read_INI("Configuration", "Visualization", 1, , True)

    If CInt(strRes) < 0 Or CInt(strRes) > 2 Or IsNumeric(strRes) = False Then strRes = 1

    If strRes = 0 Then
        frmPopUp.mnuSpecNone.Checked = True
    ElseIf strRes = 1 Then
        frmPopUp.mnuSpecBars.Checked = True
    ElseIf strRes = 2 Then
        frmPopUp.mnuSpecOsc.Checked = True
    End If

    '// format scroll
    sFormatPlayList = Trim(Read_INI("Configuration", "FormatPlayList", "%A - %S", , True))

    '// format scroll
    sFormatScroll = Trim(Read_INI("Configuration", "FormatScroll", "%S - %A (%T)", , True))

    '// scroll caption type
    strRes = Read_INI("Configuration", "ScrollType", 0, , True)
    If CInt(strRes) < 0 Or CInt(strRes) > 1 Or IsNumeric(strRes) = False Then strRes = 0
    iScrollType = CInt(strRes)
    frmMain.ScrollText(1).ScrollType = iScrollType
    frmMain.ScrollText(5).ScrollType = iScrollType

    '// scroll caption vel
    strRes = Read_INI("Configuration", "ScrollVel", 130, , True)
    If CInt(strRes) < 100 Or CInt(strRes) > 1000 Or IsNumeric(strRes) = False Then strRes = 130
    iScrollVel = CInt(strRes)
    frmMain.ScrollText(1).ScrollVelocity = iScrollVel
    frmMain.ScrollText(5).ScrollVelocity = iScrollVel

    '// Crossfade entre Tracks
    strRes = Read_INI("Configuration", "CrossfadeEnabled", 0, , True)
    If CBool(strRes) = True Then bCrossFadeEnabled = True

    '// Sprectrum refresh rate (interval of timer)
    strRes = Read_INI("Configuration", "SpectrumRefreshRate", 50, , True)
    If CInt(strRes) < 100 Or CInt(strRes) > 1000 Or IsNumeric(strRes) = False Then strRes = 50
    iSpectrum_refreshRate = CInt(strRes)

    '// Crossfade entre Tracks
    strRes = Read_INI("Configuration", "CrossfadeTrack", 100, , True)
    If CInt(strRes) < 0 Or CInt(strRes) > 400 Or IsNumeric(strRes) = False Then strRes = 100
    iCrossfadeTrack = CInt(strRes)

    '// Volume mouse Scroll change rate
    strRes = Read_INI("Configuration", "PosScrollChange", 5, , True)
    If CInt(strRes) < 0 Or CInt(strRes) > 100 Or IsNumeric(strRes) = False Then strRes = 5
    iPosScrollChange = CInt(strRes)

    '// Track seek on mouse scroll
    strRes = Read_INI("Configuration", "VolScrollChange", 5, , True)
    If CInt(strRes) < 0 Or CInt(strRes) > 100 Or IsNumeric(strRes) = False Then strRes = 5
    iVolScrollChange = CInt(strRes)

    '// Crossfade para detener
    strRes = Read_INI("Configuration", "CrossfadeStop", 100, , True)
    If CInt(strRes) < 0 Or CInt(strRes) > 400 Or IsNumeric(strRes) = False Then strRes = 100
    iCrossfadeStop = CInt(strRes)

    strRes = Read_INI("Configuration", "OptPlstCase", 0, , True)
    If CInt(strRes) < 0 Or CInt(strRes) > 2 Or IsNumeric(strRes) = False Then strRes = 0
    iOptPlstCase = CInt(strRes)

    frmMain.EnableCrossfade (bCrossFadeEnabled)

    '// add files at search
    strRes = Read_INI("Configuration", "AddFiles", 0, , True)
    If CBool(strRes) = True Then bAddFiles = True

    '// play en el comienzo
    strRes = Read_INI("Configuration", "PlayStarting", 1, , True)
    If CBool(strRes) = True Then bPlayStarting = True

    '// check proporcional
    strRes = Read_INI("Configuration", "Proportional", 0, , True)
    If CBool(strRes) = True Then OpcionesMusic.Proporcional = True

    '// check Directorio
    strRes = Read_INI("Configuration", "Directory", 0, , True)
    If CBool(strRes) = True Then OpcionesMusic.Directorio = True

    '// check show task bar
    strRes = Read_INI("Configuration", "TaskBar", 1, , True)
    If CBool(strRes) = True Then OpcionesMusic.TaskBar = True

    '//----------------------------------------------------------------------------------
    '// system tray icon
    strRes = Read_INI("Configuration", "SysTray", 0, , True)
    If CBool(strRes) = True Then
        OpcionesMusic.SysTray = True
        frmPopUp.vkSysTrayIcon(0).AddToTray (0)
    End If

    strRes = Read_INI("Configuration", "SysTrayNext", 0, , True)
    If CBool(strRes) = True Then
        PlayerTrayIcon.Next = True
        frmPopUp.vkSysTrayIcon(5).AddToTray (5)
    End If

    strRes = Read_INI("Configuration", "SysTrayStop", 0, , True)
    If CBool(strRes) = True Then
        PlayerTrayIcon.Stop = True
        frmPopUp.vkSysTrayIcon(4).AddToTray (4)
    End If

    strRes = Read_INI("Configuration", "SysTrayPause", 0, , True)
    If CBool(strRes) = True Then
        PlayerTrayIcon.Pause = True
        frmPopUp.vkSysTrayIcon(3).AddToTray (3)
    End If

    strRes = Read_INI("Configuration", "SysTrayPlay", 0, , True)
    If CBool(strRes) = True Then
        PlayerTrayIcon.Play = True
        frmPopUp.vkSysTrayIcon(2).AddToTray (2)
    End If

    strRes = Read_INI("Configuration", "SysTrayPrevious", 0, , True)
    If CBool(strRes) = True Then
        PlayerTrayIcon.Previous = True
        frmPopUp.vkSysTrayIcon(1).AddToTray (1)
    End If
    '//----------------------------------------------------------------------------------

    '// play files format
    strRes = Read_INI("Configuration", "FileType", "1;0;0;0", , True)

    Dim arryFiles(3) As String
    arryFiles(0) = "mp3"
    arryFiles(1) = "wma"
    arryFiles(2) = "wav"
    arryFiles(3) = "ogg"

    arryFormat = Split(strRes, ";", , vbTextCompare)

    For i = 0 To UBound(arryFormat)
        If CBool(arryFormat(i)) = True Then
            If i <= UBound(arryFiles) Then strPathern = strPathern & "*." & arryFiles(i) & ";"
        End If
    Next i

    sFileType = strRes

    '// Trasparencia del form
    strRes = Read_INI("Configuration", "Alpha", 100, , True)
    If strRes < 10 Or strRes > 100 Then strRes = 100
    OpcionesMusic.Alpha = strRes
    Make_Transparent frmMain.hwnd, OpcionesMusic.Alpha    '// Poner Trasparente
    Make_Transparent frmPLST.hwnd, OpcionesMusic.Alpha    '// Poner Trasparente
    'Make_Transparent FrmVisualisation.hWnd, OpcionesMusic.Alpha
    For i = 0 To 9
        If Left(frmPopUp.mnuAlpha(i).Caption, Len(frmPopUp.mnuAlpha(i).Caption) - 1) = OpcionesMusic.Alpha Then
            frmPopUp.mnuAlpha(i).Checked = True
            frmPopUp.mnuAlphaPer.Caption = "Custom..."
            frmPopUp.mnuAlphaPer.Checked = False
            Exit For
        Else
            frmPopUp.mnuAlphaPer.Caption = "Custom " & " [ " & OpcionesMusic.Alpha & "% ]"
            frmPopUp.mnuAlphaPer.Checked = True
        End If
    Next i

    '// Olways on top
    strRes = Read_INI("Configuration", "AlwaysTop", 0, , True)
    If CBool(strRes) = True Then OpcionesMusic.SiempreTop = True

    '// Ajustar Volumen
    strRes = Read_INI("Configuration", "Volume", 255, , True)
    If strRes < 0 Or strRes > 255 Then strRes = 255
    frmMain.slider(1).Value = CInt(strRes)
    frmMain.slider(4).Value = CInt(strRes)
    frmMain.VolumeNActuaL = CInt(strRes)

    '// load lenguaje y cambiarlo
    'strRes = Read_INI("Configuration", "Language", "English", , True)
    'OpcionesMusic.Language = strRes
    'Load_Language OpcionesMusic.Language

    '// -------------------------------------------------------------------------------
    If bolNormal = True Then    '// if it is loaded normally

        '---------------------------------------------------------------------------------------
        'Do while you read something in the ini file.
        frmPopUp.fileBmps.Pattern = strPathern

        '//Reproduced previous album
        strRes = Read_INI("Configuration", "AlbumPlaying", 1, , True)


        '// Playing track number of previous
        strRes = Read_INI("Configuration", "TrackNumber", 0, , True)


        strRes = Read_INI("Configuration", "Intro", 0, , True)
        If CBool(strRes) = True Then frmMain.Intro

        strRes = Read_INI("Configuration", "Mute", 0, , True)
        If CBool(strRes) = True Then frmMain.Player_Mute

        strRes = Read_INI("Configuration", "Repeat", 0, , True)
        If CBool(strRes) = True Then PlayerLoop = True: frmMain.Player_Repeat (PlayerLoop)

        '// The album Random Order
        strRes = Read_INI("Configuration", "RandomizeAlbum", 0, , True)
        If CBool(strRes) = True Then PlayerLoop = True: frmMain.Player_Repeat (PlayerLoop)

    End If

    '===============================================================================
    ' EQUALIZER

End Sub

Sub Save_Settings_INI(Optional Normal As Boolean = False)
    Dim Fnum As Integer, i As Integer
    Dim ArchivoINI As String
    Dim intClave As Integer
    Dim INIcheck As String
    On Error Resume Next

    '// delete systray icons
    If Normal = True Then
    End If

    On Error GoTo BITCH

    ArchivoINI = tAppConfig.AppPath & App.EXEName & ".ini"

    '// Check attributes
    INIcheck = Dir(ArchivoINI, vbNormal + vbSystem + vbHidden + vbReadOnly + vbArchive)

    '// If you are not doing it ...
    If INIcheck = "" Then
        Fnum = FreeFile  '// random number to assign to the file
        Open ArchivoINI For Output As Fnum
        Close
        'SetAttr ArchivoINI, vbHidden + vbSystem
    End If

    Write_INI "Configuration", "AppConfiguration", tAppConfig.AppConfig, ArchivoINI
    Write_INI "Configuration", "SplashScreen", OpcionesMusic.Splash, ArchivoINI
    Write_INI "Configuration", "MulInstances", OpcionesMusic.Instancias, ArchivoINI
    Write_INI "Configuration", "OptPlstCase", iOptPlstCase, ArchivoINI
    Write_INI "Configuration", "Skin", tAppConfig.Skin, ArchivoINI
    Write_INI "Configuration", "LoadRegionFile", bLoadRegionFile, ArchivoINI
    Write_INI "Configuration", "MX", frmMain.Left, ArchivoINI
    Write_INI "Configuration", "MY", frmMain.Top, ArchivoINI
    Write_INI "Configuration", "PX", frmPLST.Left, ArchivoINI
    Write_INI "Configuration", "PY", frmPLST.Top, ArchivoINI
    Write_INI "Configuration", "PW", frmPLST.Width, ArchivoINI
    Write_INI "Configuration", "PH", frmPLST.Height, ArchivoINI
    Write_INI "Configuration", "LX", frmLibrary.Left, ArchivoINI
    Write_INI "Configuration", "LY", frmLibrary.Top, ArchivoINI
    Write_INI "Configuration", "LW", frmLibrary.Width, ArchivoINI
    Write_INI "Configuration", "LH", frmLibrary.Height, ArchivoINI

    Write_INI "Configuration", "MediaLibraryShow", boolMediaLibraryShow, ArchivoINI

    Write_INI "Configuration", "Volume", frmMain.slider(1).Value, ArchivoINI
    Write_INI "Configuration", "Mini", bMiniMask, ArchivoINI

    If OpcionesMusic.NoAlteraR = True Then
        intClave = 0
    ElseIf OpcionesMusic.Mosaico = True Then
        intClave = 1
    ElseIf OpcionesMusic.Centrar = True Then
        intClave = 2
    Else
        intClave = 3
    End If

    Write_INI "Configuration", "Wallpaper", intClave, ArchivoINI

    If frmPopUp.mnuSpecNone.Checked = True Then
        intClave = 0
    ElseIf frmPopUp.mnuSpecBars.Checked = True Then
        intClave = 1
    ElseIf frmPopUp.mnuSpecOsc.Checked = True Then
        intClave = 2
    End If
    Write_INI "Configuration", "Visualization", intClave, ArchivoINI
    Write_INI "Configuration", "Proportional", OpcionesMusic.Proporcional, ArchivoINI
    Write_INI "Configuration", "Directory", OpcionesMusic.Directorio, ArchivoINI
    'Write_INI "Configuration", "Language", OpcionesMusic.Language, ArchivoINI
    Write_INI "Configuration", "FileType", sFileType, ArchivoINI
    Write_INI "Configuration", "FormatPlayList", sFormatPlayList, ArchivoINI
    Write_INI "Configuration", "FormatScroll", sFormatScroll, ArchivoINI
    Write_INI "Configuration", "ScrollType", iScrollType, ArchivoINI
    Write_INI "Configuration", "ScrollVel", iScrollVel, ArchivoINI
    Write_INI "Configuration", "CrossfadeEnabled", CBool(bCrossFadeEnabled), ArchivoINI
    Write_INI "Configuration", "CrossfadeStop", iCrossfadeStop, ArchivoINI
    Write_INI "Configuration", "SpectrumRefreshRate", iSpectrum_refreshRate, ArchivoINI
    Write_INI "Configuration", "PosScrollChange", iPosScrollChange, ArchivoINI
    Write_INI "Configuration", "VolScrollChange", iVolScrollChange, ArchivoINI
    Write_INI "Configuration", "PlayStarting", bPlayStarting, ArchivoINI
    Write_INI "Configuration", "Alpha", OpcionesMusic.Alpha, ArchivoINI
    Write_INI "Configuration", "AlwaysTop", OpcionesMusic.SiempreTop, ArchivoINI
    Write_INI "Configuration", "TaskBar", OpcionesMusic.TaskBar, ArchivoINI
    Write_INI "Configuration", "SysMenu", OpcionesMusic.SysMenu, ArchivoINI
    Write_INI "Configuration", "SysTray", OpcionesMusic.SysTray, ArchivoINI
    Write_INI "Configuration", "SysTrayPrevious", PlayerTrayIcon.Previous, ArchivoINI
    Write_INI "Configuration", "SysTrayPlay", PlayerTrayIcon.Play, ArchivoINI
    Write_INI "Configuration", "SysTrayPause", PlayerTrayIcon.Pause, ArchivoINI
    Write_INI "Configuration", "SysTrayStop", PlayerTrayIcon.Stop, ArchivoINI
    Write_INI "Configuration", "SysTrayNext", PlayerTrayIcon.Next, ArchivoINI
    Write_INI "Configuration", "Intro", frmPopUp.mnuIntro.Checked, ArchivoINI
    Write_INI "Configuration", "Mute", frmPopUp.mnuSilencio.Checked, ArchivoINI
    Write_INI "Configuration", "Repeat", frmPopUp.mnuRepeatTrack.Checked, ArchivoINI
    'Write_INI "Configuration", "RandomizeCollection", frmPopUp.mnuAleatorioTodaColec.Checked, ArchivoINI
    'Write_INI "Configuration", "RandomizeAlbum", frmPopUp.mnuAleatorioActAlbum.Checked, ArchivoINI
    'Write_INI "Configuration", "AlbumPlaying", intActiveAlbum, ArchivoINI
    Write_INI "Configuration", "TrackNumber", CurrentTrack_Index, ArchivoINI
    Write_INI "Configuration", "IndexVis", IndexVisualization, ArchivoINI
    Write_INI "Configuration", "AddFiles", bAddFiles, ArchivoINI
    Write_INI "Configuration", "EqDisplay", CBool(frmMain.Button(24).Selected), ArchivoINI
    Write_INI "Configuration", "PlstDisplay", CBool(frmMain.Button(26).Selected), ArchivoINI
    Write_INI "Configuration", "VisDisplay", CBool(boolVisShow), ArchivoINI
    Write_INI "Configuration", "EqFreeMove", CBool(bEQFreeMove), ArchivoINI

    '===============================================================================
    ' EQUALIZER
    Write_INI "Equalizer", "Enabled", CBool(frmMain.Button(16).Selected), ArchivoINI
    For i = 0 To 9
        Write_INI "Equalizer", "EQ_" & i, frmMain.Eq_SliderCtrl(i).Value, ArchivoINI
    Next i

    '===============================================================================
    ' SOUND EFFECTS



    '===============================================================================
    ' ALBUMS
    '// Can store section for reproducing the actual albums

    Exit Sub
BITCH:

End Sub

Public Sub Always_on_Top()

    Const flag As Long = SWP_NOMOVE Or SWP_NOSIZE
    'Const flag As Long = SWP_SHOWWINDOW
    If OpcionesMusic.SiempreTop = True Then
        If boolVisLoaded = True Then SetWindowPos FrmVisualisation.hwnd, HWND_TOPMOST, 0, 0, 0, 0, flag
        If boolOptionsLoaded = True Then SetWindowPos frmOptions.hwnd, HWND_TOPMOST, 0, 0, 0, 0, flag
        If boolSearchShow = True Then SetWindowPos frmSearch.hwnd, HWND_TOPMOST, 0, 0, 0, 0, flag
        If boolDspShow = True Then SetWindowPos frmDSP.hwnd, HWND_TOPMOST, 0, 0, 0, 0, flag
        SetWindowPos frmPLST.hwnd, HWND_TOPMOST, 0, 0, 0, 0, flag
        SetWindowPos frmMain.hwnd, HWND_TOPMOST, 0, 0, 0, 0, flag
        SetWindowPos frmPopUp.hwnd, HWND_TOPMOST, 0, 0, 0, 0, flag
    Else
        SetWindowPos frmPopUp.hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, flag
        SetWindowPos frmMain.hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, flag
        SetWindowPos frmPLST.hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, flag
        If boolVisLoaded = True Then SetWindowPos FrmVisualisation.hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, flag
        If boolOptionsLoaded = True Then SetWindowPos frmOptions.hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, flag
        If boolSearchShow = True Then SetWindowPos frmSearch.hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, flag
        If boolDspShow = True Then SetWindowPos frmDSP.hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, flag

    End If

End Sub

'+----------------------------------------------------------------------------------------+
'|             TRASPARENCIA                                                               |
'+----------------------------------------------------------------------------------------+

Sub Make_Transparent(lhWnd As Long, Percentage As Integer)
    On Error GoTo hell
    '// transparent procedure for the forms
    '// parameteters
    '// [LHwnD] -> form's handle
    '// [Percentage] -> pus que va ser el ...che Percentage

    '// only work with win 2000 and later

    Dim OSV As OSVERSIONINFO

    '/* Get OS compatability flag
    OSV.OSVSize = Len(OSV)
    If GetVersionEx(OSV) <> 1 Then Exit Sub

    If OSV.PlatformID = 1 And OSV.dwVerMinor >= 10 Then Exit Sub    '/* Win 98/ME
    If OSV.PlatformID = 2 And OSV.dwVerMajor >= 5 Then    '/* Win 2000/XP
        Call SetWindowLong(lhWnd, GWL_EXSTYLE, GetWindowLong(lhWnd, GWL_EXSTYLE) Or WS_EX_LAYERED)
        Call SetLayeredWindowAttributes(lhWnd, 0, (Percentage * 255) / 100, LWA_ALPHA)
    End If
    Exit Sub
hell:
End Sub


'+-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-+
'|  PROCEDURE TO Drag form with Mouse             |
'+-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-+
Public Sub FormDrag(TheForm As Form)
    ReleaseCapture
    Call SendMessage(TheForm.hwnd, &HA1, 2, 0&)
End Sub
'+-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-+

Public Function cRead_INI(Section As String, Value As String, Default As Variant, FilePath As String) As Long
    Dim ColorArr As Variant
    Dim str As String
    str = String(255, Chr(0))
    str = Left(str, GetPrivateProfileString(Section, ByVal Value, "NO_TA", str, Len(str), FilePath))
    If str = "NO_TA" Then
        cRead_INI = Default
    Else
        ColorArr = Split(str, ",")
        If UBound(ColorArr) <> 2 Then
            cRead_INI = Default
        Else
            ' function returns rgb type long
            cRead_INI = RGB(ColorArr(0), ColorArr(1), ColorArr(2))
        End If
    End If
End Function

