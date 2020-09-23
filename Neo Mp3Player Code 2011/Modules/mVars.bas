Attribute VB_Name = "mStart"
Option Explicit
'+-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-+
'|   VARIABLES UTILIZADAS PARA TODO EL PROGRAMA                                          |
'+-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-+

Public strRutaCaratula As String
Public CopyMp3Totales As Integer
Public CopyTotalAlbums As Integer
Public bEQFreeMove As Boolean
Public bolCaratulaShow As Boolean, bolDirectoriosShow As Boolean
Attribute bolDirectoriosShow.VB_VarUserMemId = 1073741828
Public bolAcercaShow As Boolean, boolOptionsShow As Boolean
Attribute bolAcercaShow.VB_VarUserMemId = 1073741830
Attribute boolOptionsShow.VB_VarUserMemId = 1073741830
Public bolLyricsShow As Boolean, boolTagsShow As Boolean
Attribute bolLyricsShow.VB_VarUserMemId = 1073741832
Attribute boolTagsShow.VB_VarUserMemId = 1073741832
Public boolVisShow As Boolean, boolSearchShow As Boolean, bolSearchShow As Boolean
Attribute boolVisShow.VB_VarUserMemId = 1073741834
Attribute boolSearchShow.VB_VarUserMemId = 1073741834
Attribute bolSearchShow.VB_VarUserMemId = 1073741834
Public boolMediaLibraryShow As Boolean, bolMainShow As Boolean
Attribute boolMediaLibraryShow.VB_VarUserMemId = 1073741837
Attribute bolMainShow.VB_VarUserMemId = 1073741837
Public boolDspShow As Boolean
Attribute boolDspShow.VB_VarUserMemId = 1073741839
Public CurrentTrack_Index As Integer
Attribute CurrentTrack_Index.VB_VarUserMemId = 1073741840
Public bMultiSelect As Boolean
Attribute bMultiSelect.VB_VarUserMemId = 1073741841
Public NormalSelection As Integer
Attribute NormalSelection.VB_VarUserMemId = 1073741842
Public PlayerisClosing As Boolean
Attribute PlayerisClosing.VB_VarUserMemId = 1073741843

Public Enum peCrossfade
    CrossfadeNormal = 0
    FadeIn = 1
    FadeOut = 2
End Enum

Public boolVisLoaded As Boolean, boolOptionsLoaded As Boolean
Attribute boolVisLoaded.VB_VarUserMemId = 1073741844
Attribute boolOptionsLoaded.VB_VarUserMemId = 1073741844
Public boolTagsLoaded As Boolean
Attribute boolTagsLoaded.VB_VarUserMemId = 1073741846
Public boolDspLoaded As Boolean
Attribute boolDspLoaded.VB_VarUserMemId = 1073741847
Public boolOpluginLoaded As Boolean
Attribute boolOpluginLoaded.VB_VarUserMemId = 1073741848

Public OriginalWallpaperStyle As Integer
Attribute OriginalWallpaperStyle.VB_VarUserMemId = 1073741849
Public OriginalTileWallpaper As Integer
Attribute OriginalTileWallpaper.VB_VarUserMemId = 1073741850
Public OriginalRutaWallpaper As String
Attribute OriginalRutaWallpaper.VB_VarUserMemId = 1073741851

Public bolCaratulaDefault As Boolean
Attribute bolCaratulaDefault.VB_VarUserMemId = 1073741852
Public bLoadRegionFile As Boolean
Attribute bLoadRegionFile.VB_VarUserMemId = 1073741853
Public bolSplashScreen As Boolean
Attribute bolSplashScreen.VB_VarUserMemId = 1073741854
Public intActiveAlbum As Integer
Attribute intActiveAlbum.VB_VarUserMemId = 1073741855
Public TotalAlbumS As Integer
Attribute TotalAlbumS.VB_VarUserMemId = 1073741856
Public MP3totales As Integer
Attribute MP3totales.VB_VarUserMemId = 1073741857
Public Random_Order_track() As Integer    '// RANDOMISED ARRAY FOR PLAYLIST
Attribute Random_Order_track.VB_VarUserMemId = 1073741858


Public sTextScroll As String
Attribute sTextScroll.VB_VarUserMemId = 1073741859
Public sFileMainPlaying As String
Attribute sFileMainPlaying.VB_VarUserMemId = 1073741860
Public PlayerState As String
Attribute PlayerState.VB_VarUserMemId = 1073741861
Public bSearching As Boolean
Attribute bSearching.VB_VarUserMemId = 1073741862
Public bMinimize As Boolean
Attribute bMinimize.VB_VarUserMemId = 1073741863
Public bLoading As Boolean
Attribute bLoading.VB_VarUserMemId = 1073741864
Public strPathern As String
Attribute strPathern.VB_VarUserMemId = 1073741865
Public sFileType As String
Attribute sFileType.VB_VarUserMemId = 1073741866
Public sFormatPlayList As String
Attribute sFormatPlayList.VB_VarUserMemId = 1073741867
Public sFormatScroll As String
Attribute sFormatScroll.VB_VarUserMemId = 1073741868
Public iScrollType As Integer
Attribute iScrollType.VB_VarUserMemId = 1073741869
Public iScrollVel As Integer
Attribute iScrollVel.VB_VarUserMemId = 1073741870
Public iIdAlbumRC As Integer
Attribute iIdAlbumRC.VB_VarUserMemId = 1073741871
Public bAddFiles As Boolean
Attribute bAddFiles.VB_VarUserMemId = 1073741872
Public IndexVisualization As Integer
Attribute IndexVisualization.VB_VarUserMemId = 1073741873
Public tCurrentID3 As cMP3
Attribute tCurrentID3.VB_VarUserMemId = 1073741874

Public iCrossfadeTrack As Integer
Attribute iCrossfadeTrack.VB_VarUserMemId = 1073741875
Public iCrossfadeStop As Integer
Attribute iCrossfadeStop.VB_VarUserMemId = 1073741876
Public bPlayStarting As Boolean
Attribute bPlayStarting.VB_VarUserMemId = 1073741877
Public bPlayDFXStarting As Boolean
Attribute bPlayDFXStarting.VB_VarUserMemId = 1073741878
Public mainExecuted As Boolean
Attribute mainExecuted.VB_VarUserMemId = 1073741879
Public iOptPlstCase As Integer
Attribute iOptPlstCase.VB_VarUserMemId = 1073741880

'=======================================================
' VISUALIZACION
Public Type ptVisSpect
    Exist As Boolean
    BackColor As Long
    Mirrored As Boolean
    DrawSource As Integer
    ScaleUp As Integer
    ImageFile As String
    DrawBars As Boolean
    Gradient As String
    GrandientIndex As Integer
    Bars As Integer
    Spacio As Integer
    BackColorBar As Long
    DrawPeaks As Boolean
    BackColorPeak As Long
    arryPeaks() As Single
    arryWaitPeak() As String
    PeakHeight As Integer
    PeakGravity As Integer
End Type

Public Type FileTrack
    trackName As String
    trackPath As String
    Duration As String
End Type

Public Type ptVisScope
    LinesScope As Integer
    BackColorScope As Long
    Align As Integer
End Type

Public tConfigVis As ptVisSpect
Attribute tConfigVis.VB_VarUserMemId = 1073741881
Public tConfigScope() As ptVisScope
Attribute tConfigScope.VB_VarUserMemId = 1073741882

'=======================================================
Public Type Entry
    NoAlteraR As Boolean
    Mosaico As Boolean
    Centrar As Boolean
    Proporcional As Boolean
    Expander As Boolean
    Directorio As Boolean
    Language As String
    Ingles As Boolean
    Alpha As Integer
    SiempreTop As Boolean
    Splash As Boolean
    Instancias As Boolean
    TaskBar As Boolean
    SysTray As Boolean
    SysMenu As Boolean
End Type

Public Type ptSpec
    bDrawBars As Boolean
    iBars As Integer
    iSpacio As Integer
    lBackColorBar As Long
    lLineColorBar As Long
    bDrawPeaks As Boolean
    lBackColorPeak As Long
    iPeakHeight As Integer
    iPeakGravity As Integer
    iLinesScope As Integer
    lBackColorScope As Long
End Type

Public Type ptSlider
    Width As Integer
    Height As Integer
End Type

Public Type ptApp
    AppPath As String
    AppConfig As String
    Skin As String
End Type


Public tAppConfig As ptApp
Attribute tAppConfig.VB_VarUserMemId = 1073741883
Public tSpectrum As ptSpec
Attribute tSpectrum.VB_VarUserMemId = 1073741884

Public Type TrayIcon
    Previous As Boolean
    Play As Boolean
    Pause As Boolean
    Stop As Boolean
    Next As Boolean
End Type

Public bMiniMask As Boolean
Attribute bMiniMask.VB_VarUserMemId = 1073741885
Public OpcionesMusic As Entry
Attribute OpcionesMusic.VB_VarUserMemId = 1073741886
Public PlayerTrayIcon As TrayIcon
Attribute PlayerTrayIcon.VB_VarUserMemId = 1073741887
Public peCrossFadeType As peCrossfade
Attribute peCrossFadeType.VB_VarUserMemId = 1073741888
Public bCrossFadeEnabled As Boolean
Attribute bCrossFadeEnabled.VB_VarUserMemId = 1073741889
Public PlayerLoop As Boolean
Attribute PlayerLoop.VB_VarUserMemId = 1073741890
'-----Following for Enqueue Process------
Private nBuffer() As Byte
Attribute nBuffer.VB_VarUserMemId = 1073741891
Private Const WM_COPYDATA = &H4A
Private Type COPYDATASTRUCT
    dwData As Long
    cbData As Long
    lpData As Long
End Type
Private nCopyData As COPYDATASTRUCT
Attribute nCopyData.VB_VarUserMemId = 1073741892
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)


'+-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-+
'|  INICIO DE LA APLICATION                                                              |
'+-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-+

Sub Main()

    On Error Resume Next
    Dim strPath As String, args As String
    Dim strRes
    Call EnqueueProcess


    bLoading = True
    boolOptionsShow = False
    boolDspShow = False

    Load_Settings_INI True

    args = Trim(Command$)

    If args <> "" Then
        ProcessCommandParameter (args)
        frmPLST.Update_Plst_Scrollbar
        If frmPLST.clist.ItemCount > 0 Then
            CurrentTrack_Index = 0
            sFileMainPlaying = frmPLST.clist.exItem(CurrentTrack_Index)
            frmMain.PlayerIsPlaying = "true"
        End If
        bLoading = False

    Else  'No command line argument
        LoadPlaylist
    End If



    If frmPLST.clist.ItemCount >= 1 Then
        CurrentTrack_Index = 0
        sFileMainPlaying = frmPLST.clist.exItem(CurrentTrack_Index)
        frmMain.PlayerIsPlaying = "true"
        strRes = Read_INI("Configuration", "PlayStarting", 1, , True)
        If CBool(strRes) = True Then bPlayStarting = True

        If bPlayStarting = True Then
            Load (frmLibrary)
            frmLibrary.Visible = False
            DoEvents
            frmMain.Play
            frmMain.Load_File_Tags
        Else
            frmMain.Button(3).Selected = True
            frmMain.ScrollText(1).CaptionText = "--NEO MP3 PLAYER--By Mahesh Kurmi"
            frmMain.ScrollText(2).CaptionText = " "
            frmMain.ScrollText(3).CaptionText = " "
        End If
    End If


    If OpcionesMusic.Splash = True Then Unload frmSplash
    frmMain.Show
    DoEvents
    DoEvents
    DoEvents
    frmLibrary.Visible = boolMediaLibraryShow
    bLoading = False
    bLoadingSkin = False
    If frmMain.Button(26).Selected Then
        frmPopUp.mnuShowPlst.Checked = True
        frmPLST.Show
        Call CheckMenu(11, True)
    End If


    PlayerisClosing = False

    Call Hook(frmMain.hwnd)
    Call EnableFileDrops(frmMain)

End Sub


'PROCESSES WHETHER OR NOT IS FIRST INSTANCE - IF IT ISNT, IT SENDS THE FILEPATH OF PATH TO FIRST INSTANCE
Public Sub EnqueueProcess()
    Dim sCommand As String
    sCommand = Command$
    If App.PrevInstance And sCommand <> "%1" And sCommand <> "" Then
        Dim lhWnd As Long

        lhWnd = CLng(Val(GetSetting(App.Title, "ActiveWindow", "Handle")))
        ReDim nBuffer(1 To Len(Command$) + 1)
        Call CopyMemory(nBuffer(1), ByVal Command$, Len(Command$))

        With nCopyData
            .dwData = 3
            .cbData = Len(sCommand) + 1
            .lpData = VarPtr(nBuffer(1))
        End With

        Call SendMessage(lhWnd, WM_COPYDATA, lhWnd, nCopyData)
        End
    Else
        SaveSetting App.Title, "ActiveWindow", "Handle", str(frmMain.hwnd)
    End If
End Sub


Public Sub ProcessCommandParameter(args As String)

    On Error GoTo errorHandler:
    If FolderExists(args) Then    'If it is folder
        Call frmPLST.Enqueue_Track(args, "RUND")    'If it is file
    ElseIf FileExists(args) Then
        If UCase(Right(args, 3)) = "NPL" Or UCase(Right(args, 3)) = "M3U" Or UCase(Right(args, 3)) = "PLS" Then
            Call frmPLST.Enqueue_Track(args, "RUNP")
        Else
            Call frmPLST.Enqueue_Track(args, "")
        End If
    ElseIf Left(args, 1) = """" Then
        Dim Arr() As String
        Dim i As Integer
        Arr = Split(args, """", , vbTextCompare)
        For i = LBound(Arr) + 1 To UBound(Arr) - 1
            If FolderExists(Arr(i)) Then
                Call frmPLST.Enqueue_Track(Arr(i), "ADDD")
            ElseIf FileExists(Arr(i)) Then
                Call frmPLST.Enqueue_Track(Arr(i), "ADDF")
            End If
        Next
    Else
        Call frmPLST.Enqueue_Track(Right(args, Len(args) - 4), Left(args, 4))
    End If
    Exit Sub
errorHandler:
    MsgBox err.Description & args & "processcommand"
End Sub
