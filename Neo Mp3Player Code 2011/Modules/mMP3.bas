Attribute VB_Name = "mMP3"


Option Explicit



Public Declare Function CreateFile Lib "kernel32" Alias "CreateFileA" (ByVal lpFileName As String, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, lpSecurityAttributes As SECURITY_ATTRIBUTES, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As Long) As Long
Public Type SECURITY_ATTRIBUTES
    nLength As Long
    lpSecurityDescriptor As Long
    bInheritHandle As Long
End Type
Public Const GENERIC_WRITE = &H40000000
Public Const OPEN_EXISTING = 3
Public Const FILE_FLAG_BACKUP_SEMANTICS = &H2000000

Public Declare Function SetFilePointer Lib "kernel32" (ByVal hFile As Long, ByVal lDistanceToMove As Long, lpDistanceToMoveHigh As Long, ByVal dwMoveMethod As Long) As Long
Public Declare Function SetEndOfFile Lib "kernel32" (ByVal hFile As Long) As Long
Public Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long

Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Const EM_LINEFROMCHAR = &HC9
Public Const EM_LINEINDEX = &HBB
Public Const EM_GETLINECOUNT = &HBA
Public Const EM_LINELENGTH = &HC1
Public Const ZERO = 0

Public Type ID3v1Tag
    ID As String * 3
    Title As String * 30
    Artist As String * 30
    Album As String * 30
    Year As String * 4
    Comment As String * 30
    Genre As Byte
End Type

Public Type ptMPEG
    FileSize As String
    bitrate As Long
    Frequency As Long
    Version As Integer
    Layer As Integer
    Header As Integer
    Mode As String
    Emphasis As String
    Original As String
    Copyrighted As String
    Private As String
    CRCs As String
    Duration As String
    Framesize As Integer
    Frames As Long
    Padding As String
    VBR As Boolean
End Type

Public Type ptID3
    Title As String
    Artist As String
    Album As String
    Year As String
    Comment As String
    Genre As Variant
    GenreName As String
    Lyrics As String
    Lyrics3Tag As Boolean
    ID3v1Tag As Boolean
End Type

Public tCurrentMPEGInfo As ptMPEG
Private sLyrics As String
Private HasLyrics3Tag As Boolean
Private HasID3v1Tag As Boolean
Private LSZ As String
Private s As String
Private posLyrics As Long
Private gsGenres As String

Public Function ReadFile_Tags(FileName As String) As ptID3

    On Error GoTo errorHandler
    Dim ITemp As Integer
    Dim NewGenre As String

    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' use the filename to get ID3 info
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Dim lngFilesize As Long
    Dim fn As Integer
    Dim Tag1 As ID3v1Tag
    Dim LyricEndID As String * 6

    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' Open the file
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

    fn = FreeFile

    Open FileName For Binary As #fn         'Open the file so we can read it
    lngFilesize = LOF(fn)                   'Size of the file, in bytes

    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' Check for an ID3v1 tag
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

    'ID3v1 tag

    Get #fn, lngFilesize - 127, Tag1.ID

    If Tag1.ID = "TAG" Then    'If "TAG" is present, then we have a valid ID3v1 tag and will extract all available ID3v1 info from the file
        Get #fn, , Tag1.Title   'Always limited to 30 characters
        Get #fn, , Tag1.Artist  'Always limited to 30 characters
        Get #fn, , Tag1.Album   'Always limited to 30 characters
        Get #fn, , Tag1.Year    'Always limited to 4 characters
        Get #fn, , Tag1.Comment    'Always limited to 30 characters
        Get #fn, , Tag1.Genre   'Always limited to 1 byte (?)

        HasID3v1Tag = True

        ReadFile_Tags.Title = Replace(Trim(Tag1.Title), Chr(0), "")
        ReadFile_Tags.Artist = Replace(Trim(Tag1.Artist), Chr(0), "")
        ReadFile_Tags.Album = Replace(Trim(Tag1.Album), Chr(0), "")
        ReadFile_Tags.Year = Replace(Trim(Tag1.Year), Chr(0), "")
        ReadFile_Tags.Comment = Trim(Replace(Trim(Tag1.Comment), Chr(0), ""))
        ReadFile_Tags.Genre = Tag1.Genre

        '// si no hay nada poner default 'Other'
        Populate_Genres
        If Trim(Tag1.Genre) = "" Then ITemp = 12 Else ITemp = CInt(Tag1.Genre)
        If ITemp <= 0 Then ITemp = 1 Else ITemp = ITemp * 22
        NewGenre = Trim$(Mid$(gsGenres, ITemp, 22))
        If Trim(NewGenre) = "" Then NewGenre = "Other"
        ReadFile_Tags.GenreName = NewGenre
        gsGenres = ""
    Else

        HasID3v1Tag = False

    End If

    ReadFile_Tags.ID3v1Tag = HasID3v1Tag

    'lyrics3 tag
    If HasID3v1Tag = True Then
        Get #fn, lngFilesize - 136, LyricEndID    'look for a lyrics 3 tag
    Else
        Get #fn, lngFilesize - 8, LyricEndID   'look for a lyrics 3 tag
    End If

    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' Close the file
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Close

    'lyrics3 tag?
    If LyricEndID = "LYRICS" Then    'got one, go get it
        HasLyrics3Tag = GetLyrics3Tag(FileName)
        ReadFile_Tags.Lyrics = sLyrics
        ReadFile_Tags.Lyrics3Tag = HasLyrics3Tag
    Else
        HasLyrics3Tag = False
        ReadFile_Tags.Lyrics3Tag = False
    End If

    Exit Function

errorHandler:
    err.Clear
    Close
End Function


'// sakar los valores del mpeginfo
Public Function MPEGInfo(FileName As String) As ptMPEG
    On Error Resume Next

    Dim ByteArray(4) As Byte
    Dim XingH As String * 4
    Dim FIO As Long
    Dim i As Long
    Dim z As Long
    Dim TempVar As Byte
    Dim HeadStart As Long
    Dim Frames As Long
    Dim BitString As String
    Dim Temp As Variant
    Dim Brate As Variant
    Dim freq As Variant
    Dim dFilesize As Double
    Dim HH As Long, MM As Long, ss As Long, sTempTime As String, lSeconds As Long

    Dim YesNo, NoYes

    YesNo = Array("Yes", "No")
    NoYes = Array("No", "Yes")


    'tables
    Dim VersionLayer(3) As String
    VersionLayer(0) = 0
    VersionLayer(1) = 3
    VersionLayer(2) = 2
    VersionLayer(3) = 1

    Dim SMode(3) As String
    SMode(0) = "stereo"
    SMode(1) = "joint stereo"
    SMode(2) = "dual channel"
    SMode(3) = "single channel"

    Dim Emphasis(3) As String
    Emphasis(0) = "None"
    Emphasis(1) = "50/15"
    Emphasis(2) = "Reserved"
    Emphasis(3) = "CCITT J 17"

    FIO = FreeFile

    'read the header
    Open FileName For Binary Access Read As FIO
    If LOF(FIO) < 256 Then
        Close FIO
        Exit Function
    End If

    '=========================================================================
    ' COMENZAR A CHEKAR LA POSICION DEL ENCABEZADO( HEADER)
    ' SI COMIENZA DIFERENTE 1 ENTONCES TIENE ID3V2 TAGS
    '=========================================================================

    For i = 1 To LOF(FIO)           ' chekar en todo el archivo para el header
        Get #FIO, i, TempVar
        If TempVar = 255 Then       ' header siempre comienza con 255 seguido por 250 o 251
            Get #FIO, i + 1, TempVar
            If TempVar > 249 And TempVar < 252 Then
                HeadStart = i       ' guardar posicion del header
                MPEGInfo.Header = HeadStart
                Exit For
            End If
        End If
    Next i

    '=========================================================================
    ' no hay header
    If HeadStart = 0 Then
        Exit Function
    End If
    '=========================================================================


    '=========================================================================
    ' comenzar a buscar si hay XingHeader
    Get #FIO, HeadStart + 36, XingH
    If XingH = "Xing" Then
        MPEGInfo.VBR = True
        For z = 1 To 4    '
            Get #1, HeadStart + 43 + z, ByteArray(z)  ' asignar framelength a un arreglo
        Next z
        Frames = BinToDec(ByteToBit(ByteArray))       ' calkular numero de frames
    Else
        MPEGInfo.VBR = False
    End If
    '=========================================================================

    '=========================================================================
    ' extraer los primeros 4 bytes(32bits) a un arreglo
    For z = 1 To 4    '
        Get #FIO, HeadStart + z - 1, ByteArray(z)
    Next z
    '=========================================================================
    Close FIO

    'header string
    BitString = ByteToBit(ByteArray)

    'asignar mpegversion de una tabla
    MPEGInfo.Version = VersionLayer(BinToDec(Mid(BitString, 12, 2)))

    'obtener la layer de la tabla
    MPEGInfo.Layer = VersionLayer(BinToDec(Mid(BitString, 14, 2)))


    'obtener el mode de la tabla
    MPEGInfo.Mode = SMode(BinToDec(Mid(BitString, 25, 2)))

    'obtener emphasis
    MPEGInfo.Emphasis = Emphasis(BinToDec(Mid(BitString, 31, 2)))

    'deacuerdo con la version construir la tabla adecuada
    Select Case MPEGInfo.Version
    Case 1       'para la version 1
        freq = Array(44100, 48000, 32000)
    Case 2 Or 25    'para la version 2 o 2.5
        freq = Array(22050, 24000, 16000)
    Case Else
        MPEGInfo.Frequency = 0
        Exit Function
    End Select

    'buscar la frequencia de la tabla freq
    If (freq(BinToDec(Mid(BitString, 21, 2))) = "") Then
        MPEGInfo.Frequency = "44100"
    Else
        MPEGInfo.Frequency = freq(BinToDec(Mid(BitString, 21, 2)))
    End If


    If MPEGInfo.VBR = True Then
        'redifinir para calcular el correcto bitrate
        Temp = Array(, 12, 144, 144)
        MPEGInfo.bitrate = (FileLen(FileName) * MPEGInfo.Frequency) / (Int(Frames)) / 1000 / Temp(MPEGInfo.Layer)
    Else
        'Correcto bitrate
        Select Case Val(MPEGInfo.Version & MPEGInfo.Layer)
        Case 11  'Version 1, Layer 1
            Brate = Array(0, 32, 64, 96, 128, 160, 192, 224, 256, 288, 320, 352, 384, 416, 448)
        Case 12  'Version 1 Layer 2
            Brate = Array(0, 32, 48, 56, 64, 80, 96, 112, 128, 160, 192, 224, 256, 320, 384)
        Case 13  'Version 1 Layer 3
            Brate = Array(0, 32, 40, 48, 56, 64, 80, 96, 112, 128, 160, 192, 224, 256, 320)
        Case 21 Or 251    'V2 L1 and V2.5 L1
            Brate = Array(0, 32, 48, 56, 64, 80, 96, 112, 128, 144, 160, 176, 192, 224, 256)
        Case 22 Or 252 Or 23 Or 253  'V2 L2 and 'V2.5 L2 etc...
            Brate = Array(0, 8, 16, 24, 32, 40, 48, 56, 64, 80, 96, 112, 128, 144, 160)
        Case Else   'si el bitrate es variable
            MPEGInfo.bitrate = 1
            Exit Function
        End Select

        MPEGInfo.bitrate = Brate(BinToDec(Mid(BitString, 17, 4)))
    End If

    'si hay un lugar decimal quitarlo
    If InStr(1, MPEGInfo.bitrate, ".") Then
        MPEGInfo.bitrate = Left(MPEGInfo.bitrate, InStr(1, MPEGInfo.bitrate, ".") - 1)
    End If

    'obtener original
    MPEGInfo.Original = NoYes(Mid(BitString, 30, 1))

    'obtener copyrighted
    MPEGInfo.Copyrighted = NoYes(Mid(BitString, 29, 1))

    'obtener private
    MPEGInfo.Private = NoYes(Mid(BitString, 24, 1))

    'i am not so sure if the padding thing here is right
    MPEGInfo.Padding = NoYes(Mid(BitString, 23, 1))

    'obtener CRC
    MPEGInfo.CRCs = YesNo(Mid(BitString, 16, 1))

    'obtener Frame Size
    MPEGInfo.Framesize = (MPEGInfo.bitrate * 144000) / (MPEGInfo.Frequency)

    If MPEGInfo.Padding = "Yes" Then MPEGInfo.Framesize = MPEGInfo.Framesize + 1

    MPEGInfo.Frames = Int(FileLen(FileName) / MPEGInfo.Framesize)


    'calcular duracion
    lSeconds = Int((FileLen(FileName) * 8) / MPEGInfo.bitrate / 1000)

    HH = lSeconds \ 3600      ' Horas
    MM = lSeconds \ 60 Mod 60    ' Minutos
    ss = lSeconds Mod 60      ' Segundos

    If HH > 0 Then sTempTime = HH & ":"
    MPEGInfo.Duration = Trim(sTempTime & Format$(MM, "00") & ":" & Format$(ss, "00"))

    'Tamaño del archivo MBs
    dFilesize = FileLen(FileName) / 1024
    MPEGInfo.FileSize = CLng(dFilesize / 1024 * 100) / 100 & " MB"


End Function

Private Function BinToDec(BinValue As String) As Long
    On Error Resume Next

    Dim i As Long
    BinToDec = 0
    For i = 1 To Len(BinValue)
        If Mid(BinValue, i, 1) = 1 Then
            BinToDec = BinToDec + 2 ^ (Len(BinValue) - i)
        End If
    Next i

    Exit Function

End Function

Private Function ByteToBit(ByteArray) As String
    On Error GoTo ErrHand

    Dim z As Integer
    Dim i As Integer
    'convert 4*1 byte array to 4*8 bits'''''
    ByteToBit = ""
    For z = 1 To 4
        For i = 7 To 0 Step -1
            If Int(ByteArray(z) / (2 ^ i)) = 1 Then
                ByteToBit = ByteToBit & "1"
                ByteArray(z) = ByteArray(z) - (2 ^ i)
            ElseIf ByteToBit <> "" Then
                ByteToBit = ByteToBit & "0"
            End If
        Next
    Next z

    Exit Function

ErrHand:
    err.Raise err.Description

End Function


Public Function WriteTag(strFileName As String, tID3 As ptID3) As Boolean

    Dim cID3v1Tag As ID3v1Tag
    Dim WholeTag As String
    Dim tagsize As String * 6
    Dim Position As Long
    Dim MoveMP3Tag As Boolean
    Dim fn As Integer
    Dim UseOldInfo As Boolean
    Dim strAuthor As String
    Dim NewLyr As Boolean
    WholeTag = ""

    On Error GoTo hell

    '// you can add more tags(fields) at file for exemple i add LYR -> Lyrics
    '// (but you can add more ex: AUT -> author)
    If tID3.Lyrics <> "" Then NewLyr = True

    '//Author field
    '   strAuthor = "Raul Martinez"
    '   If strAuthor <> "" Then NewLyr = True

    'build the tag
    If NewLyr = True Then
        WholeTag = "LYRICSBEGIN"
        If tID3.Lyrics <> "" Then
            WholeTag = WholeTag & "LYR" & Format(Len(tID3.Lyrics), "00000") & tID3.Lyrics
        End If

        '      If strAuthor <> "" Then
        '         WholeTag = WholeTag & "AUT" & Format(Len(strAuthor), "00000") & strAuthor
        '      End If

        tagsize = Format(Len(WholeTag), "000000")

        'append the end identifier
        WholeTag = WholeTag & tagsize & "LYRICS200"

    End If

    'prepare for writing
    fn = FreeFile
    Open strFileName For Binary As #fn

    If HasID3v1Tag = True Then
        'set to just before the current id3 tag.
        Position = LOF(fn) - 127
    Else
        Position = LOF(fn) + 1
    End If

    'if there is a Lyrics3 tag, then go back to the beginning of the old one
    If HasLyrics3Tag Then Position = posLyrics


    ' write the lyrics3 tag if there is one...
    If WholeTag <> "" Then
        Put #fn, Position, WholeTag
        Position = Seek(fn)
    End If

    'write the id3tag

    cID3v1Tag.ID = "TAG"
    cID3v1Tag.Title = tID3.Title
    cID3v1Tag.Artist = tID3.Artist
    cID3v1Tag.Album = tID3.Album
    cID3v1Tag.Genre = tID3.Genre
    cID3v1Tag.Year = tID3.Year
    cID3v1Tag.Comment = tID3.Comment

    Put #fn, Position, cID3v1Tag

    'set the last byte of the file
    Position = Seek(fn) - 1
    Close
    'make sure this is the end of the file which is needed if this tag is smaller than the old tag.
    SetFileLength strFileName, Position
    Exit Function
hell:
End Function

Public Function GetLyrics3Tag(sFilename As String) As Boolean
    On Error GoTo hell
    Dim Position As Long
    Dim FieldData As String   '// save Text of field
    Dim FieldID As String * 3    '// save Field ex: LYR, ALB, AUT... etc...
    Dim LengthField As String
    Dim fn As Integer
    Dim Size As Long          '// size of File
    Dim tagtype As String * 9    '// save end tag lyrics
    Dim Byte11Buffer As String * 11    '// save LYRICSBEGIN
    Dim Byte5Buffer As String * 5   '// Length of Field
    Dim lEndLyr As Long

    '//reset size of tag
    LengthField = "000000"

    '//open the file
    fn = FreeFile
    Open sFilename For Binary As #fn

    'get filesize
    Size = LOF(fn)

    'get the tag END
    If HasID3v1Tag = True Then
        'get the tag END [ (127 -> ID3v1Tag) + (9  -> LYRICS200) ] = 136
        lEndLyr = Size - 136
        Get #fn, lEndLyr, tagtype
    Else
        lEndLyr = Size - 8
        Get #fn, lEndLyr, tagtype
    End If

    'if tag is valid then
    If tagtype = "LYRICS200" Then

        'get the size of the tag  ( 136 - 6 ) = 142
        Get #fn, lEndLyr - 6, LengthField

        'set the position to the first byte of Lyrics
        Position = lEndLyr - 6 - Val(LengthField)

        'get the beginning of tag
        Get #fn, Position, Byte11Buffer
        'save beginning lyrics tag
        posLyrics = Position

        If Byte11Buffer <> "LYRICSBEGIN" Then
            'invalid Lyrics3 version 2 tag! we don't support version 1...
            Close
            Exit Function
        End If

        'first field ( Position + LYRICSBEGIN )
        Position = Position + 11

        'keep getting fields until we get to the end of the tag
        Do Until Position >= lEndLyr - 9

            'the field id -> LYR-ALB-AUT. etc...
            Get #fn, Position, FieldID

            'the size of the field
            Get #fn, Position + 3, Byte5Buffer
            LengthField = Val(Byte5Buffer)

            'make room for the data
            FieldData = Space(Val(LengthField))
            'get the data
            Get #fn, Position + 8, FieldData
            'and fill the approprate field
            Select Case FieldID
            Case "LYR"    '// Lyrics
                sLyrics = Trim$(FieldData)
                '           Case "AUT" '// Author
                '              MsgBox "Copyright : " & Trim$(FieldData)
            End Select

            'now set the postion to the beginning of the next field
            Position = Position + 8 + Val(LengthField)
        Loop
        'set the flag
        GetLyrics3Tag = True
    End If

    Close

    Exit Function
hell:
    'close all open files
    Close
End Function

Private Sub SetFileLength(strFileName As String, ByVal NewLength As Long)

'Will cut the length of a file to the length specified.

    Dim hFile As Long
    Dim l As Long
    Dim lpSecurity As SECURITY_ATTRIBUTES
    'if file is smaller than or equal to requsted length, exit.
    If FileLen(strFileName) <= NewLength Then Exit Sub
    'open the file
    hFile = CreateFile(strFileName, GENERIC_WRITE, ZERO, lpSecurity, OPEN_EXISTING, FILE_FLAG_BACKUP_SEMANTICS, 0)
    'if file not open exit
    If hFile = -1 Then Exit Sub
    'seek to position
    l = SetFilePointer(hFile, NewLength, ZERO, ZERO)
    'and mark here as end of file
    SetEndOfFile hFile
    'close the file
    l = CloseHandle(hFile)
End Sub

Private Sub Populate_Genres()
    gsGenres = "Blues                 "
    gsGenres = gsGenres & "Classic Rock          "
    gsGenres = gsGenres & "Country               "
    gsGenres = gsGenres & "Dance                 "
    gsGenres = gsGenres & "Disco                 "
    gsGenres = gsGenres & "Gazal                 "
    gsGenres = gsGenres & "Funk                  "
    gsGenres = gsGenres & "Grunge                "
    gsGenres = gsGenres & "Hip-Hop               "
    gsGenres = gsGenres & "Jazz                  "
    gsGenres = gsGenres & "Metal                 "
    gsGenres = gsGenres & "New Age               "
    gsGenres = gsGenres & "Oldies                "
    gsGenres = gsGenres & "Other                 "
    gsGenres = gsGenres & "Pop                   "
    gsGenres = gsGenres & "R&B                   "
    gsGenres = gsGenres & "Rap                   "
    gsGenres = gsGenres & "Reggae                "
    gsGenres = gsGenres & "Rock                  "
    gsGenres = gsGenres & "Techno                "
    gsGenres = gsGenres & "Industrial            "
    gsGenres = gsGenres & "Alternative           "
    gsGenres = gsGenres & "Ska                   "
    gsGenres = gsGenres & "Death Metal           "
    gsGenres = gsGenres & "Pranks                "
    gsGenres = gsGenres & "Soundtrack            "
    gsGenres = gsGenres & "Euro-Techno           "
    gsGenres = gsGenres & "Ambient               "
    gsGenres = gsGenres & "Trip-Hop              "
    gsGenres = gsGenres & "Vocal                 "
    gsGenres = gsGenres & "Jazz+Funk             "
    gsGenres = gsGenres & "Fusion                "
    gsGenres = gsGenres & "Trance                "
    gsGenres = gsGenres & "Classical             "
    gsGenres = gsGenres & "Instrumental          "
    gsGenres = gsGenres & "Acid                  "
    gsGenres = gsGenres & "House                 "
    gsGenres = gsGenres & "Game                  "
    gsGenres = gsGenres & "Sound Clip            "
    gsGenres = gsGenres & "Gospel                "
    gsGenres = gsGenres & "Noise                 "
    gsGenres = gsGenres & "Alternative Rock      "
    gsGenres = gsGenres & "Bass                  "
    gsGenres = gsGenres & "Soul                  "
    gsGenres = gsGenres & "Punk                  "
    gsGenres = gsGenres & "Space                 "
    gsGenres = gsGenres & "Meditative            "
    gsGenres = gsGenres & "Instrumental Pop      "
    gsGenres = gsGenres & "Instrumental Rock     "
    gsGenres = gsGenres & "Ethnic                "
    gsGenres = gsGenres & "Gothic                "
    gsGenres = gsGenres & "Darkwave              "
    gsGenres = gsGenres & "Techno-Industrial     "
    gsGenres = gsGenres & "Electronic            "
    gsGenres = gsGenres & "Pop-Folk              "
    gsGenres = gsGenres & "Eurodance             "
    gsGenres = gsGenres & "Dream                 "
    gsGenres = gsGenres & "Southern Rock         "
    gsGenres = gsGenres & "Comedy                "
    gsGenres = gsGenres & "Cult                  "
    gsGenres = gsGenres & "Gangsta               "
    gsGenres = gsGenres & "Top 40                "
    gsGenres = gsGenres & "Christian Rap         "
    gsGenres = gsGenres & "Pop/Funk              "
    gsGenres = gsGenres & "Jungle                "
    gsGenres = gsGenres & "Native US             "
    gsGenres = gsGenres & "Cabaret               "
    gsGenres = gsGenres & "New Wave              "
    gsGenres = gsGenres & "Psychadelic           "
    gsGenres = gsGenres & "Rave                  "
    gsGenres = gsGenres & "Showtunes             "
    gsGenres = gsGenres & "Trailer               "
    gsGenres = gsGenres & "Lo-Fi                 "
    gsGenres = gsGenres & "Tribal                "
    gsGenres = gsGenres & "Acid Punk             "
    gsGenres = gsGenres & "Acid Jazz             "
    gsGenres = gsGenres & "Polka                 "
    gsGenres = gsGenres & "Retro                 "
    gsGenres = gsGenres & "Musical               "
    gsGenres = gsGenres & "Rock & Roll           "
    gsGenres = gsGenres & "Hard Rock             "
    gsGenres = gsGenres & "Folk                  "
    gsGenres = gsGenres & "Folk-Rock             "
    gsGenres = gsGenres & "National Folk         "
    gsGenres = gsGenres & "Swing                 "
    gsGenres = gsGenres & "Fast Fusion           "
    gsGenres = gsGenres & "Bebob                 "
    gsGenres = gsGenres & "Latin                 "
    gsGenres = gsGenres & "Revival               "
    gsGenres = gsGenres & "Celtic                "
    gsGenres = gsGenres & "Bluegrass             "
    gsGenres = gsGenres & "Avantgarde            "
    gsGenres = gsGenres & "Gothic Rock           "
    gsGenres = gsGenres & "Progressive Rock      "
    gsGenres = gsGenres & "Psychedelic Rock      "
    gsGenres = gsGenres & "Symphonic Rock        "
    gsGenres = gsGenres & "Slow Rock             "
    gsGenres = gsGenres & "Big Band              "
    gsGenres = gsGenres & "Chorus                "
    gsGenres = gsGenres & "Easy Listening        "
    gsGenres = gsGenres & "Acoustic              "
    gsGenres = gsGenres & "Humour                "
    gsGenres = gsGenres & "Speech                "
    gsGenres = gsGenres & "Chanson               "
    gsGenres = gsGenres & "Opera                 "
    gsGenres = gsGenres & "Chamber Music         "
    gsGenres = gsGenres & "Sonata                "
    gsGenres = gsGenres & "Symphony              "
    gsGenres = gsGenres & "Booty Bass            "
    gsGenres = gsGenres & "Primus                "
    gsGenres = gsGenres & "Porn Groove           "
    gsGenres = gsGenres & "Satire                "
    gsGenres = gsGenres & "Slow Jam              "
    gsGenres = gsGenres & "Club                  "
    gsGenres = gsGenres & "Tango                 "
    gsGenres = gsGenres & "Samba                 "
    gsGenres = gsGenres & "Folklore              "
    gsGenres = gsGenres & "Ballad                "
    gsGenres = gsGenres & "Power Ballad          "
    gsGenres = gsGenres & "Rhytmic Soul          "
    gsGenres = gsGenres & "Freestyle             "
    gsGenres = gsGenres & "Duet                  "
    gsGenres = gsGenres & "Punk Rock             "
    gsGenres = gsGenres & "Drum Solo             "
    gsGenres = gsGenres & "Acapella              "
    gsGenres = gsGenres & "Euro-House            "
    gsGenres = gsGenres & "Dance Hall            "
    gsGenres = gsGenres & "Goa                   "
    gsGenres = gsGenres & "Drum & Bass           "
    gsGenres = gsGenres & "Club-House            "
    gsGenres = gsGenres & "Hardcore              "
    gsGenres = gsGenres & "Terror                "
    gsGenres = gsGenres & "Indie                 "
    gsGenres = gsGenres & "BritPop               "
    gsGenres = gsGenres & "Negerpunk             "
    gsGenres = gsGenres & "Polsk Punk            "
    gsGenres = gsGenres & "Beat                  "
    gsGenres = gsGenres & "Christian Gangsta Rap "
    gsGenres = gsGenres & "Heavy Metal           "
    gsGenres = gsGenres & "Black Metal           "
    gsGenres = gsGenres & "Crossover             "
    gsGenres = gsGenres & "Contemporary Christian"
    gsGenres = gsGenres & "Christian Rock        "
    gsGenres = gsGenres & "Merengue              "
    gsGenres = gsGenres & "Salsa                 "
    gsGenres = gsGenres & "Trash Metal           "
    gsGenres = gsGenres & "Anime                 "
    gsGenres = gsGenres & "Jpop                  "
    gsGenres = gsGenres & "Synthpop              "
End Sub


