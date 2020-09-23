VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cMP3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'##   PIRATEADO Y MODIFICADO POR RAUL MARTINEZ ##

' *************************************************************************************************
' This MP3 tagging module reads, writes and deletes ID3 v 1.0, 1.1, 2.2, 2.3 and 2.4 tags,
' with support for extended tag headers, unsynchronized tags, tag footers, appended tags and
' MPEG file information.
'
' This module is a part of Magic MP3 Tagger. Visit www.magic-tagger.com if you're interested in
' this top MP3 tagger!
'
' Copyright by Mathias Kunter.
' You're free to use this module within your programs. These programs may also be commercial.
' The only condition for usage of this module is that you show the following line within the
' credits of your program:
'
' ID3 tagging module by Mathias Kunter (www.magic-tagger.com)
'
' *************************************************************************************************
Option Explicit
Private HasLyrics3Tag As Boolean
Private PosTagLyr As String
Private m_sFileName As String


Private Type tOtherInfo
    Lyrics As String
    Quality As String
    Situation As String
    Mood As String
    Author As String
    OtherInfo As String
    Images As String
    UnknownFields As String
    Album As String
    Artist As String
    Title As String
End Type

Private m_LastUpdate As String
Private m_fReadMPEGInfo As Boolean


'***** The ID3 tag for data input / output with this module *****
Private Type ID3Tag
    Artist As String
    Album As String
    Title As String
    SongYear As String
    Genre As String
    Comment As String
    TrackNr As String

    'The following items are only used with ID3 v2 tags.
    EncodedBy As String
    LinkTo As String
    Copyright As String
    Composer As String
    OrigArtist As String
End Type

Private Type MPEGInfo
    MPEGVersion As String
    HasCRC As Boolean
    bitrate As Long     'kbit / sec (where k = 1000, not 1024)
    Frequency As Long   'hz
    HasVBR As Boolean
    ChannelMode As String
    IsCopyrighted As Boolean
    IsOriginal As Boolean
    HasEmphasis As Boolean
    Length As Long      'sec
    ID3v1Version As Long    '-1 = no tag
    ID3v2Version As Long    '-1 = no tag
    FileSize As Long    'bytes
    LengthTime As String
    SizeMB As String
End Type

Private Enum ID3v2_Version
    VERSION_2_3
    VERSION_2_4
End Enum

Private myTag As ID3Tag
Private myOtherinfo As tOtherInfo
Private myMPEGInfo As MPEGInfo

Private Const ID3_GENRE_COUNT As Long = 148



'***** ID3v1 Types *****
Private Type v1Tag
    Identifier(2) As Byte
    Title(29) As Byte
    Artist(29) As Byte
    Album(29) As Byte
    SongYear(3) As Byte
    Comment(29) As Byte
    Genre As Byte
End Type



'***** ID3v2 Types *****

'Tag header for all versions (2.2 - 2.4)
'The v 2.4 footer also uses this structure, since it's a copy of the header.
Private Type v2TagHeader
    Identifier(2) As Byte
    Version(1) As Byte
    Flags As Byte
    Size(3) As Byte
End Type

'Extended tag header for v 2.3
Private Type v2_3ExtHeader
    Size(3) As Byte
    ExtendedFlags(1) As Byte
    PaddingSize(3) As Byte
End Type

'Extended tag header for v 2.4
Private Type v2_4ExtHeader
    Size(3) As Byte
    NumFlags As Byte
    ExtendedFlags As Byte
End Type

'Frame header for v 2.2
Private Type v2_2FrameHeader
    Ident(2) As Byte
    Size(2) As Byte
End Type

'Frame header for v 2.3 and v 2.4
Private Type v2_34FrameHeader
    Ident(3) As Byte
    Size(3) As Byte
    Flags(1) As Byte
End Type

'A v 2.x frame in memory.
Private Type v2MemFrame
    ID As String
    Size As Long
    Data() As Byte
End Type

Private Type v2TagHeaderFlags
    bUnsynch As Boolean
    bExtHeader As Boolean
    bFooter As Boolean
    bUpdateTag As Boolean
End Type

Private Type v2TagSizes
    ReadSize As Long
    ExtHeaderSize As Long
    FooterSize As Long
End Type

Private Enum v2_StrEncoding
    ENC_ISO = 0
    ENC_UNICODE_UTF16_BOM = 1
    ENC_UNICODE_UTF16 = 2
    ENC_UNICODE_UTF8 = 3
End Enum

'The frames stored in memory, used when updating a currently existing v2 tag.
Private mMemFrames() As v2MemFrame, mMemFraCnt As Long
Attribute mMemFraCnt.VB_VarUserMemId = 1073938440

Private Const cGenreString As String = _
        "Blues#Classic Rock#Country#Dance#Disco#Gazal#Funk#Grunge#Hip-Hop#Jazz#Metal#New Age#Oldies#Other#Pop#" & _
        "R&B#Rap#Reggae#Rock#Techno#Industrial#Alternative#Ska#Death Metal#Pranks#Soundtrack#Euro-Techno#" & _
        "Ambient#Trip-Hop#Vocal#Jazz+Funk#Fusion#Trance#Classical#Instrumental#Acid#House#Game#Sound Clip#" & _
        "Gospel#Noise#Alt. Rock#Bass#Soul#Punk#Space#Meditative#Instrumental Pop#Instrumental Rock#Ethnic#" & _
        "Gothic#Darkwave#Techno-Industrial#Electronic#Pop-Folk#Eurodance#Dream#Southern Rock#Comedy#Cult#" & _
        "Gangsta#Top 40#Christian Rap#Pop/Funk#Jungle#Native American#Cabaret#New Wave#Psychadelic#Rave#" & _
        "Showtunes#Trailer#Lo-Fi#Tribal#Acid Punk#Acid Jazz#Polka#Retro#Musical#Rock & Roll#Hard Rock#Folk#" & _
        "Folk-Rock#National Folk#Swing#Fast Fusion#Bebob#Latin#Revival#Celtic#Bluegrass#Avantgarde#Gothic Rock#" & _
        "Progressive Rock#Psychedelic Rock#Symphonic Rock#Slow Rock#Big Band#Chorus#Easy Listening#Acoustic#" & _
        "Humour#Speech#Chanson#Opera#Chamber Music#Sonata#Symphony#Booty Bass#Primus#Porn Groove#Satire#Slow Jam#" & _
        "Club#Tango#Samba#Folklore#Ballad#Power Ballad#Rhythmic Soul#Freestyle#Duet#Punk Rock#Drum Solo#A capella#" & _
        "Euro-House#Dance Hall#Goa#Drum & Bass#Club-House#Hardcore#Terror#Indie#BritPop#Negerpunk#Polsk Punk#" & _
        "Beat#Christian Gangsta#Heavy Metal#Black Metal#Crossover#Contemporary C#Christian Rock#Merengue#Salsa#" & _
        "Thrash Metal#Anime#JPop#SynthPop"
Private GenreField() As String, bBuiltGenre As Boolean
Attribute GenreField.VB_VarUserMemId = 1073938442
Attribute bBuiltGenre.VB_VarUserMemId = 1073938442


'API
Private Const GENERIC_READ = &H80000000
Private Const GENERIC_WRITE = &H40000000
Private Const FILE_SHARE_READ = &H1
Private Const OPEN_EXISTING = &H3

Private Const FILE_BEGIN = 0
Private Const FILE_CURRENT = 1
Private Const FILE_END = 2

Private Const FILE_ATTRIBUTE_READONLY = &H1
Private Const INVALID_HANDLE_VALUE = -1
Private Const ERROR_NO_MORE_FILES As Long = 18&

Private Type WIN32_FIND_DATA
    dwFileAttributes As Long
    ftCreationTime As FILETIME
    ftLastAccessTime As FILETIME
    ftLastWriteTime As FILETIME
    nFileSizeHigh As Long
    nFileSizeLow As Long
    dwReserved0 As Long
    dwReserved1 As Long
    cFileName As String * 260
    cAlternate As String * 14
End Type

Private Declare Function CreateFile Lib "kernel32" Alias "CreateFileA" (ByVal lpFileName As String, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, lpSecurityAttributes As Any, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As Long) As Long
Private Declare Function GetFileSize Lib "kernel32" (ByVal hFile As Long, lpFileSizeHigh As Long) As Long
Private Declare Function SetFilePointer Lib "kernel32" (ByVal hFile As Long, ByVal lDistanceToMove As Long, lpDistanceToMoveHigh As Long, ByVal dwMoveMethod As Long) As Long
Private Declare Function ReadFile Lib "kernel32" (ByVal hFile As Long, lpBuffer As Any, ByVal nNumberOfBytesToRead As Long, lpNumberOfBytesRead As Long, lpOverlapped As Any) As Long
Private Declare Function WriteFile Lib "kernel32" (ByVal hFile As Long, lpBuffer As Any, ByVal nNumberOfBytesToWrite As Long, lpNumberOfBytesWritten As Long, lpOverlapped As Any) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Declare Function SetEndOfFile Lib "kernel32" (ByVal hFile As Long) As Long
Private Declare Function GetFileAttributes Lib "kernel32" Alias "GetFileAttributesA" (ByVal lpFileName As String) As Long
Private Declare Function SetFileAttributes Lib "kernel32" Alias "SetFileAttributesA" (ByVal lpFileName As String, ByVal dwFileAttributes As Long) As Long
Private Declare Function FindFirstFile Lib "kernel32" Alias "FindFirstFileA" (ByVal lpFileName As String, lpFindFileData As WIN32_FIND_DATA) As Long

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Private Type SECURITY_ATTRIBUTES
    nLength As Long
    lpSecurityDescriptor As Long
    bInheritHandle As Long
End Type
Private Const FILE_FLAG_BACKUP_SEMANTICS = &H2000000


Public Property Let Read_MPEGInfo(Tag As Boolean)
    m_fReadMPEGInfo = Tag
End Property
Public Property Get Read_MPEGInfo() As Boolean
    Read_MPEGInfo = m_fReadMPEGInfo
End Property

Public Property Get MPEG_Version() As String
    MPEG_Version = myMPEGInfo.MPEGVersion
End Property


Public Property Get MPEG_FileSizeMB() As String
    MPEG_FileSizeMB = myMPEGInfo.SizeMB
End Property

Public Property Get MPEG_Bit_Rate() As Long
    MPEG_Bit_Rate = myMPEGInfo.bitrate
End Property

Public Property Get MPEG_Frequency() As Long
    MPEG_Frequency = myMPEGInfo.Frequency
End Property

Public Property Get MPEG_Emphasis() As String
    MPEG_Emphasis = myMPEGInfo.HasEmphasis
End Property
Public Property Get MPEG_Original() As String
    MPEG_Original = myMPEGInfo.IsOriginal
End Property
Public Property Get MPEG_Copyrighted() As String
    MPEG_Copyrighted = myMPEGInfo.IsCopyrighted
End Property
Public Property Get MPEG_CRCs() As String
    MPEG_CRCs = myMPEGInfo.HasCRC
End Property
Public Property Get MPEG_DurationTime() As String
    MPEG_DurationTime = myMPEGInfo.LengthTime

End Property
Public Property Get MPEG_VBR() As Boolean
    MPEG_VBR = myMPEGInfo.HasVBR
End Property
Public Property Get MPEG_ChannelMode() As String
    MPEG_ChannelMode = myMPEGInfo.ChannelMode
End Property


Public Property Get TagID3V1() As String
    If myMPEGInfo.ID3v1Version = -1 Then
        TagID3V1 = "none"
    Else
        TagID3V1 = "version 1." & myMPEGInfo.ID3v1Version
    End If
End Property

Public Property Get TagID3V2() As String
    If myMPEGInfo.ID3v2Version = -1 Then
        TagID3V2 = "none"
    Else
        TagID3V2 = "version 2." & myMPEGInfo.ID3v2Version
    End If
End Property


Public Property Get Has_Lyrics3_Tag() As Boolean
    Has_Lyrics3_Tag = HasLyrics3Tag
End Property

Public Property Get Lyrics() As String
    Lyrics = myOtherinfo.Lyrics
End Property

Public Property Let Lyrics(NewLyrics As String)
    myOtherinfo.Lyrics = NewLyrics
End Property

Public Property Get Quality() As String
    Quality = myOtherinfo.Quality
End Property

Public Property Let Quality(NewQuality As String)
    myOtherinfo.Quality = NewQuality
End Property

Public Property Get Situation() As String
    Situation = myOtherinfo.Situation
End Property

Public Property Let Situation(NewSituation As String)
    myOtherinfo.Situation = NewSituation
End Property

Public Property Get Mood() As String
    Mood = myOtherinfo.Mood
End Property

Public Property Let Mood(NewMood As String)
    myOtherinfo.Mood = NewMood
End Property

Public Property Get Author() As String
    Author = myOtherinfo.Author
End Property

Public Property Let Author(NewAuthor As String)
    myOtherinfo.Author = NewAuthor
End Property

Public Property Get OtherInfo() As String
    OtherInfo = myOtherinfo.OtherInfo
End Property

Public Property Let OtherInfo(NewOtherInfo As String)
    myOtherinfo.OtherInfo = NewOtherInfo
End Property

Public Property Get Images() As String
    Images = myOtherinfo.Images
End Property

Public Property Let Images(NewImages As String)
    myOtherinfo.Images = NewImages
End Property

Public Property Let Year(ByVal NewYear As String)
    myTag.SongYear = NewYear
End Property

Public Property Get Year() As String
    Year = myTag.SongYear
End Property

Public Property Get DurationInSecs()
    DurationInSecs = myMPEGInfo.Length
End Property

Public Property Let Genre(ByVal NewGenre As String)
    myTag.Genre = NewGenre
End Property

Public Property Get Genre() As String
    Genre = myTag.Genre
End Property


Public Property Let Comment(ByVal NewComment As String)
    myTag.Comment = NewComment
End Property

Public Property Get Comment() As String
    Comment = myTag.Comment
End Property

Public Property Let Album(ByVal NewAlbum As String)
    myTag.Album = NewAlbum
End Property

Public Property Get Album() As String
    Album = myTag.Album
End Property

Public Property Let TrackNr(ByVal NewTrack As String)
    myTag.TrackNr = NewTrack
End Property

Public Property Get TrackNr() As String
    TrackNr = myTag.TrackNr
End Property

Public Property Let Composer(ByVal NewComposer As String)
    myTag.Composer = NewComposer
End Property

Public Property Get Composer() As String
    Composer = myTag.Composer
End Property

Public Property Let OrigArtist(ByVal NewOrigArtist As String)
    myTag.OrigArtist = NewOrigArtist
End Property

Public Property Get OrigArtist() As String
    OrigArtist = myTag.OrigArtist
End Property

Public Property Let Copyright(ByVal NewCopyright As String)
    myTag.Copyright = NewCopyright
End Property

Public Property Get Copyright() As String
    Copyright = myTag.Copyright
End Property

Public Property Let LinkTo(ByVal NewlinkTo As String)
    myTag.LinkTo = NewlinkTo
End Property

Public Property Get LinkTo() As String
    LinkTo = myTag.LinkTo
End Property

Public Property Let EncodedBy(ByVal NewEncodedBy As String)
    myTag.EncodedBy = NewEncodedBy
End Property

Public Property Get EncodedBy() As String
    EncodedBy = myTag.EncodedBy
End Property

Public Property Let Artist(ByVal NewArtist As String)
    myTag.Artist = NewArtist
End Property

Public Property Get Artist() As String
    Artist = myTag.Artist
End Property

Public Property Let Title(ByVal NewTitle As String)
    myTag.Title = NewTitle
End Property

Public Property Get Title() As String
    Title = myTag.Title
End Property

Public Property Get FileName() As String
    FileName = m_sFileName
End Property

Public Property Get FileSize() As String
    FileSize = myMPEGInfo.FileSize
End Property


'**********************************ID3v 1.0, 1.1****************************************

Private Function ReadID3v1(ByVal FileName As String, ByRef outTag As ID3Tag) As Boolean
    Dim fh As Long, fsize As Long
    Dim RdTag As v1Tag

    fh = CreateFile(FileName, GENERIC_READ, FILE_SHARE_READ, ByVal 0, OPEN_EXISTING, 0, 0)
    If fh = INVALID_HANDLE_VALUE Then Exit Function
    fsize = GetFileSize(fh, 0)
    If fsize < 128 Then
        CloseHandle fh
        Exit Function
    End If
    'The ID3v1 tag is located 128 bytes from the end of the file, so look at this position.
    SetFilePointer fh, fsize - 128, 0, FILE_BEGIN
    ReadFile fh, RdTag, Len(RdTag), 0, ByVal 0
    CloseHandle fh

    If Data2String(VarPtr(RdTag.Identifier(0)), 3, ENC_ISO) = "TAG" Then
        ResetTag outTag

        'An ID3v1 tag is present.
        outTag.Artist = Trim$(Data2String(VarPtr(RdTag.Artist(0)), 30, ENC_ISO))
        outTag.Album = Trim$(Data2String(VarPtr(RdTag.Album(0)), 30, ENC_ISO))
        outTag.Title = Trim$(Data2String(VarPtr(RdTag.Title(0)), 30, ENC_ISO))
        If RdTag.Comment(28) = 0 And Not RdTag.Comment(29) = 0 Then
            'ID3v1.1 tag: byte 28 of the comment tag is zeroed, byte 29 is the track number.
            outTag.Comment = Trim$(Data2String(VarPtr(RdTag.Comment(0)), 28, ENC_ISO))
            outTag.TrackNr = RdTag.Comment(29)
        Else
            'ID3v1.0 tag: contains no track number.
            outTag.Comment = Trim$(Data2String(VarPtr(RdTag.Comment(0)), 30, ENC_ISO))
        End If
        outTag.Genre = GetGenreName(RdTag.Genre)
        outTag.SongYear = Trim$(Data2String(VarPtr(RdTag.SongYear(0)), 4, ENC_ISO))

        ReadID3v1 = True
    End If
End Function

Public Sub Read_File_Tags(sFile As String)
    Dim TAGv1 As ID3Tag
    Dim TAGv2 As ID3Tag

    ' On Error Resume Next
    m_sFileName = sFile

    ResetTag myTag
    myOtherinfo.Lyrics = ""
    myOtherinfo.Images = ""
    'If Dir(m_sFileName) = "" Then Exit Sub
    If Not FileExists(m_sFileName) Or UCase(getFileExtensionFromPath(sFile)) <> "MP3" Then Exit Sub    'the problem with dir is that it doesn't read hidden files so I had to use fileexist functio
    'fileExist() uses finfirstfile function to check existance of file which is much reliable and faster than vb's dir()

    'Reset the local tag structure, and read the ID3 v2 tag.

    'read tag only when it is mp3 file
    ReadID3v2 sFile, TAGv2

    '// Read ID3V1 Tags
    ReadID3v1 sFile, TAGv1

    myTag.Title = TAGv2.Title
    myTag.Artist = TAGv2.Artist
    myTag.Album = TAGv2.Album
    myTag.SongYear = TAGv2.SongYear
    myTag.Genre = TAGv2.Genre
    myTag.Comment = TAGv2.Comment


    If TAGv2.Title = "" Then myTag.Title = TAGv1.Title
    If myTag.Title = "" Then
        myTag.Title = Right(sFile, Len(sFile) - InStrRev(sFile, "\", , vbTextCompare))
        myTag.Title = Left(myTag.Title, Len(myTag.Title) - 4)
    End If
    If TAGv2.Artist = "" Then myTag.Artist = TAGv1.Artist
    If TAGv2.Album = "" Then myTag.Album = TAGv1.Album
    If TAGv2.SongYear = "" Then myTag.SongYear = TAGv1.SongYear
    If TAGv2.Genre = "" Then myTag.Genre = TAGv1.Genre
    If TAGv2.Comment = "" Then myTag.Comment = TAGv1.Comment

    '// buscar la letra de la rola y otra infomacion
    HasLyrics3Tag = GetLyrics3Tag(sFile, myOtherinfo)

    If m_fReadMPEGInfo = True Then
        ReadMPEGInfo sFile, myMPEGInfo
    End If

End Sub

Public Sub Write_File_Tags()
    On Error Resume Next
    If Dir(m_sFileName) = "" Then Exit Sub

    WriteID3v1 m_sFileName, myTag

    'The WriteID3v2 function can be passed a buffer pointer as parameter. See the comment
    'at the function implementation for details.
    WriteID3v2 m_sFileName, myTag, VERSION_2_3, False

    ' WriteLyrics3Tag   ' i have commented this as it is giving some problem with windows 7
    ' It kills some of mp3 file
End Sub

Private Function WriteLyrics3Tag() As Boolean

    Dim WholeTag As String
    Dim tagsize As String * 6
    Dim Position As Long
    Dim MoveMP3Tag As Boolean
    Dim fn As Integer
    Dim UseOldInfo As Boolean
    Dim NewLyr As Boolean
    Dim fh As Long, fsize As Long
    Dim RdTag As v1Tag
    Dim bID3v1Tag As Boolean


    WholeTag = ""

    On Error GoTo hell

    If myOtherinfo.Author <> Empty Then NewLyr = True
    If myOtherinfo.OtherInfo <> Empty Then NewLyr = True
    If myOtherinfo.Lyrics <> Empty Then NewLyr = True
    If myOtherinfo.Mood <> Empty Then NewLyr = True
    If myOtherinfo.Quality <> Empty Then NewLyr = True
    If myOtherinfo.Situation <> Empty Then NewLyr = True
    If myOtherinfo.Images <> "" Then NewLyr = True

    'build the tag
    If NewLyr = True Then
        fh = CreateFile(m_sFileName, GENERIC_READ, FILE_SHARE_READ, ByVal 0, OPEN_EXISTING, 0, 0)
        If fh = INVALID_HANDLE_VALUE Then Exit Function
        fsize = GetFileSize(fh, 0)
        If fsize < 128 Then
            CloseHandle fh
            Exit Function
        End If

        'The ID3v1 tag is located 128 bytes from the end of the file, so look at this position.
        SetFilePointer fh, fsize - 128, 0, FILE_BEGIN
        ReadFile fh, RdTag, Len(RdTag), 0, ByVal 0
        CloseHandle fh

        If Data2String(VarPtr(RdTag.Identifier(0)), 3, ENC_ISO) = "TAG" Then
            bID3v1Tag = True
        Else
            bID3v1Tag = False
        End If

        WholeTag = "LYRICSBEGIN"
        Dim IND As String
        'lyrics
        If myOtherinfo.Lyrics <> "" Then
            WholeTag = WholeTag & "LYR" & Format(Len(myOtherinfo.Lyrics), "00000") & myOtherinfo.Lyrics
        End If
        'other info
        If myOtherinfo.OtherInfo <> "" Then
            WholeTag = WholeTag & "INF" & Format(Len(myOtherinfo.OtherInfo), "00000") & myOtherinfo.OtherInfo
        End If
        'author
        If myOtherinfo.Author <> "" Then
            WholeTag = WholeTag & "AUT" & Format(Len(myOtherinfo.Author), "00000") & myOtherinfo.Author
        End If
        'images
        If myOtherinfo.Images <> "" Then
            WholeTag = WholeTag & "IMG" & Format(Len(myOtherinfo.Images), "00000") & myOtherinfo.Images
        End If
        'quality
        If myOtherinfo.Quality <> "" Then
            WholeTag = WholeTag & "QUA" & Format(Len(myOtherinfo.Quality), "00000") & myOtherinfo.Quality
        End If
        'situation
        If myOtherinfo.Situation <> "" Then
            WholeTag = WholeTag & "SIT" & Format(Len(myOtherinfo.Situation), "00000") & myOtherinfo.Situation
        End If
        'mood
        If myOtherinfo.Mood <> "" Then
            WholeTag = WholeTag & "MOO" & Format(Len(myOtherinfo.Mood), "00000") & myOtherinfo.Mood
        End If

        'calcuate the size
        tagsize = Format(Len(WholeTag), "000000")
        'append the end identifier
        WholeTag = WholeTag & tagsize & "LYRICS200"
    End If
    'prepare for writing
    fn = FreeFile
    Open m_sFileName For Binary As #fn

    'set to just before the current id3 tag if has.
    If bID3v1Tag = True Then
        Position = LOF(fn) - 127
    Else
        Position = LOF(fn) + 1
    End If

    'if there is a Lyrics3 tag, then go back to the beginning of the old one
    If HasLyrics3Tag Then Position = Val(PosTagLyr)

    ' write the lyrics3 tag if there is one...
    If WholeTag <> Empty Then
        Put #fn, Position, WholeTag
        Position = Seek(fn)
    End If
    'set the last byte of the file
    Position = Seek(fn) - 1
    Close
    'make sure this is the end of the file which is needed if this tag is smaller than the old tag.
    SetFileLength m_sFileName, Position

    Exit Function
hell:
End Function

Private Sub SetFileLength(ByVal FileName As String, ByVal NewLength As Long)
'Will cut the length of a file to the length specified.

    Dim hFile As Long
    Dim l As Long
    Dim lpSecurity As SECURITY_ATTRIBUTES
    'if file is smaller than or equal to requsted length, exit.
    If FileLen(FileName) <= NewLength Then Exit Sub
    'open the file
    hFile = CreateFile(FileName, GENERIC_WRITE, ZERO, lpSecurity, OPEN_EXISTING, FILE_FLAG_BACKUP_SEMANTICS, 0)
    'if file not open exit
    If hFile = -1 Then Exit Sub
    'seek to position
    l = SetFilePointer(hFile, NewLength, ZERO, ZERO)
    'and mark here as end of file
    SetEndOfFile hFile
    'close the file
    l = CloseHandle(hFile)
End Sub


Private Function WriteID3v1(ByVal FileName As String, ByRef inTag As ID3Tag, Optional ByVal bIgnoreReadOnly As Boolean = True) As Boolean
    Dim fh As Long, fsize As Long
    Dim rdBuf(2) As Byte, WrTag As v1Tag

    If bIgnoreReadOnly Then IgnoreReadOnly FileName
    fh = CreateFile(FileName, GENERIC_READ Or GENERIC_WRITE, FILE_SHARE_READ, ByVal 0, OPEN_EXISTING, 0, 0)
    If fh = INVALID_HANDLE_VALUE Then Exit Function
    fsize = GetFileSize(fh, 0)

    'Check if there's already an ID3v1 tag present.
    'The ID3v1 tag is located 128 bytes from the end of the file, so look at this position.
    If fsize >= 128 Then
        SetFilePointer fh, fsize - 128, 0, FILE_BEGIN
        ReadFile fh, rdBuf(0), 3, 0, ByVal 0
        If Data2String(VarPtr(rdBuf(0)), 3, ENC_ISO) = "TAG" Then
            'An ID3v1 tag is present.
            SetFilePointer fh, fsize - 128, 0, FILE_BEGIN
        Else
            'An ID3v1 tag isn't present.
            SetFilePointer fh, fsize, 0, FILE_BEGIN
        End If
    Else
        'An ID3v1 tag isn't present.
        SetFilePointer fh, fsize, 0, FILE_BEGIN
    End If

    'Build up writing tag.
    String2Data "TAG", 3, VarPtr(WrTag.Identifier(0)), False
    String2Data inTag.Artist, 30, VarPtr(WrTag.Artist(0)), False
    String2Data inTag.Album, 30, VarPtr(WrTag.Album(0)), False
    String2Data inTag.Title, 30, VarPtr(WrTag.Title(0)), False
    String2Data Replace$(inTag.Comment, vbNewLine, " "), 28, VarPtr(WrTag.Comment(0)), False
    'ID3v1.1 tag: comment(28) = 0, comment(29) = track number.
    WrTag.Comment(28) = 0
    WrTag.Comment(29) = 0
    If IsNumeric(inTag.TrackNr) And Len(Trim$(inTag.TrackNr)) <= 3 Then
        If CLng(inTag.TrackNr) > 0 And CLng(inTag.TrackNr) <= 255 Then
            WrTag.Comment(29) = CByte(inTag.TrackNr)
        End If
    End If
    WrTag.Genre = GetGenreIndex(inTag.Genre)
    String2Data inTag.SongYear, 4, VarPtr(WrTag.SongYear(0)), False

    'Write the tag to file.
    WriteFile fh, WrTag, Len(WrTag), 0, ByVal 0

    CloseHandle fh
    WriteID3v1 = True
End Function

Private Function DeleteID3v1(ByVal FileName As String, Optional ByVal bIgnoreReadOnly As Boolean = True) As Boolean
    Dim fh As Long, fsize As Long
    Dim rdBuf(2) As Byte

    If bIgnoreReadOnly Then IgnoreReadOnly FileName
    fh = CreateFile(FileName, GENERIC_READ Or GENERIC_WRITE, FILE_SHARE_READ, ByVal 0, OPEN_EXISTING, 0, 0)
    If fh = INVALID_HANDLE_VALUE Then Exit Function
    fsize = GetFileSize(fh, 0)
    If fsize < 128 Then
        CloseHandle fh
        Exit Function
    End If

    'Check if there's an ID3v1 tag present.
    'The ID3v1 tag is located 128 bytes from the end of the file, so look at this position.
    SetFilePointer fh, fsize - 128, 0, FILE_BEGIN
    ReadFile fh, rdBuf(0), 3, 0, ByVal 0
    If Data2String(VarPtr(rdBuf(0)), 3, ENC_ISO) = "TAG" Then
        'An ID3v1 tag is present, remove it.
        SetFilePointer fh, fsize - 128, 0, FILE_BEGIN
        SetEndOfFile fh
    End If

    CloseHandle fh
    DeleteID3v1 = True
End Function

Private Function GetLyrics3Tag(sFile As String, ByRef outOtherInfo As tOtherInfo) As Boolean
    On Error GoTo hell
    Dim Position As Long
    Dim FieldData As String   '// save Text of field
    Dim FieldID As String * 3    '// save Field ex: LYR, ALB, AUT... etc...
    Dim LengthField As String
    Dim Size As Long          '// size of File
    Dim tagtype As String * 9    '// save end tag lyrics
    Dim Byte11Buffer As String * 11    '// save LYRICSBEGIN
    Dim Byte5Buffer As String * 5   '// Length of Field
    Dim lEndLyr As Long
    Dim LyricEndID As String * 6
    Dim fh As Long, fsize As Long
    Dim RdTag As v1Tag
    Dim bID3v1Tag As Boolean

    fh = CreateFile(sFile, GENERIC_READ, FILE_SHARE_READ, ByVal 0, OPEN_EXISTING, 0, 0)
    If fh = INVALID_HANDLE_VALUE Then Exit Function
    fsize = GetFileSize(fh, 0)
    If fsize < 128 Then
        CloseHandle fh
        Exit Function
    End If

    'The ID3v1 tag is located 128 bytes from the end of the file, so look at this position.
    SetFilePointer fh, fsize - 128, 0, FILE_BEGIN
    ReadFile fh, RdTag, Len(RdTag), 0, ByVal 0
    CloseHandle fh

    If Data2String(VarPtr(RdTag.Identifier(0)), 3, ENC_ISO) = "TAG" Then
        bID3v1Tag = True
    Else
        bID3v1Tag = False
    End If

    '//reset size of tag
    LengthField = "000000"

    '//open the file
    fh = FreeFile
    Open sFile For Binary As #fh
    'get filesize
    Size = LOF(fh)

    'get the tag END
    If bID3v1Tag = True Then
        'get the tag END [ (127 -> ID3v1Tag) + (9  -> LYRICS200) ] = 136
        lEndLyr = Size - 136
        Get #fh, lEndLyr, tagtype
    Else
        lEndLyr = Size - 8
        Get #fh, lEndLyr, tagtype
    End If

    'if tag is valid then
    If tagtype = "LYRICS200" Then

        'get the size of the tag  ( 136 - 6 ) = 142
        Get #fh, lEndLyr - 6, LengthField

        'set the position to the first byte of Lyrics
        Position = lEndLyr - 6 - Val(LengthField)

        'get the beginning of tag
        Get #fh, Position, Byte11Buffer
        'save beginning lyrics tag
        PosTagLyr = Position

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
            Get #fh, Position, FieldID

            'the size of the field
            Get #fh, Position + 3, Byte5Buffer
            LengthField = Val(Byte5Buffer)

            'make room for the data
            FieldData = Space(Val(LengthField))
            'get the data
            Get #fh, Position + 8, FieldData
            'and fill the approprate field
            Select Case FieldID
            Case "LYR"
                outOtherInfo.Lyrics = Trim$(FieldData)
            Case "INF"
                outOtherInfo.OtherInfo = Trim$(FieldData)
            Case "AUT"
                outOtherInfo.Author = Trim$(FieldData)
            Case "EAL"
                outOtherInfo.Album = Trim$(FieldData)
            Case "EAR"
                outOtherInfo.Artist = Trim$(FieldData)
            Case "ETT"
                outOtherInfo.Title = Trim$(FieldData)
            Case "IMG"
                outOtherInfo.Images = Trim$(FieldData)
            Case "QUA"
                outOtherInfo.Quality = Trim$(FieldData)
            Case "SIT"
                outOtherInfo.Situation = Trim$(FieldData)
            Case "MOO"
                outOtherInfo.Mood = Trim$(FieldData)
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



'**********************************ID3v 2.2, 2.3, 2.4****************************************

Private Function ReadID3v2(ByVal FileName As String, ByRef outTag As ID3Tag) As Boolean
    Dim fh As Long, fp As Long, TagFooter As v2TagHeader
    Dim rdBuf(3) As Byte
    On Error GoTo hell
    fh = CreateFile(FileName, GENERIC_READ, FILE_SHARE_READ, ByVal 0, OPEN_EXISTING, 0, 0)
    If fh = INVALID_HANDLE_VALUE Then Exit Function

    'Search for an ID3v2 tag.
    Do
        SetFilePointer fh, fp, 0, FILE_BEGIN
        ReadFile fh, rdBuf(0), 3, 0, ByVal 0
        If Data2String(VarPtr(rdBuf(0)), 3, ENC_ISO) = "ID3" Then
            'We've got an ID3 v2 tag.
            fp = ReadV2Tag(fh, fp, outTag, False)
            ReadID3v2 = True
        Else
            fp = -1
        End If
    Loop While Not fp = -1

    If Not ReadID3v2 Then
        'No tag found at the beginning of the file.
        'Version 2.4 allows the ID3 tag to be appended to the audio data. So, look if we can
        'find such a tag, identified by a footer. According to id3.org, the ID3v2 footer should
        'be added before other tagging systems, therefore prepended to an possibly existing ID3v1 tag.

        fp = GetFileSize(fh, 0)
        SetFilePointer fh, fp - 128, 0, FILE_BEGIN
        ReadFile fh, rdBuf(0), 3, 0, ByVal 0
        If Data2String(VarPtr(rdBuf(0)), 3, ENC_ISO) = "TAG" Then
            'Look before the ID3 v1 tag.
            fp = fp - 128 - Len(TagFooter)
        Else
            'Look at the end of the file.
            fp = fp - Len(TagFooter)
        End If
        If fp < 0 Then fp = 0

        SetFilePointer fh, fp, 0, FILE_BEGIN
        ReadFile fh, TagFooter, Len(TagFooter), 0, ByVal 0
        If Data2String(VarPtr(TagFooter.Identifier(0)), 3, ENC_ISO) = "3DI" Then
            'Got a footer. Calculate the header position from the size.
            fp = fp - Data2Long(VarPtr(TagFooter.Size(0)), True) - Len(TagFooter)
            If Not fp < 0 Then
                ReadV2Tag fh, fp, outTag, False
                ReadID3v2 = True
            End If
        End If
    End If
hell:
    CloseHandle fh
End Function

'Returns the file pointer for the next tag, or -1 if there is no tag appended.
Private Function ReadV2Tag(ByVal fh As Long, ByVal fp As Long, ByRef RdTag As ID3Tag, ByVal bToMem As Boolean) As Long
    Dim TagSizes As v2TagSizes, arrp As Long
    Dim TagHeader As v2TagHeader, TagHeaderFlags As v2TagHeaderFlags
    Dim ExtHeader23 As v2_3ExtHeader, ExtHeader24 As v2_4ExtHeader
    Dim FraHeader22 As v2_2FrameHeader, FraHeader234 As v2_34FrameHeader
    Dim FraDataSize As Long, FraUnsynchDecr As Long, fraID As String, FraText As String
    Dim FraOffset As Long, bSkipFra As Boolean
    Dim TagData() As Byte, SizeBuf(3) As Byte


    ReadV2Tag = -1

    '------------------------------------------------------
    '********************* TAG HEADER *********************
    '------------------------------------------------------

    SetFilePointer fh, fp, 0, FILE_BEGIN
    ReadFile fh, TagHeader, Len(TagHeader), 0, ByVal 0
    If TagHeader.Version(0) < 2 Or TagHeader.Version(0) > 4 Then Exit Function
    'The size stored in the header excludes itself, and excludes the footer (if present).
    TagSizes.ReadSize = Data2Long(VarPtr(TagHeader.Size(0)), True)

    'v 2.2 flags: %ab000000 a = unsynch, b = compression (not implemented, ignored)
    'v 2.3 flags: %abc00000 a = unsynch, b = extended header, c = experimental (ignored)
    'v 2.4 flags: %abcd0000 a = unsynch, b = extended header, c = experimental (ignored), d = footer present
    If TagHeader.Flags And &H80& Then TagHeaderFlags.bUnsynch = True
    If TagHeader.Version(0) >= 3 Then
        If TagHeader.Flags And &H40& Then TagHeaderFlags.bExtHeader = True
    End If
    If TagHeader.Version(0) = 4 Then
        If TagHeader.Flags And &H10& Then
            TagHeaderFlags.bFooter = True
            TagSizes.FooterSize = Len(TagHeader)
            'Add the size of the footer (which is the same size than the header) to the read size.
            'Padding must not exist when a footer is present.
            TagSizes.ReadSize = TagSizes.ReadSize + TagSizes.FooterSize
        End If
    End If

    'Read in the rest of the tag.
    ReDim TagData(TagSizes.ReadSize - 1)
    ReadFile fh, TagData(0), TagSizes.ReadSize, 0, ByVal 0

    'Perform de-unsynchoronisation of the whole read tag, if nescessary.
    'In version 2.2 and 2.3, the unsynch is done for all data after the header if the flag is set.
    'In version 2.4, the unsynch is applied on a per frame base.
    If TagHeader.Version(0) <= 3 And TagHeaderFlags.bUnsynch Then
        TagSizes.ReadSize = TagSizes.ReadSize - DeUnsynch(VarPtr(TagData(0)), TagSizes.ReadSize)
    End If


    '---------------------------------------------------------------
    '********************* EXTENDED TAG HEADER *********************
    '---------------------------------------------------------------

    If TagHeaderFlags.bExtHeader Then
        If TagHeader.Version(0) = 3 Then
            'Ignore the v 2.3 extended header.
            CopyMemory ExtHeader23, TagData(0), Len(ExtHeader23)
            TagSizes.ExtHeaderSize = Data2Long(VarPtr(ExtHeader23.Size(0)), False) + 4
        ElseIf TagHeader.Version(0) = 4 Then
            'Read the extended header update flag.
            CopyMemory ExtHeader24, TagData(0), Len(ExtHeader24)
            TagSizes.ExtHeaderSize = Data2Long(VarPtr(ExtHeader24.Size(0)), True)
            If ExtHeader24.ExtendedFlags And &H40& Then TagHeaderFlags.bUpdateTag = True
        End If
    End If

    If Not TagHeaderFlags.bUpdateTag Then
        'Reset output tag.
        ResetTag RdTag
    End If


    '--------------------------------------------------
    '********************* FRAMES *********************
    '--------------------------------------------------

    arrp = TagSizes.ExtHeaderSize
    Do
        If TagHeader.Version(0) = 2 Then

            'Check if there's enough space to read another tag header. We don't have to take
            'care of a footer here because v 2.2 doesn't support footers.
            If arrp < 0 Or arrp > TagSizes.ReadSize - Len(FraHeader22) Then Exit Do

            'Get the frame header.
            CopyMemory FraHeader22, TagData(arrp), Len(FraHeader22)
            arrp = arrp + Len(FraHeader22)

            'Read out the frame ID.
            fraID = Data2String(VarPtr(FraHeader22.Ident(0)), 3, ENC_ISO, False)
            If fraID = String$(3, Chr(0)) Then Exit Do

            'Read out the frame data size.
            CopyMemory SizeBuf(1), FraHeader22.Size(0), 3
            FraDataSize = Data2Long(VarPtr(SizeBuf(0)), False)
            If arrp > TagSizes.ReadSize - FraDataSize Then Exit Do

            'Read out the frame itself.
            If bToMem Then
                'CRM: Encrypted meta frame. Only supported in v 2.2, not converted.
                If Not fraID = "CRM" And FraDataSize > 0 Then
                    'Store this frame in the memory buffer, we only need the data to update
                    'a currently existing tag.
                    ReDim Preserve mMemFrames(mMemFraCnt)
                    With mMemFrames(mMemFraCnt)
                        .ID = fraID
                        .Size = FraDataSize
                        ReDim .Data(.Size - 1)
                        CopyMemory .Data(0), TagData(arrp), .Size
                    End With
                    mMemFraCnt = mMemFraCnt + 1
                    ConvertV23 mMemFraCnt - 1
                End If
                FraText = ""
            Else
                FraText = ReadFrame(VarPtr(TagData(arrp)), fraID, FraDataSize)
            End If

        ElseIf TagHeader.Version(0) >= 3 Then

            'Check if there's enough space to read another tag header.
            If arrp < 0 Or arrp > TagSizes.ReadSize - TagSizes.FooterSize - Len(FraHeader234) Then Exit Do

            'Get the frame header.
            CopyMemory FraHeader234, TagData(arrp), Len(FraHeader234)
            arrp = arrp + Len(FraHeader234)

            'Read out the frame ID.
            fraID = Data2String(VarPtr(FraHeader234.Ident(0)), 4, ENC_ISO, False)
            If fraID = String$(4, Chr(0)) Then
                'Got the padding. Frame parsing is finished now.
                Exit Do
            End If

            'Read out the frame data size.
            'v 2.3: No synchsafe integer.
            'v 2.4: Synchsafe integer. However, there are sloppy programs which code version 2.4 with
            'non-synchsafe integers. Do a test therefore, by checking if the next frame is a valid one.
            FraDataSize = Data2Long(VarPtr(FraHeader234.Size(0)), IIf(TagHeader.Version(0) = 3, False, True))
            If arrp > TagSizes.ReadSize - TagSizes.FooterSize - FraDataSize Then Exit Do
            If arrp <= TagSizes.ReadSize - TagSizes.FooterSize - FraDataSize - 4 Then
                'Check if the next frame would be valid.
                If Not IsValidFrameName(VarPtr(TagData(arrp + FraDataSize))) Then
                    'Try to decode the number as non-synchsafe.
                    FraDataSize = Data2Long(VarPtr(FraHeader234.Size(0)), False)
                    If arrp > TagSizes.ReadSize - TagSizes.FooterSize - FraDataSize Then Exit Do
                End If
            End If

            'Read out the frame flags.
            FraOffset = 0
            bSkipFra = False
            FraUnsynchDecr = 0
            If TagHeader.Version(0) = 3 Then
                'v 2.3 frame header flags: %abc00000 %ijk00000
                'abc flags are ignored. i = compression, j = encryption, k = grouping
                If FraHeader234.Flags(1) And &HC0& Then
                    'Either compressed or encrypted frame. Skip it.
                    bSkipFra = True
                End If
                If FraHeader234.Flags(1) And &H20& Then
                    'This frame belongs to a group. I don't care about that, but we have to
                    'consider about the 1 byte added to the header which defines the group.
                    FraOffset = 1
                End If
            Else
                'v 2.4 frame header flags: %0abc0000 %0h00kmnp
                'abc flags are ignored. h = grouping, k = compression, m = encryption,
                'n = unsynch, p = data length indicator
                If FraHeader234.Flags(1) And &H40& Then
                    'See comment above about grouping.
                    FraOffset = 1
                End If
                If FraHeader234.Flags(1) And &HC& Then
                    'Either compressed or encrypted frame. Skip it.
                    bSkipFra = True
                End If
                If (FraHeader234.Flags(1) And &H2&) Or TagHeaderFlags.bUnsynch Then
                    If Not bSkipFra Then
                        'Unsynched frame. DeUnsynch it now.
                        FraUnsynchDecr = DeUnsynch(VarPtr(TagData(arrp)), FraDataSize)
                    End If
                End If
                If FraHeader234.Flags(1) And &H1& Then
                    'Data length indicator. Is ignored, but add 4 bytes offset, because it
                    'appends a synchsafe integer to the header.
                    FraOffset = FraOffset + 4
                End If
            End If

            'Read out the frame itself.
            If Not bSkipFra Then
                'Skipped frames also are NOT added to the memory, because the new frames
                'are written with a frame header that doesn't include the flags for
                'compression or encryption.
                If bToMem And Not fraID = "SEEK" Then
                    'Store this frame in the memory buffer, we only need the data to update
                    'a currently existing tag. The v 2.4 seek frame isn't copied to memory,
                    'because the new written file is going to have different offsets anyway.
                    'However, read out the frame to ensure processing of the appended tag.
                    If FraDataSize - FraOffset - FraUnsynchDecr > 0 Then
                        ReDim Preserve mMemFrames(mMemFraCnt)
                        With mMemFrames(mMemFraCnt)
                            .ID = fraID
                            .Size = FraDataSize - FraOffset - FraUnsynchDecr
                            ReDim .Data(.Size - 1)
                            CopyMemory .Data(0), TagData(arrp + FraOffset), .Size
                        End With
                        mMemFraCnt = mMemFraCnt + 1
                        ConvertV23 mMemFraCnt - 1
                    End If
                    FraText = ""
                Else
                    FraText = ReadFrame(VarPtr(TagData(arrp + FraOffset)), fraID, FraDataSize - FraOffset - FraUnsynchDecr)
                End If
            Else
                FraText = ""
            End If
        End If

        'Increase the array pointer.
        arrp = arrp + FraDataSize

        'Write the tags into the structure.
        If fraID = "TP1" Or fraID = "TPE1" Then
            RdTag.Artist = FraText
        ElseIf fraID = "TAL" Or fraID = "TALB" Then
            RdTag.Album = FraText
        ElseIf fraID = "TT2" Or fraID = "TIT2" Then
            RdTag.Title = FraText
        ElseIf fraID = "TRK" Or fraID = "TRCK" Then
            RdTag.TrackNr = FraText
        ElseIf fraID = "TDRC" Then
            'The v 2.4 timestamp frame includes the release year.
            RdTag.SongYear = Left$(FraText, 4)
        ElseIf fraID = "TYE" Or fraID = "TYER" Then
            RdTag.SongYear = FraText
        ElseIf fraID = "TCO" Or fraID = "TCON" Then
            RdTag.Genre = Frame2Genre(FraText)
        ElseIf fraID = "COM" Or fraID = "COM " Or fraID = "COMM" Then
            RdTag.Comment = FraText
        ElseIf fraID = "TEN" Or fraID = "TENC" Then
            RdTag.EncodedBy = FraText
        ElseIf fraID = "WXX" Or fraID = "WXXX" Then
            RdTag.LinkTo = FraText
        ElseIf fraID = "TCR" Or fraID = "TCOP" Then
            RdTag.Copyright = FraText
        ElseIf fraID = "TOA" Or fraID = "TOPE" Then
            RdTag.OrigArtist = FraText
        ElseIf fraID = "TCM" Or fraID = "TCOM" Then
            RdTag.Composer = FraText
        ElseIf fraID = "SEEK" Then
            If IsNumeric(FraText) Then
                'FraText actually has to be numeric, because its value is generated from Data2Long.
                ReadV2Tag = fp + Len(TagHeader) + TagSizes.ReadSize + CLng(FraText)
            End If
        End If
    Loop
End Function

Private Function ReadFrame(ByVal pData As Long, ByVal FrameID As String, ByVal Framesize As Long) As String
    Dim EncFormat As Byte, SkipOffset As Long

    If Framesize = 0 Then Exit Function

    CopyMemory EncFormat, ByVal pData, 1
    If Left$(FrameID, 1) = "T" Then
        'Text frame.
        ReadFrame = Data2String(pData + 1, Framesize - 1, EncFormat)
    ElseIf FrameID = "COMM" Or FrameID = "COM " Or FrameID = "COM" Then
        'Comment frame.
        'Skip the 3 bytes language identifier.
        SkipOffset = 4
        'Skip the content description.
        ReadFrame = Data2String(pData + SkipOffset, Framesize - SkipOffset, EncFormat)
        If EncFormat = ENC_ISO Or EncFormat = ENC_UNICODE_UTF8 Then
            SkipOffset = SkipOffset + Len(ReadFrame) + 1
        Else
            SkipOffset = SkipOffset + Len(ReadFrame) * 2 + 2
        End If
        'Read the actual text.
        ReadFrame = Data2String(pData + SkipOffset, Framesize - SkipOffset, EncFormat)
    ElseIf FrameID = "WXXX" Or FrameID = "WXX" Then
        'Link frame.
        SkipOffset = 1
        ReadFrame = Data2String(pData + SkipOffset, Framesize - SkipOffset, EncFormat)
        If EncFormat = ENC_ISO Or EncFormat = ENC_UNICODE_UTF8 Then
            SkipOffset = SkipOffset + Len(ReadFrame) + 1
        Else
            SkipOffset = SkipOffset + Len(ReadFrame) * 2 + 2
        End If
        ReadFrame = Data2String(pData + SkipOffset, Framesize - SkipOffset, EncFormat)
    ElseIf FrameID = "SEEK" Then
        If Framesize < 4 Then Exit Function
        'Got a seek frame (v 2.4 only). It points to the next tag in the file.
        'According to id3.org, it seems that this value isn't synchsafe... ??
        ReadFrame = Data2Long(pData, False)
    End If
End Function

Private Function IsValidFrameName(ByVal pData As Long) As Boolean
    Dim i As Long, curData As Byte

    For i = 0 To 3
        CopyMemory curData, ByVal pData + i, 1
        'Accept $00 $00 $00 $00 as a valid frame name, because this would be the start of the pattern.
        If Not curData = 0 And Not (curData >= Asc("A") And curData <= Asc("Z")) And Not (curData >= Asc("0") And curData <= Asc("9")) Then
            Exit Function
        End If
    Next i
    IsValidFrameName = True
End Function

'Unsynch is simply done by replacing &HFF &H00 with &HFF.
'Returns the number of &H00 bytes kicked out.
Private Function DeUnsynch(ByVal pData As Long, ByVal Length As Long) As Long
    Dim i As Long, subIdx As Long
    Dim curData(1) As Byte

    Do
        'The following line performs pData(subIdx) = pData(i)
        CopyMemory ByVal pData + subIdx, ByVal pData + i, 1
        subIdx = subIdx + 1
        If i <= Length - 2 Then
            CopyMemory curData(0), ByVal pData + i, 2
            If curData(0) = &HFF& And curData(1) = &H0& Then
                'Jump over the &H0& byte.
                i = i + 2
                DeUnsynch = DeUnsynch + 1
                If i = Length Then Exit Do
            Else
                i = i + 1
            End If
        Else
            Exit Do
        End If
    Loop
End Function

'Converts a frame in memory from v 2.2 to v 2.3 in order to safe it in a new version.
Private Sub ConvertV23(ByVal memIdx As Long)
    Dim i As Long
    Dim sPicType As String, expSize As Long

    If memIdx < 0 Or memIdx > mMemFraCnt - 1 Then Exit Sub
    With mMemFrames(memIdx)
        If .ID = "PIC" Then
            'Convert the picture type desc from v 2.2 to 2.3 (and 2.4)
            sPicType = UCase$(Data2String(VarPtr(.Data(1)), 3, ENC_ISO))
            If sPicType = "PNG" Then
                sPicType = "image/png"
            ElseIf sPicType = "JPG" Then
                sPicType = "image/jpeg"
            ElseIf Not sPicType = "-->" Then
                sPicType = "image/" & LCase$(sPicType)
            End If
            'Expand the data.
            expSize = Len(sPicType) - 2
            If expSize <= 0 Then Exit Sub
            .Size = .Size + expSize
            ReDim Preserve .Data(.Size - 1)
            'Copy the data back, because we need more space at the beginning for the string.
            For i = .Size - 1 To 4 + expSize Step -1
                .Data(i) = .Data(i - expSize)
            Next i
            'Write the string into the data field.
            String2Data sPicType, 0, VarPtr(.Data(1)), True
        ElseIf .ID = "LNK" Then
            'Convert the link frame from v 2.2 to 2.3 (and 2.4)
            sPicType = Data2String(VarPtr(.Data(0)), 3, ENC_ISO)
            .Size = .Size + 1
            ReDim Preserve .Data(.Size - 1)
            For i = .Size - 1 To 4 Step -1
                .Data(i) = .Data(i - 1)
            Next i
            String2Data GetNewFrameID(sPicType), 4, VarPtr(.Data(0)), False
        ElseIf .ID = "TDOR" Then
            'Convert the TDOR timestamp from 2.4 to 2.3, simply by using only the first
            '4 bytes, identifying the year.
            .Size = 4
        End If
        .ID = GetNewFrameID(.ID)
    End With
End Sub

Private Function GetNewFrameID(ByVal sID As String) As String
    If sID = "UFI" Then
        sID = "UFID"
    ElseIf sID = "TT1" Then
        sID = "TIT1"
    ElseIf sID = "TT2" Then
        sID = "TIT2"
    ElseIf sID = "TT3" Then
        sID = "TIT3"
    ElseIf sID = "TP1" Then
        sID = "TPE1"
    ElseIf sID = "TP2" Then
        sID = "TPE2"
    ElseIf sID = "TP3" Then
        sID = "TPE3"
    ElseIf sID = "TP4" Then
        sID = "TPE4"
    ElseIf sID = "TCM" Then
        sID = "TCOM"
    ElseIf sID = "TXT" Then
        sID = "TEXT"
    ElseIf sID = "TLA" Then
        sID = "TLAN"
    ElseIf sID = "TCO" Then
        sID = "TCON"
    ElseIf sID = "TAL" Then
        sID = "TALB"
    ElseIf sID = "TPA" Then
        sID = "TPOS"
    ElseIf sID = "TRK" Then
        sID = "TRCK"
    ElseIf sID = "TRC" Then
        sID = "TSRC"
    ElseIf sID = "TYE" Then
        sID = "TYER"
    ElseIf sID = "TDA" Then
        sID = "TDAT"
    ElseIf sID = "TIM" Then
        sID = "TIME"
    ElseIf sID = "TRD" Then
        sID = "TRDA"
    ElseIf sID = "TMT" Then
        sID = "TMED"
    ElseIf sID = "TFT" Then
        sID = "TFLT"
    ElseIf sID = "TBP" Then
        sID = "TBPM"
    ElseIf sID = "TCR" Then
        sID = "TCOP"
    ElseIf sID = "TPB" Then
        sID = "TPUB"
    ElseIf sID = "TEN" Then
        sID = "TENC"
    ElseIf sID = "TSS" Then
        sID = "TSSE"
    ElseIf sID = "TOF" Then
        sID = "TOFN"
    ElseIf sID = "TLE" Then
        sID = "TLEN"
    ElseIf sID = "TSI" Then
        sID = "TSIZ"
    ElseIf sID = "TDY" Then
        sID = "TDLY"
    ElseIf sID = "TKE" Then
        sID = "TKEY"
    ElseIf sID = "TOT" Then
        sID = "TOAL"
    ElseIf sID = "TOL" Then
        sID = "TOLY"
    ElseIf sID = "TOR" Or sID = "TDOR" Then
        sID = "TORY"
    ElseIf sID = "TXX" Then
        sID = "TXXX"
    ElseIf sID = "WAF" Then
        sID = "WOAF"
    ElseIf sID = "WAR" Then
        sID = "WOAR"
    ElseIf sID = "WAS" Then
        sID = "WOAS"
    ElseIf sID = "WCM" Then
        sID = "WCOM"
    ElseIf sID = "WCP" Then
        sID = "WCOP"
    ElseIf sID = "WPB" Then
        sID = "WPUB"
    ElseIf sID = "WXX" Then
        sID = "WXXX"
    ElseIf sID = "IPL" Or sID = "TIPL" Then
        sID = "IPLS"
    ElseIf sID = "MCI" Then
        sID = "MCDI"
    ElseIf sID = "ETC" Then
        sID = "ETCO"
    ElseIf sID = "MLL" Then
        sID = "MLLT"
    ElseIf sID = "STC" Then
        sID = "SYTC"
    ElseIf sID = "ULT" Then
        sID = "USLT"
    ElseIf sID = "SLT" Then
        sID = "SYLT"
    ElseIf sID = "COM" Then
        sID = "COMM"
    ElseIf sID = "RVA" Then
        sID = "RVAD"
    ElseIf sID = "EQU" Then
        sID = "EQUA"
    ElseIf sID = "REV" Then
        sID = "RVRB"
    ElseIf sID = "PIC" Then
        sID = "APIC"
    ElseIf sID = "GEO" Then
        sID = "GEOB"
    ElseIf sID = "CNT" Then
        sID = "PCNT"
    ElseIf sID = "POP" Then
        sID = "POPM"
    ElseIf sID = "BUF" Then
        sID = "RBUF"
    ElseIf sID = "CRA" Then
        sID = "AENC"
    ElseIf sID = "LNK" Then
        sID = "LINK"
    End If
    GetNewFrameID = sID
End Function


'The version parameter defines in which ID3v2 version the tag is written (2.3 or 2.4).
'You possibly will prefer v 2.3 because it has greater support by other programs and still
'provides all necessary functionality. For example, WinAmp only supports v 2.3.

'The pBuffer parameter can point to a static memory buffer which is used if the file has to
'be rewritten. The buffer should be at least as large as the file size. If no buffer is
'provided or if the buffer is too small, the memory is allocated dynamic, which is slower.
'Especially when tagging more files, it's recommended to allocate a byte array of some MB and
'pass VarPtr(Buffer(0)) for the pBuffer parameter and UBound(Buffer) + 1 for the buffer size.
'This is the fastest file rewriting method possible.
Private Function WriteID3v2(ByVal FileName As String, ByRef inTag As ID3Tag, ByVal Version As ID3v2_Version, ByVal bClearExistingTag As Boolean, Optional ByVal bIgnoreReadOnly As Boolean = True, Optional ByVal pbuffer As Long = 0, Optional ByVal BufferSize As Long = 0) As Boolean
    Dim i As Long, fh As Long, fp As Long, fsize As Long
    Dim TagHeader As v2TagHeader, ExistSize As Long, arrp As Long
    Dim WrData() As Byte, WrSize As Long, WrBuf() As Byte, BuildTag As ID3Tag
    Dim tmpHeader As v2_34FrameHeader, rdBuf(3) As Byte

    If bIgnoreReadOnly Then IgnoreReadOnly FileName
    fh = CreateFile(FileName, GENERIC_READ Or GENERIC_WRITE, FILE_SHARE_READ, ByVal 0, OPEN_EXISTING, 0, 0)
    If fh = INVALID_HANDLE_VALUE Then Exit Function

    'Never mind on possibly existing appended tags or footer tags. Just leave them in the file.
    ExistSize = GetTagSize(fh, 0)

    'Read the existing tag into memory if it shouldn't be cleared.
    mMemFraCnt = 0
    If Not ExistSize = 0 And Not bClearExistingTag Then
        Do
            SetFilePointer fh, fp, 0, FILE_BEGIN
            ReadFile fh, rdBuf(0), 3, 0, ByVal 0
            If Data2String(VarPtr(rdBuf(0)), 3, ENC_ISO) = "ID3" Then
                'We've got an ID3 v2 tag. Write it temporary in the BuildTag variable.
                fp = ReadV2Tag(fh, fp, BuildTag, True)
            Else
                fp = -1
            End If
        Loop While Not fp = -1
    End If

    BuildTag = inTag
    BuildTag.Genre = Genre2Frame(BuildTag.Genre)

    'Determine required size for the new tag.
    WrSize = Len(TagHeader)
    With BuildTag
        WrSize = WrSize + GetFrameSize("TPE1", .Artist, True)
        WrSize = WrSize + GetFrameSize("TALB", .Album, True)
        WrSize = WrSize + GetFrameSize("TIT2", .Title, True)
        WrSize = WrSize + GetFrameSize(IIf(Version = VERSION_2_3, "TYER", "TDRC"), .SongYear, True)
        WrSize = WrSize + GetFrameSize("TCON", .Genre, True)
        WrSize = WrSize + GetFrameSize("COMM", .Comment, True)
        WrSize = WrSize + GetFrameSize("TRCK", .TrackNr, True)
        WrSize = WrSize + GetFrameSize("TENC", .EncodedBy, True)
        WrSize = WrSize + GetFrameSize("WXXX", .LinkTo, True)
        WrSize = WrSize + GetFrameSize("TCOP", .Copyright, True)
        WrSize = WrSize + GetFrameSize("TCOM", .Composer, True)
        WrSize = WrSize + GetFrameSize("TOPE", .OrigArtist, True)
    End With
    'Add the size of the existing frames which should be preserved.
    For i = 0 To mMemFraCnt - 1
        If IsToStore(mMemFrames(i).ID) Then
            WrSize = WrSize + mMemFrames(i).Size + Len(tmpHeader)
        End If
    Next i

    If WrSize <= ExistSize Then
        'Overwrite the existing tag completely.
        WrSize = ExistSize
    Else
        'We have to rewrite the file anyway.
        'Add padding, and not too less. A few KB don't matter in a 3+ MB file, and the file won't
        'have to be rewritten soon.
        If WrSize Mod 2048 < 1024 Then
            WrSize = (WrSize \ 2048) * 2048 + 2048
        Else
            WrSize = (WrSize \ 2048) * 2048 + 4096
        End If
        WrSize = WrSize + Len(TagHeader)
    End If
    ReDim WrData(WrSize - 1)


    '***** Build up tag header *****
    String2Data "ID3", 3, VarPtr(TagHeader.Identifier(0)), False
    TagHeader.Version(0) = IIf(Version = VERSION_2_3, 3, 4)
    Long2Data WrSize - Len(TagHeader), True, VarPtr(TagHeader.Size(0))
    CopyMemory WrData(0), TagHeader, Len(TagHeader)
    arrp = Len(TagHeader)

    '***** No extended header is used, since it even seems that many programs don't support them *****

    '***** Start writing the frame data into the array *****
    With BuildTag
        'Note: we don't use VarPtr(WrData(arrp)) here because if arrp == <array size>, the function
        'call would fail, although there is no more data written to the array.
        arrp = arrp + WriteFrame(VarPtr(WrData(0)) + arrp, "TPE1", .Artist, Version)
        arrp = arrp + WriteFrame(VarPtr(WrData(0)) + arrp, "TALB", .Album, Version)
        arrp = arrp + WriteFrame(VarPtr(WrData(0)) + arrp, "TIT2", .Title, Version)
        arrp = arrp + WriteFrame(VarPtr(WrData(0)) + arrp, IIf(Version = VERSION_2_3, "TYER", "TDRC"), .SongYear, Version)
        arrp = arrp + WriteFrame(VarPtr(WrData(0)) + arrp, "TCON", .Genre, Version)
        arrp = arrp + WriteFrame(VarPtr(WrData(0)) + arrp, "COMM", .Comment, Version)
        arrp = arrp + WriteFrame(VarPtr(WrData(0)) + arrp, "TRCK", .TrackNr, Version)
        arrp = arrp + WriteFrame(VarPtr(WrData(0)) + arrp, "TENC", .EncodedBy, Version)
        arrp = arrp + WriteFrame(VarPtr(WrData(0)) + arrp, "WXXX", .LinkTo, Version)
        arrp = arrp + WriteFrame(VarPtr(WrData(0)) + arrp, "TCOP", .Copyright, Version)
        arrp = arrp + WriteFrame(VarPtr(WrData(0)) + arrp, "TCOM", .Composer, Version)
        arrp = arrp + WriteFrame(VarPtr(WrData(0)) + arrp, "TOPE", .OrigArtist, Version)
    End With
    For i = 0 To mMemFraCnt - 1
        If IsToStore(mMemFrames(i).ID) Then
            arrp = arrp + WriteMemFrame(VarPtr(WrData(0)) + arrp, i, Version)
        End If
    Next i

    '***** Padding is added "automatically" because the WrData() array is zeroed from VB *****


    'Write the data out to the file.
    If WrSize <= ExistSize Then
        'Here we go, just overwrite the existing tag.
        SetFilePointer fh, 0, 0, FILE_BEGIN
        WriteFile fh, WrData(0), WrSize, 0, ByVal 0
    Else
        'We need to rewrite the entire file...
        fsize = GetFileSize(fh, 0)
        If pbuffer = 0 Or BufferSize < fsize - ExistSize Then
            'Allocate a dynamic buffer.
            ReDim WrBuf(fsize - ExistSize - 1)
            pbuffer = VarPtr(WrBuf(0))
        End If

        'Use the buffer to store the audio data of the file.
        SetFilePointer fh, ExistSize, 0, FILE_BEGIN
        ReadFile fh, ByVal pbuffer, fsize - ExistSize, 0, ByVal 0
        'Compute the new file size and expand the existing file.
        fsize = WrSize + fsize - ExistSize
        SetFilePointer fh, fsize, 0, FILE_BEGIN
        SetEndOfFile fh
        'Write the new tag to the file.
        SetFilePointer fh, 0, 0, FILE_BEGIN
        WriteFile fh, WrData(0), WrSize, 0, ByVal 0
        'Append the audio data again.
        WriteFile fh, ByVal pbuffer, fsize - WrSize, 0, ByVal 0
    End If

    CloseHandle fh
    WriteID3v2 = True
End Function

'Determines if a frame in memory has to be stored in the new written file.
Private Function IsToStore(ByVal fraID As String) As Boolean
    IsToStore = True
    If fraID = "TPE1" Or fraID = "TALB" Or fraID = "TIT2" Then
        IsToStore = False
    ElseIf fraID = "TYER" Or fraID = "TDRC" Or fraID = "TCON" Or fraID = "COMM" Then
        IsToStore = False
    ElseIf fraID = "TRCK" Or fraID = "TENC" Or fraID = "WXXX" Or fraID = "TCOP" Then
        IsToStore = False
    ElseIf fraID = "TCOM" Or fraID = "TOPE" Then
        IsToStore = False
    End If
End Function

Private Function WriteFrame(ByVal pData As Long, ByVal FrameID As String, ByVal strData As String, ByVal Version As ID3v2_Version) As Long
    Dim i As Long, fraHeader As v2_34FrameHeader
    Dim EncFormat As Byte, WrOffset As Long

    If Trim$(strData) = "" Then Exit Function

    'Create the frame header.
    String2Data FrameID, 4, VarPtr(fraHeader.Ident(0)), False
    WriteFrame = GetFrameSize(FrameID, strData, False)
    Long2Data WriteFrame, IIf(Version = VERSION_2_3, False, True), VarPtr(fraHeader.Size(0))
    CopyMemory ByVal pData, fraHeader, Len(fraHeader)
    WrOffset = Len(fraHeader)

    'Add the size of the header to the return value.
    WriteFrame = WriteFrame + Len(fraHeader)

    'Create the frame data.
    If Left$(FrameID, 1) = "T" Then
        'Text frame.
        'EncFormat = 0 = ISO encoding
        CopyMemory ByVal pData + WrOffset, EncFormat, 1
        WrOffset = WrOffset + 1
        String2Data strData, 0, pData + WrOffset, True
    ElseIf FrameID = "COMM" Then
        'Comment frame.
        CopyMemory ByVal pData + WrOffset, EncFormat, 1
        WrOffset = WrOffset + 1
        'English language.
        String2Data "eng", 0, pData + WrOffset, False
        WrOffset = WrOffset + 3
        'Create an empty content description.
        CopyMemory ByVal pData + WrOffset, EncFormat, 1
        WrOffset = WrOffset + 1
        String2Data strData, 0, pData + WrOffset, True
    ElseIf FrameID = "WXXX" Then
        'Link frame.
        For i = 0 To 1
            'i=0: encoding format, i=1: empty description string
            CopyMemory ByVal pData + WrOffset, EncFormat, 1
            WrOffset = WrOffset + 1
        Next i
        String2Data strData, 0, pData + WrOffset, True
    Else
        Exit Function
    End If
End Function

Private Function WriteMemFrame(ByVal pData As Long, ByVal fraIdx As Long, ByVal Version As ID3v2_Version) As Long
    Dim i As Long, fraHeader As v2_34FrameHeader
    Dim EncFormat As Byte, WrOffset As Long

    If fraIdx < 0 Or fraIdx > mMemFraCnt - 1 Then Exit Function
    With mMemFrames(fraIdx)
        If Version = VERSION_2_4 Then
            If .ID = "TORY" Then
                .ID = "TDOR"
            ElseIf .ID = "IPLS" Then
                .ID = "TIPL"
            End If
        End If
        'Create the frame header.
        String2Data .ID, 4, VarPtr(fraHeader.Ident(0)), False
        WriteMemFrame = .Size
        Long2Data WriteMemFrame, IIf(Version = VERSION_2_3, False, True), VarPtr(fraHeader.Size(0))
        CopyMemory ByVal pData, fraHeader, Len(fraHeader)
        WrOffset = Len(fraHeader)

        'Add the size of the header to the return value.
        WriteMemFrame = WriteMemFrame + Len(fraHeader)

        'Create the frame data.
        CopyMemory ByVal pData + WrOffset, .Data(0), .Size
    End With
End Function

Private Function GetFrameSize(ByVal FrameID As String, ByVal strData As String, ByVal bIncludeHeader As Boolean) As Long
    Dim fraHeader As v2_34FrameHeader

    If Trim$(strData) = "" Then Exit Function

    If bIncludeHeader Then GetFrameSize = Len(fraHeader)
    If Left$(FrameID, 1) = "T" Then
        'Text frame: 1 b encoding, x b string, 1 b null
        GetFrameSize = GetFrameSize + Len(strData) + 2
    ElseIf FrameID = "COMM" Then
        'Comment frame: 1 b encoding, 3 b language, 1 b empty description, x b string, 1 b null
        GetFrameSize = GetFrameSize + Len(strData) + 6
    ElseIf FrameID = "WXXX" Then
        'Link frame: 1 b encoding, 1 b empty description, x b string, 1 b null
        GetFrameSize = GetFrameSize + Len(strData) + 3
    Else
        GetFrameSize = 0
    End If
End Function


'See the WriteID3v2 function for a description on the Buffer parameters.
Private Function DeleteID3v2(ByVal FileName As String, Optional ByVal bIgnoreReadOnly As Boolean = True, Optional ByVal pbuffer As Long = 0, Optional ByVal BufferSize As Long = 0) As Boolean
    Dim fh As Long, fsize As Long
    Dim TagHeader As v2TagHeader, ExistSize As Long, WrBuf() As Byte

    If bIgnoreReadOnly Then IgnoreReadOnly FileName
    fh = CreateFile(FileName, GENERIC_READ Or GENERIC_WRITE, FILE_SHARE_READ, ByVal 0, OPEN_EXISTING, 0, 0)
    If fh = INVALID_HANDLE_VALUE Then Exit Function

    'Never mind on possibly existing appended tags or footer tags. Just leave them in the file.
    ExistSize = GetTagSize(fh, 0)
    If Not ExistSize = 0 Then
        fsize = GetFileSize(fh, 0)
        If pbuffer = 0 Or BufferSize < fsize - ExistSize Then
            'Allocate a dynamic buffer.
            ReDim WrBuf(fsize - ExistSize - 1)
            pbuffer = VarPtr(WrBuf(0))
        End If

        'Use the buffer to store the audio data of the file.
        SetFilePointer fh, ExistSize, 0, FILE_BEGIN
        ReadFile fh, ByVal pbuffer, fsize - ExistSize, 0, ByVal 0
        'Compute the new file size and crop the existing file.
        fsize = fsize - ExistSize
        SetFilePointer fh, fsize, 0, FILE_BEGIN
        SetEndOfFile fh
        'Write the audio data to the file again.
        SetFilePointer fh, 0, 0, FILE_BEGIN
        WriteFile fh, ByVal pbuffer, fsize, 0, ByVal 0
    End If

    CloseHandle fh
    DeleteID3v2 = True
End Function





'**********************************Data converting functions****************************************


'Synchsafe integers are integers that keep its highest bit (bit 7) zeroed, making seven bits
'out of eight available. Thus a 32 bit synchsafe integer can store 28 bits of information.
Private Function Data2Long(ByVal pData As Long, ByVal bSynchSafe As Boolean) As Long
    Dim i As Integer, Data(3) As Byte

    CopyMemory Data(0), ByVal pData, 4

    'Avoid converting wrong synchsafe integers. If bit 7 of any byte is set, it is not synchsafe.
    'However, we can't detect wrong coded values which have bit 7 zeroed.
    For i = 0 To 3
        If Data(i) And &H80& Then bSynchSafe = False
    Next i

    'Perform left-shifts, done by multiplication with the hex values of 2^n. Finally, bit-or the values.
    If bSynchSafe Then
        Data2Long = (Data(0) * &H200000) Or (Data(1) * &H4000&) Or (Data(2) * &H80&) Or Data(3)
    Else
        Data2Long = (Data(0) * &H1000000) Or (Data(1) * &H10000) Or (Data(2) * &H100&) Or Data(3)
    End If
End Function

Private Function Long2Data(ByVal SrcValue As Long, ByVal bSynchSafe As Boolean, ByVal pData As Long)
    Dim Data(3) As Byte

    'Mask out the corresponding bits and right-shift them afterwards (done by division).
    If bSynchSafe Then
        Data(0) = (SrcValue And &HFE00000) \ &H200000
        Data(1) = (SrcValue And &H1FC000) \ &H4000&
        Data(2) = (SrcValue And &H3F80&) \ &H80&
        Data(3) = SrcValue And &H7F&
    Else
        Data(0) = SrcValue \ &H1000000
        Data(1) = (SrcValue And &HFF0000) \ &H10000
        Data(2) = (SrcValue And &HFF00&) \ &H100&
        Data(3) = SrcValue And &HFF&
    End If

    'Copy back.
    CopyMemory ByVal pData, Data(0), 4
End Function

Private Function Data2String(ByVal pData As Long, ByVal Length As Long, ByVal EncFormat As v2_StrEncoding, Optional ByVal BreakOnNull As Boolean = True) As String
    Dim i As Long, curData As Byte, curSign As String

    For i = 0 To Length - 1
        CopyMemory curData, ByVal pData + i, 1
        'New lines are represented by &0A& (which is chr(10)) only in ID3 v2 tags.
        'However, many programs seem to code the newline with chr(13) & chr(10),
        'which is the windows default. Therefore, skip chr(13) and change chr(10) to vbNewLines.
        If curData = 13 Then
            curSign = ""
        ElseIf curData = 10 Then
            curSign = vbNewLine
        Else
            curSign = Chr(curData)
        End If
        If EncFormat = ENC_ISO Or EncFormat = ENC_UNICODE_UTF8 Then
            'Clear text, null terminated.
            If curData = 0 And BreakOnNull Then
                Exit Function
            Else
                Data2String = Data2String & curSign
            End If
        ElseIf EncFormat = ENC_UNICODE_UTF16_BOM Then
            'UNICODE text with BOM, double-null terminated.
            If i >= 2 And i Mod 2 = 0 Then
                If curData = 0 And BreakOnNull Then
                    Exit Function
                Else
                    Data2String = Data2String & curSign
                End If
            End If
        ElseIf EncFormat = ENC_UNICODE_UTF16 Then
            'UNICODE text without BOM, double-null terminated.
            If i Mod 2 = 0 Then
                If curData = 0 And BreakOnNull Then
                    Exit Function
                Else
                    Data2String = Data2String & curSign
                End If
            End If
        End If
    Next i
End Function

Private Function String2Data(ByVal SrcStr As String, ByVal MaxLength As Long, ByVal pData As Long, ByVal bTerminate As Boolean) As Long
    Dim i As Long, SrcLen As Long
    Dim curAsc As Byte

    SrcLen = Len(SrcStr)
    If MaxLength <= 0 Then
        MaxLength = SrcLen
    ElseIf MaxLength > SrcLen Then
        MaxLength = SrcLen
    End If

    For i = 0 To MaxLength - 1
        curAsc = Asc(Mid$(SrcStr, i + 1, 1))
        CopyMemory ByVal pData + i, curAsc, 1
    Next i
    If bTerminate Then
        'Add the null terminator.
        curAsc = 0
        CopyMemory ByVal pData + MaxLength, curAsc, 1
        String2Data = MaxLength + 1
    Else
        String2Data = MaxLength
    End If
End Function

Private Function Frame2Genre(ByVal strData As String) As String
    Dim idxFrom As Long, idxTo As Long
    Dim strRef As String

    Frame2Genre = strData
    idxFrom = InStr(1, strData, "(")
    idxTo = InStr(1, strData, ")")
    If idxFrom = 0 Or idxTo = 0 Then Exit Function

    strRef = Mid$(Frame2Genre, idxFrom + 1, idxTo - idxFrom - 1)
    If IsNumeric(strRef) Then
        Frame2Genre = GetGenreName(CLng(strRef))
    ElseIf strRef = "RX" Then
        Frame2Genre = "Remix"
    ElseIf strRef = "CR" Then
        Frame2Genre = "Cover"
    End If
End Function

Private Function Genre2Frame(ByVal strGenre As String) As String
    Dim GenreIdx As Long

    GenreIdx = GetGenreIndex(strGenre)
    If GenreIdx = 255 Then
        Genre2Frame = strGenre
    Else
        Genre2Frame = "(" & GenreIdx & ")" & GetGenreName(GenreIdx)
    End If
End Function

Public Function GetGenreName(ByVal GenreIndex As Long) As String
    If GenreIndex < 0 Or GenreIndex > 147 Then Exit Function

    If Not bBuiltGenre Then
        bBuiltGenre = True
        GenreField = Split(cGenreString, "#")
    End If
    GetGenreName = GenreField(GenreIndex)
End Function

Public Function GetGenreIndex(ByVal GenreName As String) As Long
    Dim i As Long

    If Not bBuiltGenre Then
        bBuiltGenre = True
        GenreField = Split(cGenreString, "#")
    End If
    GenreName = LCase$(Trim$(GenreName))
    For i = 0 To 147
        If LCase$(GenreField(i)) = GenreName Then
            GetGenreIndex = i
            Exit Function
        End If
    Next i
    GetGenreIndex = 255
End Function




'**********************************File information****************************************

Private Function ReadMPEGInfo(ByVal FileName As String, ByRef outInfo As MPEGInfo) As Boolean
    On Error GoTo errorHandler:
    Dim fh As Long, fp As Long, fsize As Long
    Dim fraHeader(3) As Byte, mVers As Long, lVers As Long
    Dim opByte As Byte, opLong As Long, VBRHeaderPos(1) As Long
    Dim VBRHeader(17) As Byte, VBRIdent As String
    Dim SamplesPerFrame As Long, VBRNrFrames As Long, VBRBytes As Long
    Dim RdTag As v1Tag, TagHeader As v2TagHeader
    Dim HH As Long, MM As Long, ss As Long, sTempTime As String, lSeconds As Long
    Dim dFilesize As Double

    fh = CreateFile(FileName, GENERIC_READ, FILE_SHARE_READ, ByVal 0, OPEN_EXISTING, 0, 0)
    If fh = INVALID_HANDLE_VALUE Then Exit Function

    fp = GetTagSize(fh, 0)
    fsize = GetFileSize(fh, 0)
    If fsize < 4 Then
        CloseHandle fh
        Exit Function
    End If

    'Search for the MP3 frame header.
    Do
        SetFilePointer fh, fp, 0, FILE_BEGIN
        ReadFile fh, fraHeader(0), 4, 0, ByVal 0
        If fraHeader(0) = &HFF& And (fraHeader(1) And &HE0&) = &HE0& Then
            'Got the frame synchronisation (11 bits set).
            'Read the information from the set bits.

            'MPEG version
            opByte = (fraHeader(1) And &H18&) \ &H8&
            If opByte = 0 Then
                outInfo.MPEGVersion = "version 2.5"
                mVers = 3
            ElseIf opByte = 2 Then
                outInfo.MPEGVersion = "version 2"
                mVers = 2
            ElseIf opByte = 3 Then
                outInfo.MPEGVersion = "version 1"
                mVers = 1
            Else
                Exit Do
            End If

            'Layer version
            opByte = (fraHeader(1) And &H6&) \ &H2&
            If opByte = 1 Then
                lVers = 3
                SamplesPerFrame = IIf(mVers = 1, 1152, 576)
            ElseIf opByte = 2 Then
                lVers = 2
                SamplesPerFrame = 1152
            ElseIf opByte = 3 Then
                lVers = 1
                SamplesPerFrame = 384
            Else
                Exit Do
            End If
            outInfo.MPEGVersion = outInfo.MPEGVersion & " layer " & lVers

            'CRC
            outInfo.HasCRC = IIf(fraHeader(1) And &H1&, False, True)

            'Channel mode.
            opByte = (fraHeader(3) And &HC0&) \ &H40&

            'The position of an possibly existing VBR header depends on the channel mode.
            If Not opByte = 3 Then
                VBRHeaderPos(0) = fp + 4 + IIf(mVers = 1, 32, 17)
            Else
                VBRHeaderPos(0) = fp + 4 + IIf(mVers = 1, 17, 9)
            End If
            VBRHeaderPos(1) = fp + 4 + 32

            'Channel mode.
            If opByte = 0 Then
                outInfo.ChannelMode = "Stereo"
            ElseIf opByte = 1 Then
                outInfo.ChannelMode = "Joint stereo"
            ElseIf opByte = 2 Then
                outInfo.ChannelMode = "Dual channel"
            ElseIf opByte = 3 Then
                outInfo.ChannelMode = "Mono"
            End If

            'Frequency
            opByte = (fraHeader(2) And &HC&) \ &H4&
            If opByte = 0 Then
                opLong = 44100
            ElseIf opByte = 1 Then
                opLong = 48000
            ElseIf opByte = 2 Then
                opLong = 32000
            End If
            outInfo.Frequency = opLong / IIf(mVers = 3, 4, mVers)

            'Check if there is a VBR header. If present, use it to read out the number of frames.
            SetFilePointer fh, VBRHeaderPos(0), 0, FILE_BEGIN
            ReadFile fh, VBRHeader(0), 16, 0, ByVal 0
            VBRIdent = Data2String(VarPtr(VBRHeader(0)), 4, ENC_ISO)
            If VBRIdent = "Xing" Or VBRIdent = "Info" Then
                'A VBR header is present.
                If VBRHeader(7) And &H1& Then
                    'The number of frames is stored.
                    VBRNrFrames = Data2Long(VarPtr(VBRHeader(8)), False)
                    If VBRHeader(7) And &H2& Then
                        VBRBytes = Data2Long(VarPtr(VBRHeader(12)), False)
                    End If
                End If
            End If
            'Check if there is a VBRI header.
            SetFilePointer fh, VBRHeaderPos(1), 0, FILE_BEGIN
            ReadFile fh, VBRHeader(0), 18, 0, ByVal 0
            If Data2String(VarPtr(VBRHeader(0)), 4, ENC_ISO) = "VBRI" Then
                VBRBytes = Data2Long(VarPtr(VBRHeader(10)), False)
                VBRNrFrames = Data2Long(VarPtr(VBRHeader(14)), False)
            End If

            If Not VBRBytes = 0 And Not VBRNrFrames = 0 Then
                'VBR bitrate.
                outInfo.HasVBR = True
                outInfo.bitrate = VBRBytes / VBRNrFrames / SamplesPerFrame / 125 * outInfo.Frequency
                outInfo.Length = VBRNrFrames / outInfo.Frequency * SamplesPerFrame
            Else
                'CBR bitrate
                outInfo.HasVBR = False
                opByte = (fraHeader(2) And &HF0&) \ &H10&
                If Not opByte = 0 And Not opByte = 15 Then
                    If mVers = 1 And lVers = 1 Then
                        outInfo.bitrate = opByte * 32
                    ElseIf mVers = 1 And lVers = 2 Then
                        If opByte = 1 Then
                            outInfo.bitrate = 32
                        ElseIf opByte = 2 Then
                            outInfo.bitrate = 48
                        ElseIf opByte <= 4 Then
                            outInfo.bitrate = 48 + (opByte - 2) * 8
                        ElseIf opByte <= 8 Then
                            outInfo.bitrate = 64 + (opByte - 4) * 16
                        ElseIf opByte <= 12 Then
                            outInfo.bitrate = 128 + (opByte - 8) * 32
                        Else
                            outInfo.bitrate = 256 + (opByte - 12) * 64
                        End If
                    ElseIf mVers = 1 And lVers = 3 Then
                        If opByte = 1 Then
                            outInfo.bitrate = 32
                        ElseIf opByte <= 5 Then
                            outInfo.bitrate = 32 + (opByte - 1) * 8
                        ElseIf opByte <= 9 Then
                            outInfo.bitrate = 64 + (opByte - 5) * 16
                        ElseIf opByte <= 13 Then
                            outInfo.bitrate = 128 + (opByte - 9) * 32
                        Else
                            outInfo.bitrate = 320
                        End If
                    ElseIf mVers >= 2 And lVers = 1 Then
                        If opByte = 1 Then
                            outInfo.bitrate = 32
                        ElseIf opByte = 2 Then
                            outInfo.bitrate = 48
                        ElseIf opByte <= 4 Then
                            outInfo.bitrate = 48 + (opByte - 2) * 8
                        ElseIf opByte <= 12 Then
                            outInfo.bitrate = 64 + (opByte - 4) * 16
                        Else
                            outInfo.bitrate = 192 + (opByte - 12) * 32
                        End If
                    Else
                        'mVers >= 2, lVers >= 2
                        If opByte <= 8 Then
                            outInfo.bitrate = opByte * 8
                        Else
                            outInfo.bitrate = 64 + (opByte - 8) * 16
                        End If
                    End If
                End If
                If outInfo.bitrate <> 0 Then outInfo.Length = (fsize - fp) / (outInfo.bitrate * 125)
            End If

            'Copyright
            outInfo.IsCopyrighted = IIf(fraHeader(3) And &H8&, True, False)

            'Original
            outInfo.IsOriginal = IIf(fraHeader(3) And &H4&, True, False)

            'Emphasis
            outInfo.HasEmphasis = IIf(fraHeader(3) And &H3&, True, False)

            ReadMPEGInfo = True
            Exit Do
        Else
            fp = fp + 1
            If fp > fsize - 4 Then Exit Do
        End If
    Loop

    'Read out the tag versions.
    'ID3 v1 tag.
    outInfo.ID3v1Version = -1
    If fsize >= 128 Then
        'The ID3v1 tag is located 128 bytes from the end of the file, so look at this position.
        SetFilePointer fh, fsize - 128, 0, FILE_BEGIN
        ReadFile fh, RdTag, Len(RdTag), 0, ByVal 0
        If Data2String(VarPtr(RdTag.Identifier(0)), 3, ENC_ISO) = "TAG" Then
            If RdTag.Comment(28) = 0 And Not RdTag.Comment(29) = 0 Then
                'ID3v1.1 tag: byte 28 of the comment tag is zeroed, byte 29 is the track number.
                outInfo.ID3v1Version = 1
            Else
                'ID3v1.0 tag: contains no track number.
                outInfo.ID3v1Version = 0
            End If
        End If
    End If

    'ID3 v2 tag.
    outInfo.ID3v2Version = -1
    SetFilePointer fh, 0, 0, FILE_BEGIN
    ReadFile fh, TagHeader, Len(TagHeader), 0, ByVal 0
    If Data2String(VarPtr(TagHeader.Identifier(0)), 3, ENC_ISO) = "ID3" Then
        outInfo.ID3v2Version = TagHeader.Version(0)
    End If

    'File size.
    outInfo.FileSize = fsize

    'calcular duracion
    If outInfo.bitrate <> 0 Then lSeconds = Int((fsize * 8) / outInfo.bitrate / 1000)

    HH = lSeconds \ 3600      ' Horas
    MM = lSeconds \ 60 Mod 60    ' Minutos
    ss = lSeconds Mod 60      ' Segundos

    If HH > 0 Then sTempTime = HH & ":"
    outInfo.LengthTime = Trim(sTempTime & Format$(MM, "00") & ":" & Format$(ss, "00"))
    Debug.Print outInfo.LengthTime
    'Tamaño del archivo MB
    dFilesize = fsize / 1024
    outInfo.SizeMB = CLng(dFilesize / 1024 * 100) / 100 & " MB"

    CloseHandle fh
    Exit Function
errorHandler:
    Debug.Print err.Description
    CloseHandle fh

End Function




'**********************************Other "shared use" functions****************************************

Private Function GetTagSize(ByVal fh As Long, ByVal fp As Long) As Long
    Dim TagHeader As v2TagHeader

    'Search for an ID3v2 tag.
    SetFilePointer fh, fp, 0, FILE_BEGIN
    ReadFile fh, TagHeader, Len(TagHeader), 0, ByVal 0
    If Data2String(VarPtr(TagHeader.Identifier(0)), 3, ENC_ISO) = "ID3" Then
        'The size stored in the header excludes itself, and excludes the footer (if present).
        GetTagSize = Data2Long(VarPtr(TagHeader.Size(0)), True) + Len(TagHeader)

        'v 2.4 (or later?) flags: %abcd0000 abc = ignored, d = footer present
        If TagHeader.Version(0) >= 4 Then
            If TagHeader.Flags And &H10& Then
                'Add the size of the footer (which is the same size than the header) to the existing size.
                GetTagSize = GetTagSize + Len(TagHeader)
            End If
        End If
    End If
End Function

Private Sub IgnoreReadOnly(ByVal FileName As String)
    Dim rdAttr As Long

    rdAttr = GetFileAttributes(FileName)
    If rdAttr = -1 Then Exit Sub
    If rdAttr And FILE_ATTRIBUTE_READONLY Then
        'Set this bit to 0. Use xor for that.
        rdAttr = rdAttr Xor FILE_ATTRIBUTE_READONLY
        SetFileAttributes FileName, rdAttr
    End If
End Sub

Private Sub ResetTag(ByRef Tag As ID3Tag)
    With Tag
        .Album = ""
        .Artist = ""
        .Comment = ""
        .Composer = ""
        .Copyright = ""
        .EncodedBy = ""
        .Genre = ""
        .LinkTo = ""
        .OrigArtist = ""
        .SongYear = ""
        .Title = ""
        .TrackNr = ""
    End With
End Sub

Private Function FileExists(ByVal sFilename$) As Boolean
'checks if a file or dir exists
'returns true if it does, returns false otherwise
    Dim hFile&, Win32FindData As WIN32_FIND_DATA
    sFilename = Trim$(sFilename)
    hFile = FindFirstFile(sFilename, Win32FindData)

    If (hFile <> INVALID_HANDLE_VALUE) And (hFile <> ERROR_NO_MORE_FILES) Then
        FileExists = True
    ElseIf GetFileAttributes(sFilename) <> (-1) Then
        ' FindFirstFile will not return the root dor of a drive so we check the attributes
        ' of sFileName in case it is the root
        FileExists = True
    End If

    Call FindClose(hFile)

End Function
