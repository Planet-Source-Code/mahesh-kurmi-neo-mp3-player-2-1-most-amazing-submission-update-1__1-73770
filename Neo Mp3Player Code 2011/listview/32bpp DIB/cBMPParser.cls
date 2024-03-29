VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cBMPParser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' No APIs are declared public. This is to prevent possibly, differently
' declared APIs, or different versions of the same API, from conflciting
' with any APIs you declared in your project. Same rule for UDTs.

Private Type SafeArrayBound
    cElements As Long
    lLbound As Long
End Type
Private Type SafeArray
    cDims As Integer
    fFeatures As Integer
    cbElements As Long
    cLocks As Long
    pvData As Long
    rgSABound(0 To 1) As SafeArrayBound    ' reusable UDT for 1 & 2 dim arrays
End Type
Private Type BITMAP
    bmType As Long
    bmWidth As Long
    bmHeight As Long
    bmWidthBytes As Long
    bmPlanes As Integer
    bmBitsPixel As Integer
    bmBits As Long
End Type

' used to transfer a stdPicture bmp,jpg,wmf to a DIB
Private Declare Function GetGDIObject Lib "gdi32.dll" Alias "GetObjectA" (ByVal hObject As Long, ByVal nCount As Long, ByRef lpObject As Any) As Long
Private Declare Function VarPtrArray Lib "msvbvm60.dll" Alias "VarPtr" (Ptr() As Any) As Long
Private Declare Function GetDC Lib "user32.dll" (ByVal hwnd As Long) As Long
Private Declare Function ReleaseDC Lib "user32.dll" (ByVal hwnd As Long, ByVal hDC As Long) As Long
Private Declare Sub FillMemory Lib "kernel32.dll" Alias "RtlFillMemory" (ByRef Destination As Any, ByVal Length As Long, ByVal Fill As Byte)
Private Declare Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" (ByRef Destination As Any, ByRef Source As Any, ByVal Length As Long)

Public Function LoadStream(inStream() As Byte, cHost As c32bppDIB, _
                           Optional ByVal streamOffset As Long = 0, _
                           Optional ByVal streamLength As Long = 0) As Boolean

' PURPOSE: Test passed stream for being a 32bpp bitmap.
' If not 32bpp, then the stream is converted to a stdPicture and the contents
' of that object are drawn to the 32bpp.

' With the exception of wmf, emf & 32bpp. This class does not handle transparency.
' Therefore, the stream should have been passed to the png, gif & icon parsers
' first.

' Parameters.
' inStream() :: the byte array containing the image
' cHost :: an initialized c32bppDIB
' streamOffset :: array position for 1st byte in the stream
' streamLength :: size of stream that contains the image
'   - If zero, then size is UBound(inStream)-streamOffset+1

' IMPORTANT: the array offset & length are not checked in this class.
'   They were checked before this class was called. If this class is to
'   be pulled out and put in another project, ensure you include the
'   validation shown in c32bppDIB.LoadPicture_Stream

    Dim lValue As Long, iValue As Integer
    Dim X As Long, Y As Long, lScanWidth As Long
    Dim Offset As Long, iBitCount As Integer
    Dim aDIB() As Byte, tSA As SafeArray
    Dim bAlpha As Boolean

    ' manually parse the bitmap header.
    ' Why? because VB's LoadPicture will convert the image into a screen
    ' compatible bitmap; where if screen resolution was true color, a
    ' 32bpp image would end up being 24bit vs 32bit
    CopyMemory iValue, inStream(streamOffset), 2&   ' get 1st 2 bytes of the stream
    If iValue = &H4D42 Then                         ' is it a bmp magic number
        CopyMemory iBitCount, inStream(streamOffset + 28), 2&    ' bit count
        CopyMemory X, inStream(streamOffset + 18), 4&    ' width
        CopyMemory Y, inStream(streamOffset + 22), 4&    ' height

        ' validate size
        ' width must be at least 1 pixel & height must be a least 1 pixel
        If X < 1 Or Y = 0 Then Exit Function    ' -Y indicates top down DIB

        On Error Resume Next
        CopyMemory Offset, inStream(streamOffset + 10), 4&    ' start of image
        ' validate enough bytes exist for the image
        lValue = (streamOffset + streamLength) - (iparseByteAlignOnWord(iBitCount, X) * Abs(Y) + Offset)
        If err Then     ' should some overflow occur
            err.Clear
            lValue = -1
        End If
        If lValue >= 0 Then              ' is array big enough?
            If iBitCount = 32 Then       ' else we will allow VB to convert it for us
                ' because it doesn't contain transparency anyway
                CopyMemory lValue, inStream(streamOffset + 30), 4&    ' compression
                If lValue = 0& Then         ' manually handle no-compression bitmaps
                    ' else allow VB to convert the bitmap
                    cHost.InitializeDIB X, Abs(Y)
                    With tSA
                        .cbElements = 1
                        .cDims = 2
                        .pvData = cHost.BitsPointer
                        .rgSABound(0).cElements = cHost.Height
                        .rgSABound(1).cElements = cHost.scanWidth
                    End With
                    CopyMemory ByVal VarPtrArray(aDIB), VarPtr(tSA), 4&

                    lScanWidth = cHost.scanWidth
                    If Y < 0 Then    ' the dib is top down vs bottom up
                        ' flip the DIB
                        Y = -Y
                        For lValue = 0 To Y - 1
                            ' start of scan line in source image
                            X = lScanWidth * (Y - lValue - 1) + Offset
                            ' copy to upside down scan line on our DIB
                            CopyMemory aDIB(0, lValue), inStream(X), lScanWidth
                        Next
                    Else    ' bottom up dib; simply copy bits
                        CopyMemory ByVal cHost.BitsPointer, inStream(Offset), cHost.Height * lScanWidth
                    End If

                    ' see if 32bpp is premulitplied or not
                    iparseValidateAlphaChannel aDIB(), True, bAlpha, lValue
                    CopyMemory ByVal VarPtrArray(aDIB), 0&, 4&    ' remove overlay
                    ' set other properties
                    cHost.Alpha = bAlpha
                    cHost.ImageType = lValue
                    LoadStream = True
                End If
            End If

        End If
    End If
    On Error GoTo 0

    If cHost.Handle = 0 Then    ' we didn't process the image above, try VB's LoadPicture

        On Error Resume Next
        Dim tPic As StdPicture
        Set tPic = iparseArrayToPicture(inStream(), streamOffset, streamLength)
        If err Then
            err.Clear
        Else
            LoadStream = ConvertstdPicTo32bpp(tPic, cHost, iBitCount)
        End If

    End If

End Function

Public Function ConvertstdPicTo32bpp(stdPic As StdPicture, cHost As c32bppDIB, ByVal bitCount As Integer) As Boolean

    Dim tSA As SafeArray
    Dim cx As Long, cy As Long
    Dim tDC As Long, bAlpha As Boolean, iType As Long
    Dim aDIB() As Byte

    If stdPic Is Nothing Then Exit Function    ' couldn't convert image
    If stdPic.Type = vbPicTypeNone Then Exit Function

    ' get the picture's width & height & initialize DIB
    cx = ConvertHimetrix2Pixels(stdPic.Width, True)
    cy = ConvertHimetrix2Pixels(stdPic.Height, False)
    cHost.InitializeDIB cx, cy

    ' WMF/EMFs are kinda weird, but here is a neat trick to determine if it
    ' has transparency. Fill the entire image with white, then when it is
    ' rendered, any "transparent" areas are not drawn over, leaving the
    ' alpha byte as 255. Those areas that are drawn over are changed to zero.
    If stdPic.Type = vbPicTypeEMetafile Or stdPic.Type = vbPicTypeMetafile Then
        FillMemory ByVal cHost.BitsPointer, cy * cHost.scanWidth, 255
    End If

    ' render the stdPic to the host's dc
    tDC = cHost.LoadDIBinDC(True)
    stdPic.Render tDC + 0&, 0&, 0&, cx + 0&, cy + 0&, _
                  0, stdPic.Height, stdPic.Width, -stdPic.Height, ByVal 0&
    ' unmanage the DC if needed
    cHost.LoadDIBinDC False

    ' map our array to the host's DIB
    With tSA
        .cbElements = 1    ' as byte array
        .cDims = 2      ' as 1 dimensional
        .pvData = cHost.BitsPointer
        .rgSABound(0).cElements = cy
        .rgSABound(1).cElements = cHost.scanWidth
    End With
    CopyMemory ByVal VarPtrArray(aDIB), VarPtr(tSA), 4&    ' apply overlay

    If stdPic.Type = vbPicTypeEMetafile Or stdPic.Type = vbPicTypeMetafile Then
        ' as mentioned above, any transparent pixels will have alpha value = 255
        For cy = 0 To cHost.Height - 1
            For cx = 3 To cHost.scanWidth - 1 Step 4
                If aDIB(cx, cy) = 255 Then    ' 100% transparent
                    CopyMemory aDIB(cx - 3, cy), 0&, 4&
                    bAlpha = True
                Else                        ' 100% opaque
                    aDIB(cx, cy) = 255    ' was 255, now 0, change back to 255
                End If
            Next
        Next
        cHost.ImageType = stdPic.Type
    Else    ' jpg or non-alpha bitmap
        ' validate first that it has no alpha bytes
        If bitCount = 0 Then
            ' when called from cHost.LoadPicture_Resource then no BitCount is known
            Dim tBMP As BITMAP
            GetGDIObject stdPic.Handle, Len(tBMP), tBMP
            bitCount = tBMP.bmBitsPixel
        End If
        If bitCount = 32 Then
            iparseValidateAlphaChannel aDIB(), True, bAlpha, iType
            cHost.ImageType = iType
        Else    ' cannot have an alpha channel
            ' Note: I have experienced a 24bpp stdPicture.Rendering into a 32bpp DIB
            ' and writing in the alpha channel. The -1 below forces next routine
            ' to fill the alpha channel with 255
            cHost.ImageType = vbPicTypeBitmap
            iparseValidateAlphaChannel aDIB(), True, False, -1
        End If
    End If
    CopyMemory ByVal VarPtrArray(aDIB), 0&, 4&  ' remove overlay
    cHost.Alpha = bAlpha
    ConvertstdPicTo32bpp = True

End Function

Private Function ConvertHimetrix2Pixels(vHiMetrix As Long, Horizontally As Boolean) As Long
' conversion from Himetrics to Pixels when ScaleX/Y is not available
    If Horizontally Then
        ConvertHimetrix2Pixels = vHiMetrix * 1440 / 2540 / Screen.TwipsPerPixelX
    Else
        ConvertHimetrix2Pixels = vHiMetrix * 1440 / 2540 / Screen.TwipsPerPixelY
    End If
End Function
