VERSION 5.00
Begin VB.UserControl ScrollText 
   AutoRedraw      =   -1  'True
   ClientHeight    =   2205
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3360
   ScaleHeight     =   147
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   224
   ToolboxBitmap   =   "ScrollText.ctx":0000
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   150
      Left            =   1065
      Top             =   1740
   End
   Begin VB.PictureBox picCaptionText 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFFF&
      Height          =   90
      Left            =   0
      MouseIcon       =   "ScrollText.ctx":0312
      ScaleHeight     =   6
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   157
      TabIndex        =   3
      Top             =   0
      Width           =   2355
   End
   Begin VB.PictureBox picTextScroll 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   90
      Left            =   360
      ScaleHeight     =   6
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   18
      TabIndex        =   2
      Top             =   1155
      Visible         =   0   'False
      Width           =   270
   End
   Begin VB.PictureBox picDefault 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   270
      Left            =   375
      ScaleHeight     =   18
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   155
      TabIndex        =   1
      Top             =   1395
      Visible         =   0   'False
      Width           =   2325
   End
   Begin VB.PictureBox picText 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   270
      Left            =   420
      ScaleHeight     =   18
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   155
      TabIndex        =   0
      Top             =   720
      Visible         =   0   'False
      Width           =   2325
   End
End
Attribute VB_Name = "ScrollText"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'ORIGINALLY DESIGNED BY: raul mortinez
'Updation by: Mahesh Kurmi
'Many bugs fixed
'Drag function introduced
'proper alignemnt of text etc

Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private iSpeedScroll As Integer    'speed of scrolling
Private iorgXScroll As Integer, iDesHeight As Integer, iDesWidth As Integer
Attribute iDesHeight.VB_VarUserMemId = 1073938433
Attribute iDesWidth.VB_VarUserMemId = 1073938433
Private bZigZagScroll As Boolean    'to and fro scroll
Attribute bZigZagScroll.VB_VarUserMemId = 1073938436
Private bscroll_moving_left As Boolean
Attribute bscroll_moving_left.VB_VarUserMemId = 1073938437
Private sScrollText As String
Attribute sScrollText.VB_VarUserMemId = 1073938438
Private bScrolling As Boolean    ' currently scrolling or not
Attribute bScrolling.VB_VarUserMemId = 1073938439
Private bScroll As Boolean    'should it scroll or not
Attribute bScroll.VB_VarUserMemId = 1073938440
Private bAutosize As Boolean
Attribute bAutosize.VB_VarUserMemId = 1073938441

Public Event Click()
Public Event DBLClick()
Public Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

Private eAlignText As AlignmentConstants
Attribute eAlignText.VB_VarUserMemId = 1073938442
Private bMoveL, bMoveR As Boolean
Attribute bMoveL.VB_VarUserMemId = 1073938443
Attribute bMoveR.VB_VarUserMemId = 1073938443
Public Enum peScrollType
    Rolling = 0
    ZigZag = 1
End Enum

Dim prevX As Integer
Attribute prevX.VB_VarUserMemId = 1073938445
Private peST As peScrollType
Attribute peST.VB_VarUserMemId = 1073938446

'//--------------------------------------------------------------------------
Public Property Get hwnd() As Variant
    hwnd = UserControl.hwnd
End Property


'//--------------------------------------------------------------------------
Public Property Get AutoSize() As Boolean
    AutoSize = bAutosize
End Property

Public Property Let AutoSize(ByVal bValue As Boolean)
    bAutosize = bValue
    If bAutosize = True Then
        UserControl.Width = picTextScroll.ScaleWidth * 15
    End If
    PropertyChanged "AutoSize"
End Property

'//--------------------------------------------------------------------------
Public Property Get Scroll() As Boolean
    Scroll = bScroll
End Property

Public Property Let Scroll(ByVal bValue As Boolean)
    bScroll = bValue
    Timer1.Enabled = False
    If bScroll = True Then Call ScrollText
    PropertyChanged "Scroll"
End Property

'//--------------------------------------------------------------------------
Public Property Get PictureText() As Picture
    Set PictureText = picText.Picture
End Property

Public Property Set PictureText(ByVal New_Picture As Picture)
    Set picText.Picture = New_Picture
    picText.AutoSize = True
    picCaptionText.Height = picText.ScaleHeight / 3
    picCaptionText.Width = UserControl.ScaleWidth
    UserControl.Height = (picText.ScaleHeight / 3) * 15

    BuildText sScrollText
    PropertyChanged "PictureText"
End Property
'//--------------------------------------------------------------------------
Public Property Get picHandle() As Long
'returns handle of main picture for tooltip etc
    picHandle = picCaptionText.hwnd
End Property


'//--------------------------------------------------------------------------
Public Property Get AlignText() As AlignmentConstants
    AlignText = eAlignText
End Property

Public Property Let AlignText(ByVal vNewValue As AlignmentConstants)
    eAlignText = vNewValue
    UpdateAlign
    PropertyChanged "AlignText"
End Property

'//--------------------------------------------------------------------------
Public Property Get CaptionText() As String
    CaptionText = sScrollText
End Property

Public Property Let CaptionText(ByVal vNewValue As String)
    sScrollText = vNewValue
    If sScrollText = "" Then sScrollText = " "
    BuildText sScrollText
    If bScroll = True Then Call ScrollText
    PropertyChanged "CaptionText"
End Property

'//--------------------------------------------------------------------------
Public Property Get ScrollType() As peScrollType
    ScrollType = peST
End Property

Public Property Let ScrollType(ByVal vNewValue As peScrollType)
    peST = vNewValue
    bZigZagScroll = peST
    If bScrolling = True Then Call ScrollText
    PropertyChanged "ScrollType"
End Property

'//--------------------------------------------------------------------------
Public Property Get ScrollVelocity() As Integer
    ScrollVelocity = Timer1.Interval
End Property

Public Property Let ScrollVelocity(ByVal vNewValue As Integer)
    Timer1.Interval = vNewValue
    PropertyChanged "ScrollVelocity"
End Property

'//--------------------------------------------------------------------------
Public Property Get ScrollingNow() As Boolean
    ScrollingNow = Timer1.Enabled
End Property

'//--------------------------------------------------------------------------
Public Sub StopScroll(ByVal bValue As Boolean)
    If bScrolling = True Then
        Timer1.Enabled = Not bValue
    End If
End Sub

'//--------------------------------------------------------------------------
Public Property Get BackColor() As ole_color
    BackColor = picCaptionText.BackColor
End Property

Public Property Let BackColor(ByVal vNewValue As ole_color)
    picCaptionText.BackColor = vNewValue
    picTextScroll.BackColor = picCaptionText.BackColor
    BuildText sScrollText
    PropertyChanged "BackColor"
End Property




Private Sub picCaptionText_Click()
    RaiseEvent Click
End Sub

Private Sub picCaptionText_DblClick()
    RaiseEvent DBLClick
End Sub

Private Sub picCaptionText_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseDown(Button, Shift, X, Y)
    prevX = X
    Timer1.Enabled = False
    'bMoveL = True
    ' iorgXScroll = picCaptionText.ScaleWidth
End Sub

Private Sub picCaptionText_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseMove(Button, Shift, X, Y)

    If bScrolling = False Then Exit Sub
    If Button = 1 Then
        If (prevX - X) >= 1 And ((iorgXScroll < picTextScroll.ScaleWidth - 2) Or ScrollType = Rolling) Then    ' move scroll leftwards
            BitBlt picCaptionText.hDC, 0, 0, iDesWidth, iDesHeight, picCaptionText.hDC, 2, 0, &HCC0020
            '//copy first pixel col of picCaptionText taken from pictextscroll and put it to end of it
            BitBlt picCaptionText.hDC, iDesWidth - 1, 0, 2, iDesHeight, picTextScroll.hDC, iorgXScroll, 0, &HCC0020

            'increment variable for further scroll
            If ScrollType = Rolling And iorgXScroll >= picTextScroll.ScaleWidth Then iorgXScroll = 0    'whole text is scrolled once

            iorgXScroll = iorgXScroll + 2
            'whole text is scrolled once
            prevX = X
        ElseIf ((X - prevX) >= 1) And ((iorgXScroll - picCaptionText.ScaleWidth > 1) Or ScrollType = Rolling) Then
            iorgXScroll = iorgXScroll - 2

            BitBlt picCaptionText.hDC, 2, 0, iDesWidth, iDesHeight, picCaptionText.hDC, 0, 0, &HCC0020

            If ScrollType = Rolling Then
                If iorgXScroll <= 0 Then iorgXScroll = picTextScroll.ScaleWidth
                If iorgXScroll <= picCaptionText.ScaleWidth Then
                    BitBlt picCaptionText.hDC, 0, 0, 2, iDesHeight, picTextScroll.hDC, picTextScroll.ScaleWidth - Abs(iorgXScroll - picCaptionText.ScaleWidth), 0, &HCC0020
                Else
                    BitBlt picCaptionText.hDC, 0, 0, 2, iDesHeight, picTextScroll.hDC, iorgXScroll - picCaptionText.ScaleWidth, 0, &HCC0020
                End If
            Else
                '//copy first pixel col of picCaptionText taken from pictextscroll and put it to end of it
                BitBlt picCaptionText.hDC, 0, 0, 2, iDesHeight, picTextScroll.hDC, iorgXScroll - picCaptionText.ScaleWidth, 0, &HCC0020
            End If
            'increment variable for further scroll

            prevX = X
        End If
    End If

    picCaptionText.Refresh
    DoEvents
End Sub

Private Sub picCaptionText_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseUp(Button, Shift, X, Y)
    If bScrolling Then Timer1.Enabled = True
End Sub

Private Sub Timer1_Timer()
    Dim i As Integer

    Static bscroll_wait_collide As Boolean
    Static stcPause As Integer    ' stcPause is used to pause scroll on collision in zigzag movement
    'Exit Sub
    If bZigZagScroll = False Then    'i.e. scroll is to be rotated

        For i = 0 To iSpeedScroll
            '// shift picture of picCaptionText to one pixel column left
            BitBlt picCaptionText.hDC, 0, 0, iDesWidth, iDesHeight, picCaptionText.hDC, 1, 0, &HCC0020
            '//copy first pixel col of picCaptionText taken from pictextscroll and put it to end of it
            BitBlt picCaptionText.hDC, iDesWidth, 0, 1, iDesHeight, picTextScroll.hDC, iorgXScroll, 0, &HCC0020



            'increment variable for further scroll


            iorgXScroll = iorgXScroll + 1
            If iorgXScroll = picTextScroll.ScaleWidth Then iorgXScroll = 0    'whole text is scrolled once
        Next i

        picCaptionText.Refresh
        Exit Sub
    End If

    If bscroll_wait_collide = True And stcPause < 15 Then    'waits for a moment if scroll has to move to and fro and
        'srolls head either reaches extreme left if going left or its tail reaches extreme right if going right
        'i.e. at the moment scrll reflects back there is a wait
        stcPause = stcPause + 1
        Exit Sub
    End If

    'if neither scroll is in pause condition nor is rolling i.e. zigzag motion is occurring
    For i = 0 To iSpeedScroll

        If bscroll_moving_left = False Then    ' move scroll leftwards

            BitBlt picCaptionText.hDC, 0, 0, iDesWidth, iDesHeight, picCaptionText.hDC, 1, 0, &HCC0020
            '//copy first pixel col of picCaptionText taken from pictextscroll and put it to end of it
            BitBlt picCaptionText.hDC, iDesWidth, 0, 1, iDesHeight, picTextScroll.hDC, iorgXScroll, 0, &HCC0020

            'increment variable for further scroll

            iorgXScroll = iorgXScroll + 1
            If iorgXScroll > picTextScroll.ScaleWidth Then    ' scroll has reached max left limit
                bscroll_moving_left = True
                'iorgXScroll = Abs(picTextScroll.ScaleWidth - picCaptionText.ScaleWidth)
                bscroll_wait_collide = True    ' reflect scroll back towards right so send control to else block
                stcPause = 0
            End If
        Else

            BitBlt picCaptionText.hDC, 1, 0, iDesWidth, iDesHeight, picCaptionText.hDC, 0, 0, &HCC0020
            '//copy first pixel col of picCaptionText taken from pictextscroll and put it to end of it
            BitBlt picCaptionText.hDC, 0, 0, 1, iDesHeight, picTextScroll.hDC, iorgXScroll - picCaptionText.ScaleWidth, 0, &HCC0020

            'increment variable for further scroll

            iorgXScroll = iorgXScroll - 1
            'If iorgXScroll = picTextScroll.ScaleWidth Then iorgXScroll = 0 'whole text is scrolled once


            ' BitBlt picCaptionText.hDc, 1, 0, iDesWidth, iDesHeight, picCaptionText.hDc, 0, 0, &HCC0020
            'BitBlt picCaptionText.hDc, 0, 0, 1, iDesHeight, picTextScroll.hDc, iorgXScroll, 0, &HCC0020


            If iorgXScroll - picCaptionText.ScaleWidth < 0 Then
                bscroll_moving_left = False
                'iorgXScroll = picCaptionText.ScaleWidth
                bscroll_wait_collide = True
                stcPause = 0
            End If
        End If

    Next i

    picCaptionText.Refresh
End Sub



Private Sub UserControl_Initialize()
'pictext is picture box to contain font_ficture as specified in skin
    picText.Picture = picDefault.Picture    'pictext has default font picture
    sScrollText = "Mahesh's Scroll Text Control"
    BuildText sScrollText

End Sub

Private Sub UserControl_Resize()
    picCaptionText.Width = UserControl.ScaleWidth
    UserControl.Height = (picText.ScaleHeight / 3) * 15
    If picTextScroll.ScaleWidth > picCaptionText.ScaleWidth Then bScrolling = True
    UpdateAlign
    If bScrolling = True Then
        ' Call ScrollText
    Else
        'picCaptionText.Height = UserControl.ScaleHeight
        'UpdateAlign
    End If
End Sub

'//--------------------------------------------------------------------------
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("PictureText", picText.Picture, picDefault.Picture)
    Call PropBag.WriteProperty("BackColor", picCaptionText.BackColor, &H0)
    Call PropBag.WriteProperty("CaptionText", sScrollText, "Scroll Text")
    Call PropBag.WriteProperty("AlignText", eAlignText, 0)
    Call PropBag.WriteProperty("ScrollType", peST, 0)
    Call PropBag.WriteProperty("ScrollVelocity", Timer1.Interval, 200)
    Call PropBag.WriteProperty("Scroll", bScroll, False)
    Call PropBag.WriteProperty("AutoSize", bAutosize, False)

End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    On Error Resume Next
    picText.Picture = PropBag.ReadProperty("PictureText", picDefault.Picture)
    If picText.Picture = 0 Then picText.Picture = picDefault.Picture
    picCaptionText.BackColor = PropBag.ReadProperty("BackColor", &H0)
    picTextScroll.BackColor = picCaptionText.BackColor
    eAlignText = PropBag.ReadProperty("AlignText", 0)
    Timer1.Interval = PropBag.ReadProperty("ScrollVelocity", 200)
    sScrollText = PropBag.ReadProperty("CaptionText", "Scroll Text")
    bScroll = PropBag.ReadProperty("Scroll", False)
    bAutosize = PropBag.ReadProperty("AutoSize", False)
    peST = PropBag.ReadProperty("ScrollType", 0)
    bZigZagScroll = peST
    BuildText sScrollText
    'If bScroll = True Then Call ScrollText
End Sub

Private Sub UpdateAlign()
'picCaptionText.Picture = LoadPicture()
    Dim unitwidth, unitheight, k As Integer
    unitwidth = picText.ScaleWidth / 31
    unitheight = picText.ScaleHeight / 3

    'picCaptionText.Picture
    'BitBlt picCaptionText.hdc, 0, 0, picText.ScaleWidth / 2, picTextScroll.ScaleHeight, picTextScroll.hdc, 0, 0, &HCC0020

    Select Case eAlignText

    Case 0    '// left
        BitBlt picCaptionText.hDC, 0, 0, picTextScroll.ScaleWidth, picTextScroll.ScaleHeight, picTextScroll.hDC, 0, 0, &HCC0020
        k = picTextScroll.ScaleWidth

        While (k <= (picCaptionText.ScaleWidth))  'fill remaining portion with spaces
            Call BitBlt(picCaptionText.hDC, k, 0, unitwidth, unitheight, picText.hDC, unitwidth * 30, 0, vbSrcCopy)
            k = k + unitwidth
        Wend

    Case 2    '// center
        BitBlt picCaptionText.hDC, (picCaptionText.ScaleWidth / 2) - (picTextScroll.ScaleWidth / 2), 0, picTextScroll.ScaleWidth, picTextScroll.ScaleHeight, picTextScroll.hDC, 0, 0, &HCC0020
        k = (picCaptionText.ScaleWidth - picTextScroll.ScaleWidth) / 2

        While (k >= 0)    'fill vacant left portion with spaces
            Call BitBlt(picCaptionText.hDC, k - unitwidth, 0, unitwidth, unitheight, picText.hDC, unitwidth * 30, 0, vbSrcCopy)
            k = k - unitwidth    ' decreamnent for right alignment
        Wend    'icellX gives pos of cell to be filled by pic
        k = (picCaptionText.ScaleWidth + picTextScroll.ScaleWidth) / 2 - 1
        While (k <= (picCaptionText.ScaleWidth))   'fill right vacant portion with spaces
            Call BitBlt(picCaptionText.hDC, k, 0, unitwidth, unitheight, picText.hDC, unitwidth * 30, 0, vbSrcCopy)
            k = k + unitwidth
        Wend

    Case 1    '//right
        BitBlt picCaptionText.hDC, (picCaptionText.ScaleWidth) - (picTextScroll.ScaleWidth), 0, picTextScroll.ScaleWidth, picTextScroll.ScaleHeight, picTextScroll.hDC, 0, 0, &HCC0020
        k = (picCaptionText.ScaleWidth) - (picTextScroll.ScaleWidth)

        While (k >= 0)    'fill remaining portion at right with spaces
            Call BitBlt(picCaptionText.hDC, k - unitwidth, 0, unitwidth, unitheight, picText.hDC, unitwidth * 30, 0, vbSrcCopy)
            k = k - unitwidth
        Wend

    Case Else
        BitBlt picCaptionText.hDC, 0, 0, picTextScroll.ScaleWidth, picTextScroll.ScaleHeight, picTextScroll.hDC, 0, 0, &HCC0020
    End Select
    ' picCaptionText.Picture = picCaptionText.Image

End Sub

Private Sub ScrollText()
    Timer1.Enabled = False
    If sScrollText = " " Then Exit Sub
    If picTextScroll.ScaleWidth > picCaptionText.ScaleWidth Then

        '// cool effect
        If peST = Rolling Then BuildText "* " & sScrollText & " *"

        picCaptionText.Picture = LoadPicture()
        picCaptionText.Picture = picTextScroll.Picture
        iDesHeight = picCaptionText.ScaleHeight
        iDesWidth = picCaptionText.ScaleWidth - 1
        iorgXScroll = picCaptionText.ScaleWidth
        iSpeedScroll = 1
        bscroll_moving_left = False
        bScrolling = True
        bMoveR = True
        Timer1.Enabled = True
    Else
        bScrolling = False
    End If

End Sub

Private Sub BuildText(sText As String)
    Dim i As Integer, iCell As Integer, iCellX As Integer
    Dim s As String
    Dim unitwidth, unitheight As Integer
    picCaptionText.Picture = LoadPicture()
    picCaptionText.Width = UserControl.ScaleWidth
    'picTextScroll.Width = UserControl.ScaleWidth
    'colums in each row one for every letter in picture
    'so its width/31 gives with of one letter
    'len()gives no of letters in string passed
    picTextScroll.Height = (picText.ScaleHeight / 3) * 15  'picture ib skin has 3 rows of letters ,so its width/3 gives height of letter
    picCaptionText.Height = picTextScroll.Height
    picTextScroll.Width = (Len(sText)) * (picText.ScaleWidth / 31)    'pictext has picture as in skin having 31
    picTextScroll.Picture = LoadPicture()

    If bAutosize = True Then UserControl.Width = picTextScroll.ScaleWidth * 15
    unitwidth = picText.ScaleWidth / 31
    unitheight = picText.ScaleHeight / 3

    'picCaptionText.Picture = picTextScroll.Image 'pic stored in its image(CANVAS) is copied back in picture since images acts as canvas for BitBlt operation in GDI32
    ' picTextScroll.Picture = LoadPicture()

    For i = 1 To Len(sText)
        s = Mid(sText, i, 1)
        iCell = IndexWord(UCase(s))    ' icell and icellx are int to show pos of cell to be filled
        iCellX = (picText.ScaleWidth / 31) * (i - 1)    'pictext has 31 colums in each of 3 rows
        'icellX gives pos of cell to be filled by pic
        CopyCell iCell, iCellX
    Next i
    picTextScroll.Picture = picTextScroll.Image    'pic stored in its image(CANVAS) is copied back in picture since images acts as canvas for BitBlt operation in GDI32
    UpdateAlign
End Sub

Private Function IndexWord(sWord As String) As Integer
    Dim iWord As Integer
    Select Case sWord
    Case "A", "Á", "À": iWord = 0
    Case "B": iWord = 1
    Case "C": iWord = 2
    Case "D": iWord = 3
    Case "E", "É": iWord = 4
    Case "F": iWord = 5
    Case "G": iWord = 6
    Case "H": iWord = 7
    Case "I", "Í", "Ì": iWord = 8
    Case "J": iWord = 9
    Case "K": iWord = 10
    Case "L": iWord = 11
    Case "M": iWord = 12
    Case "N", "Ñ": iWord = 13
    Case "O", "Ó", "Ò": iWord = 14
    Case "P": iWord = 15
    Case "Q": iWord = 16
    Case "R": iWord = 17
    Case "S": iWord = 18
    Case "T": iWord = 19
    Case "U", "Ú", "Ù", "Ü", "Û": iWord = 20
    Case "V": iWord = 21
    Case "W": iWord = 22
    Case "X": iWord = 23
    Case "Y": iWord = 24
    Case "Z": iWord = 25
    Case """": iWord = 26
    Case "@", "®": iWord = 27
    Case " ": iWord = 29
    Case " ": iWord = 29
    Case " ": iWord = 29

    Case "0": iWord = 31
    Case "1": iWord = 32
    Case "2": iWord = 33
    Case "3": iWord = 34
    Case "4": iWord = 35
    Case "5": iWord = 36
    Case "6": iWord = 37
    Case "7": iWord = 38
    Case "8": iWord = 39
    Case "9": iWord = 40
    Case "_": iWord = 41
    Case ".": iWord = 42
    Case ":", ";": iWord = 43
    Case "(", "<": iWord = 44
    Case ")", ">": iWord = 45
    Case "-", "~", "°": iWord = 46
    Case "'", "`", "´": iWord = 47
    Case "!", "¡": iWord = 48
    Case "_": iWord = 49
    Case "+": iWord = 50
    Case "\", "|": iWord = 51
    Case "/": iWord = 52
    Case "[", "{": iWord = 53
    Case "]", "}": iWord = 54
    Case "^": iWord = 55
    Case "&": iWord = 56
    Case "%": iWord = 57
    Case ",": iWord = 58
    Case "=": iWord = 59
    Case "$": iWord = 60
    Case "#": iWord = 61

    Case "Ã": iWord = 62
    Case "ö", "õ", "ô": iWord = 63
    Case "Ä": iWord = 64
    Case "?", "¿": iWord = 65
    Case "*": iWord = 66

    Case Else
        iWord = 29

    End Select
    IndexWord = iWord
End Function

Private Sub CopyCell(iIndex As Integer, orgX As Integer)
'orgX gives starting pos of cell to be filled in pictextscroll
    Dim iorgXScroll As Integer
    Dim srcY As Integer
    Dim srcWidth As Integer
    Dim srcHeight As Integer

    srcWidth = picText.ScaleWidth / 30    'width of column in skin picture of font
    srcHeight = picText.ScaleHeight / 3    'height of row in skin picture
    'col no. and row no can be retrieved knowing the pos or index of letter
    If iIndex <= 30 Then srcY = srcHeight * 0    'picture of letter in skin pic lies in first row
    If iIndex > 30 And iIndex <= 61 Then srcY = srcHeight * 1: iIndex = iIndex - 31
    If iIndex > 61 Then srcY = srcHeight * 2: iIndex = iIndex - 62

    iorgXScroll = srcWidth * iIndex
    'pictext is picture box to store skin pic permanently
    'using column no. and row no. we can copy pic of particular letter from pictext to the pcscrolltext which is shown in mediaplayer
    Call BitBlt(picTextScroll.hDC, orgX, 0, srcWidth, srcHeight, picText.hDC, iorgXScroll, srcY, vbSrcCopy)
    'pictextscroll has this picture in form of its image property (CANVAS for picture box)
End Sub



