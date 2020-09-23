VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Bars2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private Const ANGLENUMERATOR = 6.283185      ' 2 * Pi
Private Const NUMSAMPLES = 1024              ' Number of Samples
Private Const NumBits = 10                   ' Number of Bits
Private Col As Long
Private ReversedBits(0 To NUMSAMPLES - 1) As Long ' Bit reservation

Private Declare Function Rectangle Lib "gdi32" (ByVal hDC As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function FillRect Lib "user32" (ByVal hDC As Long, lpRect As RECT, ByVal hBrush As Long) As Long
Private Declare Function LineTo Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function MoveToEx Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, lpPoint As POINTAPI) As Long

Private Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type
Private NewX As Single, NewY As Single
Private Type POINTAPI
        X As Long
        Y As Long
End Type
Private Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long
Private Declare Function CreatePen Lib "gdi32" (ByVal nPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long


Public Sub DoReverse()

    Dim i As Long
    For i = LBound(ReversedBits) To UBound(ReversedBits)
        ReversedBits(i) = ReverseBits(i, NumBits)
    Next
End Sub
Public Function ReverseBits(ByVal index As Long, NumBits As Byte) As Long
    Dim i As Byte, Rev As Long
    For i = 0 To NumBits - 1
        Rev = (Rev * 2) Or (index And 1)
        index = index \ 2
    Next
    ReverseBits = Rev
End Function

' Fast Fourier Tansform: FFT
Public Sub FFTAudio(RealIn() As Integer, RealOut() As Single)

    Static ImagOut(0 To NUMSAMPLES - 1) As Single
    Static i As Long, j As Long, k As Long, N As Long, BlockSize As Long, BlockEnd As Long
    Static DeltaAngle As Single, DeltaAr As Single
    Static Alpha As Single, Beta As Single
    Static TR As Single, TI As Single, AR As Single, AI As Single
    For i = 0 To (NUMSAMPLES - 1)
        j = ReversedBits(i)
        RealOut(j) = RealIn(i)
        ImagOut(j) = 0
    Next
    BlockEnd = 1
    BlockSize = 2
    Do While BlockSize <= NUMSAMPLES
        DeltaAngle = ANGLENUMERATOR / BlockSize
        Alpha = Sin(0.5 * DeltaAngle)
        Alpha = 2! * Alpha * Alpha
        Beta = Sin(DeltaAngle)
        i = 0
        Do While i < NUMSAMPLES
            AR = 1!
            AI = 0!
            j = i
            For N = 0 To BlockEnd - 1
                k = j + BlockEnd
                TR = AR * RealOut(k) - AI * ImagOut(k)
                TI = AI * RealOut(k) + AR * ImagOut(k)
                RealOut(k) = RealOut(j) - TR
                ImagOut(k) = ImagOut(j) - TI
                RealOut(j) = RealOut(j) + TR
                ImagOut(j) = ImagOut(j) + TI
                DeltaAr = Alpha * AR + Beta * AI
                AI = AI - (Alpha * AI - Beta * AR)
                AR = AR - DeltaAr
                j = j + 1
            Next N
            i = i + BlockSize
        Loop
        BlockEnd = BlockSize
        BlockSize = BlockSize * 2
    Loop
    
 Equalize RealOut

End Sub

' my own code

'equalize: reduces impact of bassy frequencies
'and slightly amplifies higher frequencies

Public Sub Equalize(InData() As Single)

 On Error Resume Next
  Dim i As Integer, Temp As Single
  

   For i = 0 To UBound(InData)
   'scaling = -0.01
    InData(i) = InData(i) * (50 / (((UBound(InData) / 2.1) - i)))
    Next
 
End Sub
Public Sub drawVis(hDC As Long, DrawData() As Integer, Height As Single, Width As Single)
Static X As Integer, Y As Single
Dim dummyPt As POINTAPI
Dim ppen As Long, ave As Single
Dim Space As Integer
Dim Barwidth As Integer
Dim Yspace As Integer
Dim Ywidth As Integer
Dim nSample As Integer
Space = 1
Barwidth = 3
Ywidth = 2
Yspace = 1
Dim i As Integer
Dim Bandwidth As Integer
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
Y = Sqr(Abs(PeakDataExTop(X)))

NewX = Barwidth * X + Space * X
NewY = (Height - Y) / 2

ppen = CreatePen(0, 1, RGB(240, 240, 255))
SelectObject hDC, ppen
 MoveToEx hDC, NewX, NewY, dummyPt
LineTo hDC, NewX + Barwidth, NewY
SelectObject hDC, -1
DeleteObject ppen


'Col = RGB(Abs(r1 * 0.25), Abs(g1 * 0.25), Abs(b1 * 0.25))
NewY = (Height + Y) / 2

ppen = CreatePen(0, 1, RGB(80, 90, 100))
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
  
NewX = Barwidth * X + Space * X
NewY = (Height - Y) / 2

'//////////////////////////////
ppen = CreatePen(0, 1, RGB(123, 169, 220))
SelectObject hDC, ppen

For i = 0 To Abs(NewY - Height / 2) / (Ywidth + Yspace)
 'PicSpectrum.ForeColor = GetGradColor(maxx(1, (Abs(NewY - Height / 2) / 5)), i, vbRed, vbGreen, vbBlue)
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
Col = RGB(27, 37, 72)

ppen = CreatePen(0, 1, Col)
SelectObject hDC, ppen
For i = 0 To Abs(NewY - Height / 2) / (Ywidth + Yspace)
 'PicSpectrum.ForeColor = GetGradColor(maxx(1, (Abs(NewY - Height / 2) / 5)), i, vbRed, vbGreen, vbBlue)
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
Private Sub Class_Initialize()
DoReverse

End Sub




