VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PlainBars"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private Const ANGLENUMERATOR = 6.283185      ' 2 * Pi
Private Const NUMSAMPLES = 1024              ' Number of Samples
Private Const NUMBITS = 10                   ' Number of Bits

Private Col As Long

Private ReversedBits(0 To NUMSAMPLES - 1) As Long ' Bit reservation

Private Declare Function Rectangle Lib "gdi32" (ByVal hdc As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function FillRect Lib "user32" (ByVal hdc As Long, lpRect As RECT, ByVal hBrush As Long) As Long

Private Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function CreatePen Lib "gdi32" (ByVal nPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long


Public Sub DoReverse()

    Dim i As Long
    For i = LBound(ReversedBits) To UBound(ReversedBits)
        ReversedBits(i) = ReverseBits(i, NUMBITS)
    Next
End Sub
Public Function ReverseBits(ByVal Index As Long, NUMBITS As Byte) As Long
    Dim i As Byte, Rev As Long
    For i = 0 To NUMBITS - 1
        Rev = (Rev * 2) Or (Index And 1)
        Index = Index \ 2
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
    
End Sub

' my own code

'equalize: reduces impact of bassy frequencies
'and slightly amplifies higher frequencies

Public Sub Equalize(InData() As Single, Max As Single)

 'On Error Resume Next
  Dim i As Integer, Temp As Single
  

   For i = 0 To UBound(InData)
   'scaling = -0.01
    'Temp = -0.01 * Math.Log((UBound(InData) / 2) - i) ' bassy frequcies are found in the LBound of the array
    InData(i) = InData(i) * (Max / (((UBound(InData) / 2.1) - i)) * 0.25)
    'inData(i) = InData(i) * Temp
    'InData(i) =
   Next
 
End Sub
Public Sub drawVis(hdc As Long, DrawData() As Integer, Height As Single, Width As Single)

On Error GoTo e
Static PeakData(0 To 1023) As Single
Static ExData(0 To 1023) As Single
Static OutData(0 To 1023) As Single

Static x As Integer, yy As Integer, i As Single, j As Single
Static Temp As Single
Static offSet As Integer, pPen As Long, r As RECT
offSet = Abs(Width - 255) / 2

FFTAudio DrawData, OutData
Equalize OutData, Height

For x = 8 To UBound(PeakData) / 4 - 4 Step 8

If (PeakData(x * 4) < Abs(OutData(x * 4))) Then PeakData(x * 4) = OutData(x * 4) 'peak of outdata
If (ExData(x * 4) < Abs(PeakData(x * 4))) Then ExData(x * 4) = PeakData(x * 4)  'peak of peak data

yy = (Sqr(Abs(PeakData(x * 4)) / 5)) Mod Height

 Col = RGB(0, 128, 64)
pPen = CreatePen(0, 1, Col)
SelectObject hdc, pPen

r.Top = Height - yy
r.Bottom = Height
r.Left = x + offSet
r.Right = x + 6 + offSet

FillRect hdc, r, pPen

SelectObject hdc, -1
DeleteObject pPen

Col = RGB(0, yy * 2, yy * 8)
yy = (Sqr(Abs(ExData(x * 4)) / 5)) Mod Height
pPen = CreatePen(0, 1, Col)
SelectObject hdc, pPen

r.Top = Height - yy
r.Bottom = Height - yy - 2
r.Left = x + offSet
r.Right = x + 6 + offSet

FillRect hdc, r, pPen
SelectObject hdc, -1
DeleteObject pPen




PeakData(x * 4) = PeakData(x * 4) - PeakData(x * 4) * 0.5
ExData(x * 4) = ExData(x * 4) - ExData(x * 4) * 0.1
Next x


Exit Sub
e:

End Sub

Private Sub Class_Initialize()
DoReverse

End Sub



