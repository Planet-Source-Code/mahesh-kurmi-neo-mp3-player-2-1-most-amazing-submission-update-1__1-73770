VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Scope"
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

Private Declare Function SetPixel Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal crColor As Long) As Long
Private Declare Function MoveToEx Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, lpPoint As POINTAPI) As Long
Private Declare Function LineTo Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long) As Long

Private Declare Function CreatePen Lib "gdi32" (ByVal nPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function FillRect Lib "user32" (ByVal hdc As Long, lpRect As RECT, ByVal hBrush As Long) As Long

Private Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type

Private Type POINTAPI
        x As Long
        y As Long
End Type
Private pPen As Long, oldPen As Long


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
    Temp = -0.005 * Math.Log(i) ' bassy frequcies are found in the LBound of the array
    InData(i) = InData(i) * Temp
   Next
 
End Sub
Public Sub drawVis(hdc As Long, DrawData() As Integer, Height As Single, Width As Single)

On Error GoTo e
Static PeakData(0 To 1023) As Single
Static ExData(0 To 1023) As Single
Static OutData(0 To 1023) As Single



Static x As Integer, yy As Integer, j As Integer
Static Temp As Single, p As POINTAPI, r As RECT

FFTAudio DrawData, OutData

For x = 0 To UBound(OutData)
If (PeakData(x) < Abs(OutData(x))) Then PeakData(x) = OutData(x)     'peak of outdata
If (ExData(x) < Abs(PeakData(x))) Then ExData(x) = PeakData(x)     'peak of outdata
Next x

pPen = CreatePen(0, 1, RGB(128, 0, 0))
oldPen = SelectObject(hdc, pPen)
MoveToEx hdc, x, Height, p

For x = 0 To UBound(DrawData)
yy = Height - Sqr(Abs(ExData(x)) / 1.5)

LineTo hdc, x, yy
ExData(x) = ExData(x) - ExData(x) * 0.1
Next

DeleteObject pPen



For x = 0 To UBound(DrawData)

MoveToEx hdc, x, Height, p
yy = Height - Sqr(Abs(PeakData(x)) / 2)
pPen = CreatePen(0, 1, RGB(255, 128, 0))
SelectObject hdc, pPen
LineTo hdc, x, yy

PeakData(x) = PeakData(x) - PeakData(x) * 0.5
DeleteObject pPen
Next

SelectObject hdc, oldPen


Exit Sub
e:


End Sub

Private Sub Class_Initialize()
DoReverse

End Sub


