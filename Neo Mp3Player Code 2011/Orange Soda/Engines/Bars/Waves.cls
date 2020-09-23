VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "FrequencyWaves"
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
Private Type POINTAPI
        x As Long
        y As Long
End Type

Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function CreatePen Lib "gdi32" (ByVal nPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function LineTo Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long) As Long
Private Declare Function MoveToEx Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, lpPoint As POINTAPI) As Long



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
    'Temp = -0.01 * Math.Log((UBound(InData) / 2) - i) ' bassy frequcies are found in the LBound of the array
    InData(i) = InData(i) * (10 / (((UBound(InData) / 2.1) - i)))
    'inData(i) = InData(i) * Temp
   Next
 
End Sub
Public Sub drawVis(hdc As Long, DrawData() As Integer, Height As Single, Width As Single)

On Error GoTo e
Static x As Integer, y As Single, OutData(0 To 1023) As Single
Static ExData(0 To 1023) As Single
Static ColorData(0 To 1023) As Long
FFTAudio DrawData, OutData
For x = 0 To 1023 Step 20
If ExData(x) < OutData(x) Then
 ExData(x) = OutData(x)
 ColorData(x) = RGB(128, 128, 128)
End If
 y = Sqr(Abs(ExData(x)) / 10)
  drawWave hdc, y, ColorData(x), Width, Height
  
 ExData(x) = ExData(x) - ExData(x) * 0.01
  ColorData(x) = ColorData(x) - RGB(y, y, y)
Next x
e:
End Sub

Private Sub Class_Initialize()
DoReverse
End Sub

Private Sub drawWave(hdc As Long, Frequency As Single, Color As Long, Width As Single, Height As Single)

Dim Pen As Long
Dim x As Integer, y As Single, Angle As Single
Static phase As Single, offSet As Single
Static newX As Single, newY As Single

Dim lpPoint As POINTAPI
MoveToEx hdc, 0, Height / 2, lpPoint

Pen = CreatePen(0, 1, Color)
SelectObject hdc, Pen
 For x = 0 To Width Step 2
  Angle = (x * 3.141592659 / 180)
  y = Sin(Angle + phase) * (Frequency Mod (Height / 2))
  offSet = Cos(Angle + phase) * (50) + Height / 2
  
  newX = Sin(Angle + phase) * Frequency
  newY = Cos(Angle) * 50
     LineTo hdc, x + newX, y + offSet + newY
 Next x
 
SelectObject hdc, -1
DeleteObject Pen
phase = phase + (10 / 180 * 3.14159265)
If phase > 100 Then phase = 0
End Sub




