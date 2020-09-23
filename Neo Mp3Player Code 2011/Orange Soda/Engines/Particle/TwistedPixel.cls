VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TwistedPixel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private Declare Function SetPixel Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal crColor As Long) As Long
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function CreateBrushIndirect Lib "gdi32" (lpLogBrush As LOGBRUSH) As Long
Private Declare Function FloodFill Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal crColor As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long

Private Type LOGBRUSH
        lbStyle As Long
        lbColor As Long
        lbHatch As Long
End Type


Private NewX As Single, NewY As Single
Private Col As Long
Private tp_wave_r As Single, tp_wave_g As Single, tp_wave_b As Single
Private tp_wave_r1 As Single, tp_wave_g1 As Single, tp_wave_b1 As Single
Private tp_cx As Single, tp_cy As Single
Private tp_cx1 As Single, tp_cy1 As Single
Private tp_dx As Single, tp_dy As Single

Private Declare Function TextOut Lib "gdi32" Alias "TextOutA" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal lpString As String, ByVal nCount As Long) As Long
Private Declare Function LineTo Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long) As Long
Private Declare Function FillRect Lib "user32" (ByVal hdc As Long, lpRect As RECT, ByVal hBrush As Long) As Long
Private Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type

Private i(0 To 100) As Single
Private hBrush As LOGBRUSH, hBr As Long, c As Single, r As RECT
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)


Public Sub drawVis(hdc As Long, DrawData() As Integer, Height As Single, Width As Single)

Static x As Integer, y As Single
Static PeakData(0 To 1023) As Integer
Static ExData(0 To 1023) As Integer, i As Integer
Static drawFlag As Boolean

BitBlt hdc, Rnd * phase Mod 2, Rnd * phase Mod 2, Width, Height, frmDummy.picBuf3.hdc, 0, 0, vbSrcCopy

For x = 0 To UBound(DrawData)
If PeakData(x) < DrawData(x) Then PeakData(x) = DrawData(x)
If ExData(x) < PeakData(x) Then ExData(x) = PeakData(x)

y = Sqr(Abs(PeakData(x))) / 10
If phase < y Then
 phase = y
 drawFlag = True
End If


'    tp_wave_r = tp_wave_r + 0.35 * (0.6 * Sin(0.98 * y) + 0.4 * Sin(1.047 * y))
'    tp_wave_g = tp_wave_g + 0.35 * (0.6 * Sin(0.835 * y) + 0.4 * Sin(1.081 * y))
'    tp_wave_b = tp_wave_b + 0.35 * (0.6 * Sin(0.814 * y) + 0.4 * Sin(1.011 * y))
'
'    tp_cx = (tp_cx * 0.05) + 0.11 * (0.6 * Sin(0.374 * y) + 0.4 * Sin(0.294 * y))
'    tp_cy = (tp_cy * 0.05) + 0.11 * (0.6 * Cos(0.393 * y) + 0.4 * Cos(0.223 * y))
'
'    tp_dx = (0.6 * Sin(0.234 * y + phase) + 0.4 * Sin(0.277 * y))
'    tp_dy = (0.6 * Cos(0.284 * y) + 0.4 * Cos(0.247 * y + phase))
'
'
'For i = 1 To 3
'NewX = tp_cx * Width * 3 + (Width / i) * tp_dx + Width / 2
'NewY = tp_cy * Height * 3 + (Height / i) * tp_dy + Height / 2
'
'
'Col = RGB(Abs(tp_wave_r) Mod 255, Abs(tp_wave_g) Mod 255, Abs(tp_wave_b) Mod 255)
'SetPixel hdc, NewX, NewY, Col
'
'
'Col = RGB(Abs(tp_wave_r) * 0.5 Mod 255, Abs(tp_wave_g) * 0.5 Mod 255, Abs(tp_wave_b) * 0.5 Mod 255)
'SetPixel hdc, NewX, NewY + 2, Col
'SetPixel hdc, NewX, NewY - 2, Col
'SetPixel hdc, NewX + 2, NewY, Col
'SetPixel hdc, NewX - 2, NewY, Col
'Next

'exdata
y = Sqr(Abs(ExData(x)))


    tp_wave_r1 = tp_wave_r1 + 0.35 * (0.6 * Sin(0.98 * y) + 0.4 * Sin(1.047 * y))
    tp_wave_g1 = tp_wave_g1 + 0.35 * (0.6 * Sin(0.835 * y) + 0.4 * Sin(1.081 * y))
    tp_wave_b1 = tp_wave_b1 + 0.35 * (0.6 * Sin(0.814 * y) + 0.4 * Sin(1.011 * y))

    tp_cx1 = (tp_cx1 * 0.05) + 0.11 * (0.6 * Sin(0.374 * y) + 0.4 * Sin(0.294 * y))
    tp_cy1 = (tp_cy1 * 0.05) + 0.11 * (0.6 * Cos(0.393 * y) + 0.4 * Cos(0.223 * y))

    tp_dx = (0.6 * Sin(0.234 * y + (phase - phase * 0.1)) + 0.4 * Sin(0.277 * y + (phase - phase * 0.1)))
    tp_dy = (0.6 * Cos(0.284 * y + (phase - phase * 0.1)) + 0.4 * Cos(0.247 * y + (phase - phase * 0.1)))


For i = 1 To 3
NewX = tp_cx1 * Width * 3 + (Width / i) * tp_dx + Width / 2
NewY = tp_cy1 * Height * 3 + (Height / i) * tp_dy + Height / 2


Col = RGB(Abs(tp_wave_r1) Mod 128, Abs(tp_wave_g1) Mod 128, Abs(tp_wave_b1) Mod 128)
SetPixel hdc, NewX, NewY, Col


Col = RGB(Abs(tp_wave_r1) * 0.5 Mod 128, Abs(tp_wave_g1) * 0.5 Mod 128, Abs(tp_wave_b1) * 0.5 Mod 128)
SetPixel hdc, NewX, NewY + 2, Col
SetPixel hdc, NewX, NewY - 2, Col
SetPixel hdc, NewX + 2, NewY, Col
SetPixel hdc, NewX - 2, NewY, Col
Next

PeakData(x) = PeakData(x) - PeakData(x) * 0.5
ExData(x) = ExData(x) - ExData(x) * 0.05
Next

phase = phase - phase * 0.5

c = c - 1
If c <= 0 Then c = 255
If drawFlag Then
frmDummy.picBuf3.Cls
frmDummy.picBuf3.BackColor = RGB(150, 150, 150)
BitBlt frmDummy.picBuf3.hdc, 0, 0, Width, Height, hdc, 0, 0, vbSrcAnd
drawFlag = False
End If

End Sub


Private Sub Class_Initialize()
Randomize Timer
tp_wave_b = Rnd * 100 + 1
tp_wave_g = Rnd * 100 + 1
tp_wave_r = Rnd * 100 + 1
End Sub
