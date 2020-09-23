VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SpikeFlower"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private Declare Function SetPixel Lib "gdi32" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long, ByVal crColor As Long) As Long

Private NewX As Single, NewY As Single
Private Col As Long
Private tp_wave_r As Single, tp_wave_g As Single, tp_wave_b As Single
Private r As Single
Private angle As Single, Step As Integer, Temp As Single

Private Declare Function TextOut Lib "gdi32" Alias "TextOutA" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long, ByVal lpString As String, ByVal nCount As Long) As Long
Private Declare Function LineTo Lib "gdi32" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long) As Long
Private Declare Function MoveToEx Lib "gdi32" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long, lpPoint As POINTAPI) As Long
Private Type POINTAPI
        x As Long
        y As Long
End Type


Private Declare Function CreatePen Lib "gdi32" (ByVal nPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long

Private Declare Function FillRect Lib "user32" (ByVal hDC As Long, lpRect As RECT, ByVal hBrush As Long) As Long
Private Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type



Public Sub drawVis(hDC As Long, DrawData() As Integer, Height As Single, Width As Single)
Static x As Integer, y As Single
Dim dummyPt As POINTAPI
Dim ppen As Long
Static WaveData(0 To 1023) As Integer
Static WaveDataEx(0 To 1023) As Integer
y = Sqr(Abs(WaveData(0)))
NewX = Sin(x * 3.14 / 180) * (y Mod Width) + (Width / 2)
NewY = Cos(x * 3.14 / 180) * (y Mod Height) + (Height / 2)

MoveToEx hDC, Width / 2, Height / 2, dummyPt

For x = 0 To UBound(DrawData)
If WaveData(x) < DrawData(x) Then WaveData(x) = DrawData(x)
If WaveDataEx(x) < WaveData(x) Then WaveDataEx(x) = WaveData(x)
y = Sqr(Abs(WaveData(x)))


    tp_wave_r = tp_wave_r + 0.35 * (0.6 * Sin(0.98 * y) + 0.4 * Sin(1.047 * y))
    tp_wave_g = tp_wave_g + 0.35 * (0.6 * Sin(0.835 * y) + 0.4 * Sin(1.081 * y))
    tp_wave_b = tp_wave_b + 0.35 * (0.6 * Sin(0.814 * y) + 0.4 * Sin(1.011 * y))
    
    
If y < 0 Then y = 1

If Step < y Then Step = y


NewX = Sin(x * 3.14 / 180) * (y Mod Width) + (Width / 2)
NewY = Cos(x * 3.14 / 180) * (y Mod Height) + (Height / 2)



Col = RGB(Abs(tp_wave_r) Mod 255, Abs(tp_wave_g) Mod 255, Abs(tp_wave_b) Mod 255)
ppen = CreatePen(0, 1, Col)
SelectObject hDC, ppen
LineTo hDC, NewX, NewY
SelectObject hDC, -1
DeleteObject ppen

y = Sqr(Abs(WaveDataEx(x))) * 1.2
If y < 0 Then y = 1

If Step < y Then Step = y

NewX = Sin(x * 3.14 / 180) * (y Mod Width) + (Width / 2)
NewY = Cos(x * 3.14 / 180) * (y Mod Height) + (Height / 2)

Col = RGB(((Abs(tp_wave_r) Mod 255)), ((Abs(tp_wave_g) Mod 255)), ((Abs(tp_wave_b) Mod 255)))

SetPixel hDC, NewX + 1, NewY, Col
SetPixel hDC, NewX, NewY + 1, Col
SetPixel hDC, NewX, NewY - 1, Col
SetPixel hDC, NewX - 1, NewY, Col

WaveData(x) = WaveData(x) - WaveData(x) * 0.1
WaveDataEx(x) = WaveDataEx(x) - WaveDataEx(x) * 0.04

Next x


angle = angle + (0.5)
If angle > 360 Then angle = 0


End Sub



Private Sub Class_Initialize()
Randomize Timer
tp_wave_b = Rnd * 100 + 1
tp_wave_g = Rnd * 100 + 1
tp_wave_r = Rnd * 100 + 1
End Sub

