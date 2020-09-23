VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Spinner"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private Declare Function CreatePen Lib "gdi32" (ByVal nPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function MoveToEx Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, lpPoint As POINTAPI) As Long
Private Declare Function LineTo Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long) As Long
Private Declare Function SetPixel Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal crColor As Long) As Long
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Type POINTAPI
        x As Long
        y As Long
End Type

Private theta As Single, phase As Single
Private s1 As Integer, s2 As Integer, s3 As Integer, count As Integer

Private BeatValue As Single, maxvalue As Single
Private xx As Single, yy As Single, angle As Double

' spinner
' spinning thingy

Public Sub drawVis(hdc As Long, DrawData() As Integer, Height As Single, Width As Single)

Static x As Integer, y As Single, r As Single
Static NewX As Integer, NewY As Integer, p As POINTAPI, ppen As Long
Static maxvalue As Single
Static PeakData(0 To 1024) As Integer
Static EnvelopeData(0 To 1024) As Integer
frmDummy.picDummy.ScaleWidth = Width
frmDummy.picDummy.ScaleHeight = Height

BitBlt hdc, xx, yy, Width, Height, frmDummy.picBuf3.hdc, 0, 0, vbSrcCopy

MoveToEx hdc, Width / 2, Height / 2, p

BeatValue = 0
For x = 0 To UBound(DrawData)

If count > 50 Then
Randomize Timer
s1 = s1 - Rnd * 2 * IIf(Rnd * 2 < 1, 1, -1)
s2 = s2 - Rnd * 2 * IIf(Rnd * 2 < 1, 1, -1)
s3 = s3 - Rnd * 2 * IIf(Rnd * 2 < 1, 1, -1)
count = 0
End If

If PeakData(x) < DrawData(x) Then PeakData(x) = DrawData(x)
If EnvelopeData(x) < PeakData(x) Then EnvelopeData(x) = PeakData(x)

y = Sqr(Abs(PeakData(x)) / 10)
BeatValue = BeatValue + y
theta = (x * 3.14159265) / 180#
r = Sin(2 * theta + phase) * y
NewX = Sin(theta) * r + Width / 2
NewY = Cos(theta) * r + Height / 2




ppen = CreatePen(0, 1, RGB(y / 3 * s1, y / 3 * s2, y / 3 * s3))

SelectObject hdc, ppen
LineTo hdc, NewX, NewY
SelectObject hdc, -1
DeleteObject ppen

y = Sqr(Abs(EnvelopeData(x)) / 10)
r = Sin(2 * theta + phase - phase * 0.1) * y
NewX = Sin(theta) * r + Width / 2
NewY = Cos(theta) * r + Height / 2

SetPixel hdc, NewX, NewY, RGB(y / 3 * s1 * 0.5, y / 3 * s2 * 0.5, y / 3 * s3 * 0.5)


PeakData(x) = PeakData(x) - PeakData(x) * 0.1
EnvelopeData(x) = EnvelopeData(x) - EnvelopeData(x) * 0.05

Next x
count = count + 1
BeatValue = BeatValue / 1024#
If BeatValue > maxvalue Then
maxvalue = BeatValue
Randomize Timer
xx = Rnd * 10
yy = Rnd * 10
Else
maxvalue = maxvalue - BeatValue
End If

phase = phase + (maxvalue / 180#) * 3.14159256


frmDummy.picBuf3.Cls
BitBlt frmDummy.picBuf3.hdc, 0, 0, Width, Height, hdc, 0, 0, vbSrcCopy


End Sub

Private Sub Class_Initialize()
Randomize Timer
s1 = Rnd * 20 + 1
Randomize Timer
s2 = Rnd * 20 + 1
Randomize Timer
s3 = Rnd * 20 + 1

count = 0
End Sub
