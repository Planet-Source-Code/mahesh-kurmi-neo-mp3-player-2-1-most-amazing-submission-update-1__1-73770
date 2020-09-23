VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "EventHorizon"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private Declare Function SetPixel Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal crColor As Long) As Long
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long

Private s1 As Integer, s2 As Integer, s3 As Integer, count As Integer
Private Col As Long, r As Single, amt As Integer



Public Sub drawVis(hdc As Long, DrawData() As Integer, Height As Single, Width As Single)

Static x As Integer, y As Single
Static nx As Single, ny As Single
Static Temp As Single, Pen As Long
Static PeakData(0 To 1023) As Single, ExData(0 To 1023) As Single, Update As Boolean


If count > 100 Then
Randomize Timer
s1 = s1 - Rnd * 2 * IIf(Rnd * 2 < 1, -1, 1)
s2 = s2 - Rnd * 2 * IIf(Rnd * 2 < 1, -1, 1)
s3 = s3 - Rnd * 2 * IIf(Rnd * 2 < 1, -1, 1)

s1 = Abs(s1)
s2 = Abs(s2)
s3 = Abs(s3)

count = 0

'If Step > 0 Then
'Step = Step - 0.01
'Else
'Step = Step + 0.01
'End If
If Not holdStep Then Step = Step - Step * 0.1 * Sgn(Step)

If Not holdFactor Then Factor = Factor - 0.01

End If

angle = 45

For x = 0 To UBound(DrawData)

If PeakData(x) < DrawData(x) Then PeakData(x) = DrawData(x)

If ExData(x) < PeakData(x) Then
 ExData(x) = PeakData(x) * 2
 Update = True
 amt = 255
End If

y = Sqr(Abs(PeakData(x))) / 3

Col = RGB(y * s1, y * s2, y * s3)

r = Sin(Factor * x / 180# * 3.14159265 + phase) * y

NewX = (Tan(angle * 3.14 / 180)) * r * 1.3 + Width / 2
NewY = (Tan(angle * 3.14 / 180)) * y * 1.3 + Height / 2


SetPixel hdc, NewX, NewY, Col

Col = RGB(y * s1 * 0.5, y * s2 * 0.5, y * s3 * 0.5)

 SetPixel hdc, NewX, NewY + 1, Col
 SetPixel hdc, NewX, NewY - 1, Col
 SetPixel hdc, NewX + 1, NewY, Col
 SetPixel hdc, NewX - 1, NewY, Col

Col = RGB(y * s1 * 0.25, y * s2 * 0.25, y * s3 * 0.25)

 SetPixel hdc, NewX, NewY + 3, Col
 SetPixel hdc, NewX, NewY - 3, Col
 SetPixel hdc, NewX + 3, NewY, Col
 SetPixel hdc, NewX - 3, NewY, Col


y = Sqr(Abs(ExData(x))) / 3

Col = RGB(y * s1 * 0.25, y * s2 * 0.25, y * s3 * 0.25)

r = Sin(Factor * x / 180# * 3.14159265 + (phase - phase * 0.1)) * y
NewX = Tan(angle * 3.14 / 180) * r * 1.3 + Width / 2
NewY = Tan(angle * 3.14 / 180) * y * 1.3 + Height / 2



SetPixel hdc, NewX, NewY, Col

Col = RGB(y * s1 * 0.05, y * s2 * 0.05, y * s3 * 0.05)
 SetPixel hdc, NewX, NewY + 2, Col
 SetPixel hdc, NewX, NewY - 2, Col
 SetPixel hdc, NewX + 2, NewY, Col
 SetPixel hdc, NewX - 2, NewY, Col




angle = angle + Step

If angle > 360 Then angle = 0
PeakData(x) = PeakData(x) - PeakData(x) * 0.1
ExData(x) = ExData(x) - ExData(x) * 0.01
Next x


phase = phase + (Step * 5 / 180# * 3.1415926539)
count = count + 1

If Abs(phase) > 180 Then phase = 0

If Update Then
 BitBlt frmDummy.picBuf.hdc, 0, 0, Width, Height, hdc, 0, 0, vbSrcCopy
 Update = False
End If

If amt >= 0 Then
 frmDummy.picBuf2.BackColor = RGB(amt, amt, amt)
 BitBlt frmDummy.picBuf2.hdc, 0, 0, Width, Height, frmDummy.picBuf.hdc, 0, 0, vbSrcAnd
 BitBlt hdc, 1, 1, Width, Height, frmDummy.picBuf2.hdc, 0, 0, vbSrcPaint

End If

 
frmDummy.picBuf2.Cls
amt = amt - 5


End Sub


Private Sub Class_Initialize()
Randomize Timer
s1 = Rnd * 20 + 1
Randomize Timer
s2 = Rnd * 20 + 1
Randomize Timer
s3 = Rnd * 20 + 1

Randomize Timer


Randomize Timer

Step = (Rnd) * 2 + 1


Randomize Timer
Factor = Rnd * 5

count = 0
angle = 0
amt = 255
End Sub

Public Sub doConfig()
frmConfig.Show
End Sub
