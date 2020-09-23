VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Psychadelic"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private Declare Function SetPixel Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal crColor As Long) As Long
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long

Public Sub drawVis(hdc As Long, DrawData() As Integer, Height As Single, Width As Single)
Static x As Integer, y As Single, r As Single, theta As Single
Static newx As Single, newy As Single, c
Static flag1 As Integer, flag2 As Integer, flag3 As Integer
Static oldData(0 To 1023) As Single, phase As Single
'BitBlt hdc, 0, 0, Width, Height, frmBuf.buf.hdc, 0, 0, vbSrcPaint

For x = 0 To 1023 Step 20


If oldData(x) < DrawData(x) Then oldData(x) = DrawData(x)
y = Abs(oldData(x)) / 10

Randomize Timer
flag1 = IIf(Rnd * 2 > 1, 5, 0)
flag2 = IIf(Rnd * 2 > 1, 5, 0)
flag3 = IIf(Rnd * 2 > 1, 5, 0)


c = RGB(y * flag1, y * flag2, y * flag3)

For theta = 0 To 2 * 3.14159265 Step 0.1
r = Sin(theta * 10 + phase) * y
newx = Sin(theta) * r + Width / 2
newy = Cos(theta) * r + Height / 2

SetPixel hdc, newx, newy, c
SetPixel hdc, newx + 2, newy, c / 2
SetPixel hdc, newx - 2, newy, c / 2
SetPixel hdc, newx, newy + 2, c / 2
SetPixel hdc, newx, newy - 2, c / 2

Next

oldData(x) = oldData(x) - oldData(x) * 0.1
Next
phase = phase + (3.14159265 / 180)
BitBlt frmBuf.buf.hdc, 0, 0, Width, Height, hdc, 0, 0, vbSrcCopy
End Sub
