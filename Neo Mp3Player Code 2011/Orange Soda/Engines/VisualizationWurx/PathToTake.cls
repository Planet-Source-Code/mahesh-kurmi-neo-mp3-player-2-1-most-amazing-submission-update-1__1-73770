VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ColorSpin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function Ellipse Lib "gdi32" (ByVal hdc As Long, ByVal x1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function CreatePen Lib "gdi32" (ByVal nPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function MoveToEx Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, lpPoint As POINTAPI) As Long
Private Declare Function LineTo Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long) As Long
Private Declare Function SetPixel Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal crColor As Long) As Long
Private Type POINTAPI
        x As Long
        y As Long
End Type


Private s1 As Integer, s2 As Integer, s3 As Integer, Col As Long
Private Angle As Single, Step As Single, count As Integer


Private p As POINTAPI, pen As Long, op As Integer, Flag As Boolean, Phase As Single

Public Sub drawVis(hdc As Long, DrawData() As Integer, Height As Single, Width As Single)

Static i As Single, j As Single, x As Single, y As Single
Static c As Integer
Static yy As Single, newx As Single, newy As Single
Static wxData(0 To 1023) As Single, rad As Single
Static ta As Single, max As Single
'BitBlt hdc, 0, 0, Width, Height, frmBuf.buf.hdc, 0, 0, vbSrcCopy


MoveToEx hdc, 0, Height / 2, p
For x = 0 To 1023 Step 20
If (wxData(x) < DrawData(x)) Then wxData(x) = DrawData(x)
yy = Sqr(Abs(wxData(x))) / 10
If yy > max Then max = yy

newy = x
newx = Tan(Angle) * 150
rad = CDbl((360 / 255) * (x)) * (3.14159265 / 180)
Col = RGB((yy) * s1, yy * s2, yy * s3) 'yy*10,0,0
newx = Sin(rad) * Angle + Width / 2
newy = Cos(rad) * yy + Height / 2

drawStar hdc, newx, newy, yy

newx = Sin(rad - Phase) * Angle + Width / 2
newy = Cos(rad) * yy + Height / 2
drawStar hdc, newx, newy, yy


  
Angle = Angle + (Step)
 wxData(x) = wxData(x) - wxData(x) * 0.1
If (Angle > 100) Then Angle = 0

If count > 100 Then
count = 0
s1 = s1 - Rnd * 2 * IIf(Rnd * 2 < 1, -1, 1)
s2 = s2 - Rnd * 2 * IIf(Rnd * 2 < 1, -1, 1)
s3 = s3 - Rnd * 2 * IIf(Rnd * 2 < 1, -1, 1)

s1 = Abs(s1)
s2 = Abs(s2)
s3 = Abs(s3)
'Step = Rnd * 20 + 1
 
End If
max = max - max * 0.1
Next x
count = count + 1

   BitBlt frmBuf.buf.hdc, 0, 0, Width, Height, hdc, 0, 0, vbSrcCopy
   BitBlt hdc, 0, 0, Width, Height, frmBuf.buf.hdc, 0, 0, vbSrcErase
   For c = 0 To 100
ta = ta + (3.14 / 180 * max * 3)
  Randomize Timer
  
i = Rnd * 10
j = Rnd * 10

x = Cos(ta) * (max * 2 + 100)
y = Sin(ta) * (max * 2 + 100)

If op < 10 Then
   BitBlt hdc, i + x, j + y, 100, 100, frmBuf.buf.hdc, i, j, vbSrcInvert
   BitBlt hdc, i + x, j + y, 100, 100, frmBuf.buf.hdc, i, j, vbSrcPaint
   
   BitBlt hdc, i - x, j - y, 100, 100, frmBuf.buf.hdc, i, j, vbSrcInvert
   BitBlt hdc, i - x, j - y, 100, 100, frmBuf.buf.hdc, i, j, vbSrcPaint

Else
   BitBlt hdc, i + x, j + y, 100, 100, frmBuf.buf.hdc, i, j, vbSrcAnd
     BitBlt hdc, i - x, j - y, 100, 100, frmBuf.buf.hdc, i, j, vbSrcAnd
   op = 0
End If

op = op + 1
Next c

Phase = Phase + (3.14 / 180 * 5)
End Sub

Private Sub drawCircle(x As Single, y As Single, Radius As Single, Color As Long, hdc As Long)

 Dim pPen As Long, oldPen As Long
 pPen = CreatePen(0, 1, Color)
 oldPen = SelectObject(hdc, pPen)
  Ellipse hdc, x - Radius, y - Radius, x + Radius, y + Radius
 SelectObject hdc, oldPen
 DeleteObject pPen
End Sub

Private Sub Class_Initialize()
Randomize Timer
Step = Rnd * 20 + 1

s1 = Rnd * 20 + 1
s2 = Rnd * 20 + 1
s3 = Rnd * 20 + 1
count = 0
Flag = True
End Sub

Public Sub doConfig()
Step = Val(InputBox("Enter step value:"))
End Sub

Private Sub drawStar(hdc As Long, newx As Single, newy As Single, yy As Single)
SetPixel hdc, newx, newy, Col
 SetPixel hdc, newx + 1, newy, Col
 SetPixel hdc, newx, newy + 1, Col
 SetPixel hdc, newx + 1, newy + 1, Col
 
 Col = RGB((yy) * s1 * 0.75, yy * s2 * 0.75, yy * s3 * 0.75) 'yy*10,0,0
 SetPixel hdc, newx, newy + 3, Col
 SetPixel hdc, newx, newy - 3, Col
 SetPixel hdc, newx + 3, newy, Col
 SetPixel hdc, newx - 3, newy, Col
 
 Col = RGB((yy) * s1 * 0.5, yy * s2 * 0.5, yy * s3 * 0.5) 'yy*10,0,0
 SetPixel hdc, newx, newy + 5, Col
 SetPixel hdc, newx, newy - 5, Col
 SetPixel hdc, newx + 5, newy, Col
 SetPixel hdc, newx - 5, newy, Col

End Sub
