VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Chaos"
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
Private Declare Function Rectangle Lib "gdi32" (ByVal hdc As Long, ByVal x1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
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


Private s1 As Integer, s2 As Integer, s3 As Integer, Col As Long
Private Angle As Single, Step As Single, count As Integer


Private p As POINTAPI, pen As Long, op As Integer

Public Sub drawVis(hdc As Long, DrawData() As Integer, Height As Single, Width As Single)
On Error Resume Next
Static i As Single, j As Single, x As Integer, y As Integer
Static c As Integer
Static yy As Single, newx As Single, newy As Single
Static wxData(0 To 1023) As Single
Static nx As Single, ny As Single, a As Single, r As Single, t As Single


BitBlt hdc, 0, 0, Width, Height, frmBuf.buf.hdc, 0, 0, vbSrcCopy
r = r - r * 0.1

For c = 0 To 180
  Randomize Timer

 i = Rnd * (Width - 1)
 j = Rnd * (Height - 1)


  If i > Height / 2 And j > Width / 2 Then x = 2: y = Rnd * 2
  If i < Height / 2 And j < Width / 2 Then x = -2: y = Rnd * 2
  If i > Height / 2 And j < Width / 2 Then x = Rnd * 2: y = -2
  If i < Height / 2 And j < Width / 2 Then x = Rnd * -2: y = -2

If op < 10 Then
   BitBlt hdc, i + x, j + y, 40, 40, frmBuf.buf.hdc, i, j, vbSrcCopy
   BitBlt hdc, i + x, j + y, 50, 50, frmBuf.buf.hdc, i, j, vbSrcPaint
Else
   BitBlt hdc, i + x, j + y, 50, 50, frmBuf.buf.hdc, i, j, vbSrcCopy
   op = 0
End If

op = op + 1

Next




For x = 0 To 1023 Step 25
If (wxData(x) < DrawData(x)) Then wxData(x) = DrawData(x)
yy = Sqr(Abs(wxData(x))) / 10
If r < yy Then r = yy
newy = x
newx = Tan(Angle) * (Width - yy)

'newx = Cos(Angle * yy / 100) * yy + (Width / 2 - yy)
'newy = Sin(Angle * yy / 100) * yy + (Height / 2 - yy / 2)
    
Col = RGB((yy) * s1, yy * s2, yy * s3) 'yy*10,0,0

drawCircle newx, newy, yy, Col, hdc
drawCircle Width - newx, newy, yy, Col, hdc
'SetPixel hdc, newx, newy, Col


Col = RGB((yy) * s1 * 0.75, yy * s2 * 0.75, yy * s3 * 0.75) 'yy*10,0,0
drawCircle newx, newy, yy + 2, Col, hdc
drawCircle Width - newx, newy, yy + 1, Col, hdc
'SetPixel hdc, newx + 2, newy, Col
'SetPixel hdc, newx - 2, newy, Col
'SetPixel hdc, newx, newy + 2, Col
'SetPixel hdc, newx, newy - 2, Col

Col = RGB((yy) * s1 * 0.5, yy * s2 * 0.5, yy * s3 * 0.5) 'yy*10,0,0
drawCircle newx, newy, yy + 3, Col, hdc
drawCircle Width - newx, newy, yy + 2, Col, hdc
'SetPixel hdc, newx + 4, newy, Col
'SetPixel hdc, newx - 4, newy, Col
'SetPixel hdc, newx, newy + 4, Col
'SetPixel hdc, newx, newy - 4, Col


Angle = Angle + (Step)

 wxData(x) = wxData(x) - wxData(x) * 0.1
If (Angle > 100) Then Angle = 0

If count > 100 Then
count = 0
s1 = s1 - s1 * 0.1 * IIf(Rnd * 2 > 1, 1, -1)
s2 = s2 - s2 * 0.1 * IIf(Rnd * 2 > 1, 1, -1)
s3 = s3 - s3 * 0.1 * IIf(Rnd * 2 > 1, 1, -1)
Step = Step - Step * 0.1 * IIf(Rnd * 2 > 1, 1, -1)
 
End If

Next x
count = count + 1

   BitBlt frmBuf.buf.hdc, 0, 0, Width, Height, hdc, 0, 0, vbSrcCopy
   BitBlt hdc, 0, 0, Width, Height, frmBuf.buf.hdc, 0, 0, vbSrcCopy
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
Step = Rnd * 10 + 1
Step = 1

s1 = Rnd * 10 + 1
s2 = Rnd * 10 + 1
s3 = Rnd * 10 + 1
count = 0
End Sub

Public Sub doConfig()
Step = Val(InputBox("Enter step value:"))
End Sub

