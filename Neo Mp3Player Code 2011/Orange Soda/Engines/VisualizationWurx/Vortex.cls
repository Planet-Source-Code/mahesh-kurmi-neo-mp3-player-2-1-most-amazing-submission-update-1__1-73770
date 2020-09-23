VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "FrameFlow"
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


Private p As POINTAPI, pen As Long, op As Integer

Public Sub drawVis(hdc As Long, DrawData() As Integer, Height As Single, Width As Single)

Static i As Single, j As Single, x As Integer, y As Integer
Static c As Integer
Static yy As Single, newx As Single, newy As Single
Static wxData(0 To 1023) As Single
Static nx As Single, ny As Single, a As Single, r As Single

BitBlt hdc, 0, r, Width, Height, frmBuf.buf.hdc, 0, 0, vbSrcCopy

For x = 0 To 1023 Step 25
If (wxData(x) < DrawData(x)) Then wxData(x) = DrawData(x)
yy = Sqr(Abs(wxData(x))) / 10

If r * 2 < yy Then
 r = yy
End If

newx = x
newy = Sin(Angle) * yy + (Height / 2 - yy)

Col = RGB((yy) * s1, yy * s2, yy * s3) 'yy*10,0,0
drawCircle newx, newy, yy, Col, hdc
drawCircle Width - newx, newy, yy, Col, hdc


Col = RGB((yy) * s1 * 0.75, yy * s2 * 0.75, yy * s3 * 0.75) 'yy*10,0,0
drawCircle newx, newy, yy + 2, Col, hdc
drawCircle Width - newx, newy, yy + 1, Col, hdc

Col = RGB((yy) * s1 * 0.5, yy * s2 * 0.5, yy * s3 * 0.5) 'yy*10,0,0
drawCircle newx, newy, yy + 3, Col, hdc
drawCircle Width - newx, newy, yy + 2, Col, hdc



Angle = Angle + (Step)

 wxData(x) = wxData(x) - wxData(x) * 0.1
If (Angle > 80) Then Angle = 0

If count > 100 Then
count = 0
s1 = s1 - s1 * 0.1 * IIf(Rnd * 2 > 1, 1, -1)
s2 = s2 - s2 * 0.1 * IIf(Rnd * 2 > 1, 1, -1)
s3 = s3 - s3 * 0.1 * IIf(Rnd * 2 > 1, 1, -1)
Step = Step - Step * 0.5 * IIf(Rnd * 2 > 1, 1, -1)
 
End If

Next x
count = count + 1

   BitBlt frmBuf.buf.hdc, 0, 0, Width, Height, hdc, 0, 0, vbSrcCopy
     r = r - r * 0.1
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

s1 = Rnd * 20 + 1
s2 = Rnd * 20 + 1
s3 = Rnd * 20 + 1
count = 0
End Sub

Public Sub doConfig()
Step = Val(InputBox("Enter step value:"))
End Sub


