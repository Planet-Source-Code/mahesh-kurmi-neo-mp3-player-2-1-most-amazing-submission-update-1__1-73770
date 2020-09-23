VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PondRipple"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Private Declare Function CreatePen Lib "gdi32" (ByVal nPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function Ellipse Lib "gdi32" (ByVal hdc As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function FloodFill Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal crColor As Long) As Long
Private Declare Function SetPixel Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal crColor As Long) As Long

Private shpX(0 To 1023) As Single
Private shpY(0 To 1023) As Single
Private shpSize(0 To 1023) As Single
Private shpR(0 To 1023) As Single
Private shpG(0 To 1023) As Single
Private shpB(0 To 1023) As Single
Private flag As Boolean, pPen As Long, phase As Single, angle As Single
Private nx As Single, ny As Single
Public Sub drawVis(hdc As Long, DrawData() As Integer, Height As Single, Width As Single)

Static x As Integer, y As Single, t As Long
Static i As Single, j As Single, i1!, j1!
If flag Then
 For x = 0 To 1023
 Randomize Timer
  shpX(x) = Rnd * Width
  shpY(x) = Rnd * Height
  
  Randomize Timer
  shpR(x) = Rnd * 5
  shpG(x) = Rnd * 5
  shpB(x) = 128
  shpSize(x) = 0
 Next x
flag = False
End If

BitBlt hdc, 0, 0, Width, Height, frmBuf.buf.hdc, 0, 0, vbSrcCopy


For x = 0 To 1023 Step 25
y = Sqr(Abs(DrawData(x) / 10))
If shpSize(x) < y Then
shpSize(x) = y
shpX(x) = Rnd * Width
  shpY(x) = Rnd * Height
  
  Randomize Timer
  shpR(x) = Rnd * 5
  shpG(x) = Rnd * 5
  shpB(x) = 128
  
End If

t = Abs(shpSize(x) Mod 255)

For i = 1 To 20
Col = RGB(t * shpR(x) / i, t * shpG(x) / i, t * shpB(x) / i)
pPen = CreatePen(0, 1, Col)
SelectObject hdc, pPen
Ellipse hdc, shpX(x) - shpSize(x) - i, shpY(x) - shpSize(x) - i, shpX(x) + shpSize(x) + i, shpY(x) + shpSize(x) + i


SelectObject hdc, -1
DeleteObject pPen
Next i

shpSize(x) = shpSize(x) - shpSize(x) * 0.1
Next x

BitBlt frmBuf.buf.hdc, 0, 0, Width, Height, hdc, 0, 0, vbSrcCopy
 frmBuf.buf.Refresh
 
 angle = angle + (5# / 180 * 3.14)
phase = phase + 1
End Sub

Private Sub Class_Initialize()
flag = True
End Sub
