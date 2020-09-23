VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ShapeWurx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private Declare Function Rectangle Lib "gdi32" (ByVal hdc As Long, ByVal x1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function CreatePen Lib "gdi32" (ByVal nPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function Ellipse Lib "gdi32" (ByVal hdc As Long, ByVal x1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function FloodFill Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal crColor As Long) As Long
Private Declare Function CreateBrushIndirect Lib "gdi32" (lpLogBrush As LOGBRUSH) As Long
Private Type LOGBRUSH
        lbStyle As Long
        lbColor As Long
        lbHatch As Long
End Type


Private shpX(0 To 1023) As Single
Private shpY(0 To 1023) As Single
Private shpSize(0 To 1023) As Single
Private shpType(0 To 1023) As Integer
Private shpR(0 To 1023) As Single
Private shpG(0 To 1023) As Single
Private shpB(0 To 1023) As Single


Private Flag As Boolean, pPen As Long, Col As Long, t As Single
Private hBrush As LOGBRUSH, hBr As Long
Private dirX As Integer, dirY As Integer, max As Single, cFlag As Boolean

Public Sub drawVis(hdc As Long, DrawData() As Integer, Height As Single, Width As Single)

Static x As Integer

BitBlt hdc, dirX, dirY, Width, Height, frmBuf.buf.hdc, 0, 0, vbSrcCopy


If Flag Then
 For x = 0 To 1023
  Randomize Timer
  shpX(x) = Rnd * Width
  shpY(x) = Rnd * Height
  shpType(x) = (Rnd * 2)
  
  Randomize Timer
  shpR(x) = Rnd * 32
  shpG(x) = Rnd * 32
  shpB(x) = Rnd * 32
  
 Next x
Flag = False
End If

'For x = 0 To 50
'Randomize Timer
'BitBlt hdc, Width * Rnd, Height * Rnd, 100, 100, frmBuf.buf.hdc, 5 * Rnd, 5 * Rnd, vbSrcCopy
'Next x

For x = 0 To 1023 Step 25

If shpSize(x) < Sqr(Abs(DrawData(x))) / 10 Then
 shpSize(x) = Sqr(Abs(DrawData(x))) / 10
End If

shpX(x) = x

If max < Sqr(Abs(DrawData(x))) / 10 Then
 dirX = -dirX + Rnd * 2 * IIf(Rnd * 2 + 1 > 1, 1, -1)
 dirY = -dirY + Rnd * 2 * IIf(Rnd * 2 + 1 > 1, 1, -1)
 
 max = Sqr(Abs(DrawData(x))) / 10
End If

t = Abs(shpSize(x) Mod 255)

Col = RGB(t * shpR(x), t * shpG(x), t * shpB(x))
pPen = CreatePen(0, 1, Col)
SelectObject hdc, pPen

If shpType(x) > 1 Then
Rectangle hdc, shpX(x) - shpSize(x), shpY(x) - shpSize(x), shpX(x) + shpSize(x), shpY(x) + shpSize(x)
SelectObject hdc, -1
DeleteObject pPen
hBrush.lbColor = Col
hBr = CreateBrushIndirect(hBrush)
SelectObject hdc, hBr
FloodFill hdc, shpX(x), shpY(x), Col
SelectObject hdc, -1
DeleteObject hBr
Else
Ellipse hdc, shpX(x) - shpSize(x), shpY(x) - shpSize(x), shpX(x) + shpSize(x), shpY(x) + shpSize(x)
SelectObject hdc, -1
DeleteObject pPen
hBrush.lbColor = Col
hBr = CreateBrushIndirect(hBrush)
SelectObject hdc, hBr
FloodFill hdc, shpX(x), shpY(x), Col
SelectObject hdc, -1
DeleteObject hBr

End If



shpSize(x) = shpSize(x) - shpSize(x) * 0.1
Next x

'For x = 0 To 1023
'  Randomize Timer
'   shpX(x) = shpX(x) - Rnd * 10 * IIf(Rnd * 2 < 1, -1, 1)
'   shpY(x) = shpY(x) - Rnd * 10 * IIf(Rnd * 2 < 1, -1, 1)
' Next x
   BitBlt frmBuf.buf.hdc, 0, 0, Width, Height, hdc, 0, 0, vbSrcCopy
 max = max - max * 0.1
 dirX = dirX - dirX * 0.1
 dirY = dirY - dirY * 0.1
End Sub

Private Sub Class_Initialize()
Flag = True
hBrush.lbStyle = 0
hBrush.lbHatch = 0
dirX = Rnd * 2
dirY = Rnd * 2
cFlag = False
End Sub
