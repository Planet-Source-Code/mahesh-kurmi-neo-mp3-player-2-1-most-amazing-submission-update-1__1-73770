VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Tracks"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private Declare Function Rectangle Lib "gdi32" (ByVal hdc As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function CreatePen Lib "gdi32" (ByVal nPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function Ellipse Lib "gdi32" (ByVal hdc As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function FloodFill Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal y As Long, ByVal crColor As Long) As Long
Private Declare Function CreateBrushIndirect Lib "gdi32" (lpLogBrush As LOGBRUSH) As Long
Private Type LOGBRUSH
        lbStyle As Long
        lbColor As Long
        lbHatch As Long
End Type


Private shpX(0 To 1023) As Single
Private shpY(0 To 1023) As Single
Private shpSize(0 To 1023) As Single
Private shpR(0 To 1023) As Single
Private shpG(0 To 1023) As Single
Private shpB(0 To 1023) As Single


Private flag As Boolean, pPen As Long, Col As Long, t As Single
Private hBrush As LOGBRUSH, hBr As Long

Public Sub drawVis(hdc As Long, DrawData() As Integer, Height As Single, Width As Single)

Static X As Integer

BitBlt hdc, 0, 1, Width, Height, frmBuf.buf.hdc, 0, 0, vbSrcCopy


If flag Then
 For X = 0 To 1023
  Randomize Timer
  shpX(X) = Rnd * Width
  shpY(X) = Rnd * Height
    
  Randomize Timer
  shpR(X) = Rnd * 32
  shpG(X) = Rnd * 32
  shpB(X) = Rnd * 32
  
 Next X
flag = False
End If

For X = 0 To 1023 Step 15

If shpSize(X) < Sqr(Abs(DrawData(X))) / 10 Then
 shpSize(X) = Sqr(Abs(DrawData(X))) / 10
shpX(X) = -shpX(X) + Rnd * 2 * IIf(Rnd * 2 + 1 > 1, 1, -1)
 shpY(X) = -shpY(X) + Rnd * 2 * IIf(Rnd * 2 + 1 > 1, 1, -1)

End If

shpX(X) = X

t = Abs(shpSize(X) Mod 255)

Col = RGB(t * shpR(X), t * shpG(X), t * shpB(X))
pPen = CreatePen(0, 1, Col)
SelectObject hdc, pPen

Ellipse hdc, shpX(X) - shpSize(X), shpY(X) - shpSize(X), shpX(X) + shpSize(X), shpY(X) + shpSize(X)
SelectObject hdc, -1
DeleteObject pPen
hBrush.lbColor = Col
hBr = CreateBrushIndirect(hBrush)
SelectObject hdc, hBr
FloodFill hdc, shpX(X), shpY(X), Col
SelectObject hdc, -1
DeleteObject hBr




shpSize(X) = shpSize(X) - shpSize(X) * 0.1
shpX(X) = shpX(X) + shpX(X) * 0.1 * IIf(Rnd * 2 + 1 > 1, 1, -1)
shpY(X) = shpY(X) + shpY(X) * 0.1 * IIf(Rnd * 2 + 1 > 1, 1, -1)

If shpY(X) > Height Or shpX(X) > Width Then
Randomize Timer
  shpX(X) = Width - Rnd * Width
  shpY(X) = Height - Rnd * Height
End If
Next X

   BitBlt frmBuf.buf.hdc, 0, 0, Width, Height, hdc, 0, 0, vbSrcCopy
 End Sub

Private Sub Class_Initialize()
flag = True
hBrush.lbStyle = 0
hBrush.lbHatch = 0
End Sub

