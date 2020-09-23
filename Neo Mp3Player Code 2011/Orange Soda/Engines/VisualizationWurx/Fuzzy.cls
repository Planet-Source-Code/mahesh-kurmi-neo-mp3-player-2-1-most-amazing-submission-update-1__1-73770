VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Fuzzy"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Option Explicit
Private Declare Function CreatePen Lib "gdi32" (ByVal nPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function Ellipse Lib "gdi32" (ByVal hdc As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function FloodFill Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal y As Long, ByVal crColor As Long) As Long
Private Declare Function SetPixel Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal y As Long, ByVal crColor As Long) As Long
Private Declare Function CreateBrushIndirect Lib "gdi32" (lpLogBrush As LOGBRUSH) As Long
Private Type LOGBRUSH
        lbStyle As Long
        lbColor As Long
        lbHatch As Long
End Type

Private red As Long, gr As Long, bl As Long
Private theta As Single, phase As Single, nx As Single, ny As Single, t As Integer
Private r As Single, i As Single
Private hBrush As LOGBRUSH, hBr As Long
Public Sub drawVis(hdc As Long, DrawData() As Integer, Height As Single, Width As Single)

Static X As Integer, y As Single
hBrush.lbColor = Col
hBr = CreateBrushIndirect(hBrush)
SelectObject hdc, hBr
FloodFill hdc, 0, 0, RGB(i, i, i)
SelectObject hdc, -1
DeleteObject hBr

BitBlt hdc, 0, 0, Width, Height, frmBuf.buf.hdc, frmBuf.buf.ScaleWidth / 2, frmBuf.buf.ScaleHeight / 2, vbSrcAnd



For X = 0 To 1023
y = Sqr(Abs(DrawData(X)))

 bl = Abs(Sin(bl + Sin(y) * y))
 gr = Abs(Sin(gr + Cos(y)) * y)
 red = (bl + gr) * 2
 
  For t = 1 To y
  r = t * (Sin(theta + phase))
  nx = (X Mod Width) + Sin(theta) * r
  ny = (X Mod Height) + Cos(theta) * r
  SetPixel hdc, nx, ny, RGB(red Mod 255, gr Mod 255, bl Mod 255)
  theta = theta + (3.1415926539 / 180)
  Next t

Next X
phase = phase + (5 * 3.1415926539 / 180)


BitBlt frmBuf.buf.hdc, frmBuf.buf.ScaleWidth / 2, frmBuf.buf.ScaleHeight / 2, Width, Height, hdc, 0, 0, vbSrcCopy
 frmBuf.buf.Refresh
i = i - 0.1
If i < 0 Then i = 255
End Sub
