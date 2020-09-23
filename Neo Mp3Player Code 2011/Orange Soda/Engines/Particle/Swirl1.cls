VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Swirl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Declare Function SetPixel Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal crColor As Long) As Long
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function MoveToEx Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, lpPoint As POINTAPI) As Long
Private Declare Function CreatePen Lib "gdi32" (ByVal nPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function LineTo Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long) As Long

Private Type POINTAPI
        x As Long
        y As Long
End Type


Private angle As Single, phase As Single, theta As Single, p As POINTAPI
Private ppen As Long, Col As Long, r As Single, b As Single, g As Single
Private c As Integer, i As Single, j As Single, op As Long
Public Sub drawVis(hdc As Long, DrawData() As Integer, Height As Single, Width As Single)
Static x As Integer, y As Single
Static xx As Single, yy As Single
Static PeakData(0 To 1023) As Single

BitBlt hdc, 0, 0, Width, Height, frmDummy.picDummy.hdc, 0, 0, vbSrcCopy

For c = 0 To 100
  Randomize Timer
 i = Rnd * (Width - 1)
 j = Rnd * (Height - 1)

  If i > Height / 2 And j > Width / 2 Then x = 2: y = Rnd * 2
  If i < Height / 2 And j < Width / 2 Then x = -2: y = Rnd * 2
  If i > Height / 2 And j < Width / 2 Then x = Rnd * 2: y = -2
  If i < Height / 2 And j < Width / 2 Then x = Rnd * -2: y = -2
If op < 10 Then
   BitBlt hdc, i + x, j + y, 100, 100, frmDummy.picDummy.hdc, i, j, vbSrcInvert
   BitBlt hdc, i + x, j + y, 100, 100, frmDummy.picDummy.hdc, i, j, vbSrcPaint
Else
   BitBlt hdc, i + x, j + y, 100, 100, frmDummy.picDummy.hdc, i, j, vbSrcCopy
   op = 0
End If

op = op + 1
Next c

MoveToEx hdc, 0, Height, p


For x = 0 To 1023
If PeakData(x) < DrawData(x) Then PeakData(x) = DrawData(x)
y = Sqr(Abs(PeakData(x)))
r = r + 0.35 * (0.6 * Sin(0.98 * y) + 0.4 * Sin(1.047 * y))
g = g + 0.35 * (0.6 * Sin(0.835 * y) + 0.4 * Sin(1.081 * y))
b = b + 0.35 * (0.6 * Sin(0.814 * y) + 0.4 * Sin(1.011 * y))

Col = RGB(Abs(r) Mod 255, Abs(g) Mod 255, Abs(b) Mod 255)
ppen = CreatePen(0, 1, Col)
SelectObject hdc, ppen
'LineTo hdc, x, (Height / 2 - y / 2)
LineTo hdc, x, Height - y
SelectObject hdc, -1
DeleteObject ppen

PeakData(x) = PeakData(x) - PeakData(x) * 0.1
Next x

xx = Rnd * 15 * Sgn(IIf(Rnd * 2 > 1, 1, -1))
yy = Rnd * 15 * Sgn(IIf(Rnd * 2 > 1, 1, -1))

'frmDummy.picDummy.Cls
BitBlt frmDummy.picDummy.hdc, 0, 0, Width, Height, hdc, 0, 0, vbSrcCopy
frmDummy.picDummy.Refresh
End Sub

 Public Function Rand(L, U) As Long
    Dim i As Long, U2, L2
    If U < 0 Then
    U2 = U
    Else
    U2 = -U
    End If
    
    If L < 0 Then
    L2 = L
    Else
    L2 = -L
    End If
    
    For i = L To U
        If Int(Rnd * (U2 - L2)) = Int(Rnd * (U2 - L2 + 1)) Then Rand = i: Exit Function
    Next i
End Function
   
