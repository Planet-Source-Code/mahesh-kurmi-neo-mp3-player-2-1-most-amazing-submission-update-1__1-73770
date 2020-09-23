VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Stars"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Declare Function SetPixel Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal crColor As Long) As Long
Private Declare Function CreatePen Lib "gdi32" (ByVal nPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function MoveToEx Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, lpPoint As POINTAPI) As Long
Private Declare Function LineTo Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long) As Long
Private Type POINTAPI
        x As Long
        y As Long
End Type
Private StarPositionsX(0 To 1023) As Single
Private StarPositionsY(0 To 1023) As Single

Public Sub drawVis(hdc As Long, DrawData() As Integer, Height As Single, Width As Single)

Static c As Integer, x As Single, y As Single, Size As Single, ppen As Long, p As POINTAPI, i As Integer
Static tx As Single, ty As Single, r As Single
Static StarSize(0 To 1023) As Single, Col As Single, flag As Boolean, a As Single

If Not flag Then
For c = 0 To 1023
StarPositionsX(c) = Width / 2
StarPositionsY(c) = Height / 2
Next
flag = True
End If

For c = 0 To UBound(DrawData)

If StarSize(c) < DrawData(c) Then
StarSize(c) = DrawData(c)
Randomize Timer
StarPositionsX(c) = StarPositionsX(c) - StarPositionsX(c) * 0.05 * IIf(Rnd * 2 < 1, 1, -1)
StarPositionsY(c) = StarPositionsY(c) - StarPositionsY(c) * 0.05 * IIf(Rnd * 2 > 1, 1, -1)
If StarPositionsX(c) < 10 Then StarPositionsX(c) = Width / 2
If StarPositionsY(c) < 10 Then StarPositionsY(c) = Height / 2
End If


x = StarPositionsX(c)
y = StarPositionsY(c)
Size = Sqr(Abs(StarSize(c)))
Col = RGB(Abs(Size) * 2, Abs(Size) * 2, Abs(Size) * 2)

SetPixel hdc, x, y, Col
Col = RGB(Abs(Size), Abs(Size), Abs(Size))

 SetPixel hdc, x, y + 2, Col
 SetPixel hdc, x, y - 2, Col
 SetPixel hdc, x + 2, y, Col
 SetPixel hdc, x - 2, y, Col


'For i = 1 To Size
'r = Cos(i * 3.14 / 180 * 360 / Size * 2) * Size
'tx = Sin(i * 3.14 / 180 * 360 / Size) * r
'ty = Cos(i * 3.14 / 180 * 360 / Size) * r
'ppen = CreatePen(0, 1, RGB(Abs(r) * 2, Abs(r) * 2, Abs(r) * 2))
'SelectObject hdc, ppen

'MoveToEx hdc, StarPositionsX(c ), StarPositionsY(c ), p
'LineTo hdc, x + tx, y - ty
'MoveToEx hdc, StarPositionsX(c ), StarPositionsY(c ), p
'LineTo hdc, x + tx, y + ty
'MoveToEx hdc, StarPositionsX(c ), StarPositionsY(c ), p
'LineTo hdc, x - tx, y + ty
'MoveToEx hdc, StarPositionsX(c ), StarPositionsY(c ), p
'LineTo hdc, x - tx, y - ty
'SelectObject hdc, -1
'DeleteObject ppen

'Next i

StarSize(c) = StarSize(c) - StarSize(c) * 0.1
Next c

End Sub

