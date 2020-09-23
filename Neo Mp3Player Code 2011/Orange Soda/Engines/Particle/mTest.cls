VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "mTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Private Declare Function StretchBlt Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function CreatePen Lib "gdi32" (ByVal nPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function MoveToEx Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, lpPoint As POINTAPI) As Long
Private Declare Function SetPixel Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal crColor As Long) As Long
Private Declare Function LineTo Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long) As Long

Private Type POINTAPI
        x As Long
        y As Long
End Type


Public Sub drawVis(hdc As Long, DrawData() As Integer, Height As Single, Width As Single)
frmDummy.picDummy.ScaleWidth = Width
frmDummy.picDummy.ScaleHeight = Height

Static x As Integer, y As Single, i As Single, j!, p As POINTAPI, pp As Long

For x = 0 To 100
Randomize Timer
        intX = (Width - 1) * Rnd
        intY = (Height - 1) * Rnd

        If intX < Width / 2 And intY < Height / 2 Then intI = -2: intJ = Rand(-2, -2)
        If intX > Width / 2 And intY > Height / 2 Then intI = 2: intJ = Rand(2, 2)

        If intX < Width / 2 And intY > Height / 2 Then intI = Rand(-2, -2): intJ = Rand(2, 2)
        If intX > Width / 2 And intY < Height / 2 Then intI = Rand(2, 2): intJ = Rand(-2, -2)


        Call BitBlt(hdc, intX + intI, intY + intJ, 88, 88, frmDummy.picDummy.hdc, intX, intY, vbSrcCopy)
Next x


MoveToEx hdc, 0, Height / 2, p
For x = 0 To 1023

y = Sqr(Abs(DrawData(x)))

'SetPixel hdc, x, (y / 2 + Height / 2), RGB(y Mod 255, y Mod 255, (y * 2) Mod 255)
pp = CreatePen(0, 1, RGB(y Mod 255, y Mod 255, (y * 2) Mod 255))
SelectObject hdc, pp
LineTo hdc, x, y / 2 + Height / 2
SelectObject hdc, -1
DeleteObject pp

Next x

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
 
