VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SplashWave"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private Declare Function SetPixel Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal crColor As Long) As Long

Private NewX As Single, NewY As Single
Private Col As Long



Private Declare Function TextOut Lib "gdi32" Alias "TextOutA" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal lpString As String, ByVal nCount As Long) As Long
Private Declare Function LineTo Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long) As Long
Private Declare Function MoveToEx Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, lpPoint As POINTAPI) As Long
Private Type POINTAPI
        x As Long
        y As Long
End Type

Private r1 As Single, r2 As Single
Private g1 As Single, g2 As Single
Private b1 As Single, b2 As Single

Private Declare Function CreatePen Lib "gdi32" (ByVal nPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long

Private Declare Function FillRect Lib "user32" (ByVal hdc As Long, lpRect As RECT, ByVal hBrush As Long) As Long
Private Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type



Public Sub drawVis(hdc As Long, DrawData() As Integer, Height As Single, Width As Single)
Static x As Integer, y As Single
Dim dummyPt As POINTAPI
Dim ppen As Long, ave As Single


Static PeakData(0 To 1023) As Single
Static PeakDataExTop(0 To 1023) As Single

For x = 0 To 1023

If PeakData(x) < (DrawData(x)) Then PeakData(x) = DrawData(x)
If PeakDataExTop(x) < PeakData(x) Then PeakDataExTop(x) = PeakData(x) * 1.2

Next x

For x = 0 To 1023

y = Sqr(Abs(PeakData(x)))

    r1 = r1 + 0.35 * (0.6 * Sin(0.98 * y) + 0.4 * Sin(1.047 * y))
    g1 = g1 + 0.35 * (0.6 * Sin(0.835 * y) + 0.4 * Sin(1.081 * y))
    b1 = b1 + 0.35 * (0.6 * Sin(0.814 * y) + 0.4 * Sin(1.011 * y))

y = Sqr(Abs(PeakDataExTop(x)))

NewX = x
NewY = (Height - y) / 2

Col = RGB(255, 0, 0) 'RGB(Abs(r1 * 0.5), Abs(g1 * 0.5), Abs(b1 * 0.5))
SetPixel hdc, NewX, NewY, Col

'Col = RGB(Abs(r1 * 0.25), Abs(g1 * 0.25), Abs(b1 * 0.25))
Col = RGB(215, 0, 0)
NewY = (Height + y) / 2
SetPixel hdc, NewX, NewY, Col

PeakDataExTop(x) = PeakDataExTop(x) - PeakDataExTop(x) * 0.05
Next x
'------------------
MoveToEx hdc, 0, Height / 2, dummyPt

For x = 0 To 1023

y = Sqr(Abs(PeakData(x)))
    
    r2 = r2 + 0.35 * (0.6 * Sin(0.98 * y) + 0.4 * Sin(1.047 * y))
    g2 = g2 + 0.35 * (0.6 * Sin(0.835 * y) + 0.4 * Sin(1.081 * y))
    b2 = b2 + 0.35 * (0.6 * Sin(0.814 * y) + 0.4 * Sin(1.011 * y))
    

NewX = x
NewY = (Height - y) / 2


Col = RGB(Abs(r2) Mod 255, Abs(g2) Mod 255, Abs(b2) Mod 255)


ppen = CreatePen(0, 1, Col)
SelectObject hdc, ppen

MoveToEx hdc, x, Height / 2, dummyPt
LineTo hdc, NewX, NewY

SelectObject hdc, -1
DeleteObject ppen

' mirror
Col = RGB(Abs(r2 / 2) Mod 128, Abs(g2 / 2) Mod 128, Abs(b2 / 2) Mod 128)


ppen = CreatePen(0, 1, Col)
SelectObject hdc, ppen

MoveToEx hdc, x, Height / 2, dummyPt
LineTo hdc, NewX, Height - NewY * 1.1

'MoveToEx hdc, x, 0, dummyPt
'LineTo hdc, NewX, Height - NewY

SelectObject hdc, -1
DeleteObject ppen

PeakData(x) = PeakData(x) - PeakData(x) * 0.1
'PeakData(x) = PeakData(x) - 800

Next x

End Sub

Private Sub Class_Initialize()
r1 = Rnd
r2 = Rnd

g1 = Rnd
g2 = Rnd

b1 = Rnd
b2 = Rnd
End Sub
