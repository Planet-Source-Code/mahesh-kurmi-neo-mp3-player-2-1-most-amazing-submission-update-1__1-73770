VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Warp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private Declare Function SetPixel Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal crColor As Long) As Long
Private pos(0 To 10, 0 To 1023) As Integer
Private radii(0 To 10) As Integer, r As Single, nx As Single, ny As Single
Private a As Single
Public Sub drawVis(hdc As Long, DrawData() As Integer, Height As Single, Width As Single)

Dim x As Integer, y As Single
For x = 0 To 1023
a = (360 / 1023) * (3.1415 * x) / 180
y = Sqr(Abs(DrawData(x)))
nx = Cos(a) * y + Width / 2
ny = Sin(a) + Height / 2
SetPixel hdc, x, y, vbWhite
Next x

End Sub
