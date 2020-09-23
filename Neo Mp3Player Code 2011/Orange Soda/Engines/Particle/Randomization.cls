VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Rotation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private Declare Function SetPixel Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal crColor As Long) As Long
Private NewX As Single, NewY As Single, angle As Single
Private Declare Function LineTo Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long) As Long

Private s1 As Integer, s2 As Integer, s3 As Integer, count As Integer

Public Sub drawVis(hdc As Long, DrawData() As Integer, Height As Single, Width As Single)

Static x As Integer, y As Single

If count > 50 Then
Randomize Timer
s1 = Rnd * 20 + 1
s2 = Rnd * 20 + 1
s3 = Rnd * 20 + 1
count = 0
End If

For x = 0 To UBound(DrawData)
y = Sqr(Abs(DrawData(x))) / 3

NewX = Cos(angle) * x + Width / 2
NewY = Sin(angle) * y + Height / 2

SetPixel hdc, NewX, NewY, RGB(y * s1, y * s2, y * s3)

Next x


angle = angle + (3.14159265 / 180#)

count = count + 1
End Sub


Private Sub Class_Initialize()
Randomize Timer
s1 = Rnd * 20 + 1
s2 = Rnd * 20 + 1
s3 = Rnd * 20 + 1
count = 0
End Sub
