VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Randomization"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function FillRect Lib "user32" (ByVal hdc As Long, lpRect As RECT, ByVal hBrush As Long) As Long
Private Declare Function CreateBrushIndirect Lib "gdi32" (lpLogBrush As LOGBRUSH) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long

Private Type LOGBRUSH
        lbStyle As Long
        lbColor As Long
        lbHatch As Long
End Type
Private Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type


Private vchaos As Chaos
Private cSpin As ColorSpin
Private shp As ShapeWurx
Private track As Tracks
Private sWay As SpinWay

Private count As Single, v As Integer, oldV As Integer, amt As Integer
Private brush As LOGBRUSH

Private Sub Class_Initialize()
Set vchaos = New Chaos
Set cSpin = New ColorSpin
Set shp = New ShapeWurx
Set track = New Tracks
Set sWay = New SpinWay

count = 0
oldV = 0
Randomize Timer
 v = Rnd * 4
 brush.lbHatch = 0
 brush.lbStyle = 0
 amt = 255
End Sub

Public Sub drawVis(hdc As Long, DrawData() As Integer, Height As Single, Width As Single)


If count >= 100 Then
Randomize Timer
 v = Rnd * 4
 count = 0
End If


frmBuf.rbuf1.Cls
frmBuf.rbuf2.Cls
frmBuf.rbuf3.Cls

Select Case v
 Case 0: cSpin.drawVis frmBuf.rbuf3.hdc, DrawData, Height, Width
 Case 1: vchaos.drawVis frmBuf.rbuf3.hdc, DrawData, Height, Width
 Case 2: shp.drawVis frmBuf.rbuf3.hdc, DrawData, Height, Width
 Case 3: track.drawVis frmBuf.rbuf3.hdc, DrawData, Height, Width
 Case 4: sWay.drawVis frmBuf.rbuf3.hdc, DrawData, Height, Width
 Case Else: sWay.drawVis frmBuf.rbuf3.hdc, DrawData, Height, Width
 End Select
If oldV <> v Then

Select Case oldV
 
Case 0: cSpin.drawVis frmBuf.rbuf1.hdc, DrawData, Height, Width
 Case 1: vchaos.drawVis frmBuf.rbuf1.hdc, DrawData, Height, Width
 Case 2: shp.drawVis frmBuf.rbuf1.hdc, DrawData, Height, Width
 Case 3: track.drawVis frmBuf.rbuf1.hdc, DrawData, Height, Width
 Case 4: sWay.drawVis frmBuf.rbuf1.hdc, DrawData, Height, Width
   Case Else: sWay.drawVis frmBuf.rbuf1.hdc, DrawData, Height, Width
 End Select
 
frmBuf.rbuf2.BackColor = RGB(amt, amt, amt)

BitBlt frmBuf.rbuf2.hdc, 0, 0, Width, Height, frmBuf.rbuf1.hdc, 0, 0, vbSrcAnd
amt = amt - 5
End If
If oldV <> v Then
BitBlt hdc, 0, 0, Width, Height, frmBuf.rbuf3.hdc, 0, 0, vbSrcCopy
BitBlt hdc, 0, 0, Width, Height, frmBuf.rbuf2.hdc, 0, 0, vbSrcPaint
Else
BitBlt hdc, 0, 0, Width, Height, frmBuf.rbuf3.hdc, 0, 0, vbSrcCopy
End If
count = count + 1

If amt <= 0 Then
 oldV = v
 amt = 255
End If
End Sub

