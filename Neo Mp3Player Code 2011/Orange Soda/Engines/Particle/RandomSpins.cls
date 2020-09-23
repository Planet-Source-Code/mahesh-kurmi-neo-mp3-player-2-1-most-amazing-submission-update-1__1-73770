VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "RandomSpins"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
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

Option Explicit

Private s_tan As Spiro_Tan
Private s_tan_2 As Spiro_Tan_2

Private count As Single, v As Integer, oldV As Integer, amt As Integer
Private brush As LOGBRUSH

Private Sub Class_Initialize()
Set s_tan = New Spiro_Tan
Set s_tan_2 = New Spiro_Tan_2

count = 0
oldV = 0
Randomize Timer
 v = Rnd * 6
 brush.lbHatch = 0
 brush.lbStyle = 0
 amt = 255
End Sub

Public Sub drawVis(hdc As Long, DrawData() As Integer, Height As Single, Width As Single)
Static clientRect As RECT, br As Long
clientRect.Bottom = Height
clientRect.Top = 0

clientRect.Left = 0
clientRect.Right = Width

If count >= 100 Then
Randomize Timer
 v = Rnd * 6
 count = 0
End If

frmDummy.picDummy.Cls
frmDummy.picBuf.Cls
frmDummy.picBuf2.Cls

Select Case v
 Case 0: s_tan.drawVis frmDummy.picBuf2.hdc, DrawData, Height, Width
 Case 1: s_tan_2.drawVis frmDummy.picBuf2.hdc, DrawData, Height, Width
 End Select
If oldV <> v Then

Select Case oldV
 
 Case 0: s_tan.drawVis frmDummy.picDummy.hdc, DrawData, Height, Width
 Case 1: s_tan_2.drawVis frmDummy.picDummy.hdc, DrawData, Height, Width
 End Select

frmDummy.picBuf.BackColor = RGB(amt, amt, amt)

BitBlt frmDummy.picBuf.hdc, 0, 0, Width, Height, frmDummy.picDummy.hdc, 0, 0, vbSrcAnd
amt = amt - 5
End If
If oldV <> v Then
BitBlt hdc, 0, 0, Width, Height, frmDummy.picBuf2.hdc, 0, 0, vbSrcCopy
BitBlt hdc, 0, 0, Width, Height, frmDummy.picBuf.hdc, 0, 0, vbSrcPaint
Else
BitBlt hdc, 0, 0, Width, Height, frmDummy.picBuf2.hdc, 0, 0, vbSrcCopy
End If
count = count + 1

If amt <= 0 Then
 oldV = v
 amt = 255
End If
End Sub

