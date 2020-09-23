VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Distortion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function StretchBlt Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long


Public Sub drawVis(hdc As Long, DrawData() As Integer, Height As Single, Width As Single)

Static x As Integer, i As Integer, j As Integer, y As Integer, x1 As Integer
Static yy As Single, Angle As Single

BitBlt hdc, 0, 0, Width, Height, frmBuf.picBuf3.hdc, 0, 0, vbSrcCopy
For x = 0 To 1023 Step 2

yy = Sqr(Abs(DrawData(x))) / 2

 i = x
 j = yy / 2
 
  If i > Height / 2 And j < Width / 2 Then x1 = j: y = -2
  If i < Height / 2 And j < Width / 2 Then x1 = -2: y = -j
  
   BitBlt hdc, i + x1, j + y, 50, Height, frmBuf.picBuf3.hdc, i, j, vbSrcAnd
   BitBlt hdc, i + x1, j + y, 50, Height, frmBuf.picBuf3.hdc, i, j, vbSrcPaint

Next

End Sub
