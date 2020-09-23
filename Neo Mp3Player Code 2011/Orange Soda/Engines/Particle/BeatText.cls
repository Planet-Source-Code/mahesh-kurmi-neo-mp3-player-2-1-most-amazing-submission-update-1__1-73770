VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BeatText"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Private Declare Function StretchBlt Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long

Public Sub drawVis(hdc As Long, DrawData() As Integer, Height As Single, Width As Single)
Static x As Integer, y As Single, maxvalue As Single



y = 0
For x = 0 To 1023
y = Sqr(Abs(DrawData(x)))
If maxvalue < y Then maxvalue = y
Next x

y = maxvalue
StretchBlt hdc, Width / 2 - (Width - (y Mod Width)) / 2, Height / 2 - (Height - (y Mod Height)) / 2, Width, Height, frmDummy.buf.hdc, 0, 0, Width - (y Mod Width), Height - (y Mod Height), vbSrcCopy

maxvalue = maxvalue - maxvalue * 0.1


End Sub

Public Sub doConfig()
Dim msg As String

msg = InputBox("Enter new Text:")
If msg <> "" Then
frmDummy.buf.Cls
frmDummy.buf.Print msg

End If

SaveSetting App.EXEName, "Settings", "Text", msg
End Sub

Private Sub Class_Initialize()
Dim msg As String
msg = GetSetting(App.EXEName, "Settings", "Text", "MaheshMp3 Player")
frmDummy.buf.Cls
frmDummy.buf.ForeColor = RGB(146, 215, 230)
frmDummy.buf.Print msg
End Sub
