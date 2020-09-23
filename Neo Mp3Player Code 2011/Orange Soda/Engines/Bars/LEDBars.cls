VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "LEDBars"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, _
                                        ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, _
                                        ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, _
                                        ByVal dwRop As Long) As Long
Private Declare Function SetPixel Lib "gdi32" _
                                        (ByVal hDc As Long, ByVal X As Long, ByVal Y As Long, _
                                        ByVal crColor As Long) As Long

Private Declare Function StretchBlt Lib "gdi32" (ByVal hDc As Long, _
                                        ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, _
                                        ByVal nHeight As Long, ByVal hSrcDC As Long, _
                                        ByVal xSrc As Long, ByVal ySrc As Long, _
                                        ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, _
                                        ByVal dwRop As Long) As Long

Dim Data(0 To 255) As Single  ' Stores y position for BitBlt
Dim Data2(0 To 255) As Single
Dim Data3(0 To 255) As Single ' Stores FFT Data
Dim bolStart(0 To 255) As Boolean
Dim bolDelay(0 To 255) As Single
Dim peakDelay As Long

Public Sub drawVis(hDc As Long, Indata() As Single, Height As Single, Width As Single)
    Dim i As Long
    Dim X As Long
    Dim Y As Long
    Dim H As Long
    Dim r As Long
    Dim tmpColor As Long
  '
  '  If frmConfig.picTop.ScaleHeight <> Height Then frmConfig.picTop.Height = Height
   ' If frmConfig.picTop.ScaleWidth <> Width Then
    '    frmConfig.picTop.Width = Width
        If tSpec.BarNumber > 0 Then
            tSpec.BarWidth = (Width - tSpec.BarNumber * tSpec.ScaleW) \ tSpec.BarNumber
        End If
    End If
    
    ' Loop
    For i = 1 To 255
        Data2(i) = Indata(i)
        If (Data2(i) + (Data2(i) * (i * 0.25))) * 40 > Data(i) Then
            Data(i) = (Data2(i) + (Data2(i) * (i * 0.25))) * 40
        Else
            Data(i) = Data(i) - 1
        End If
    Next
    
    r = 0
    X = 1
    Select Case tSpec.DrawStyle
        Case 0
            '******************Draw Bar********************
            For r = 0 To 248
                Y = (Data(r) + Data(r + 1) + Data(r + 2) + Data(r + 3) + Data(r + 4) + Data(r + 5) + Data(r + 6) + Data(r + 7)) / 8 * tSpec.Zoom
                If Y > Height Then Y = Height
                    '-------------------Bar------------------
                    If tSpec.ShowBar Then
                        Select Case tSpec.PosY
                            Case 0 'Top
                                Select Case tSpec.PosX
                                    Case 0 'Left
                                        Select Case tSpec.FillMode
                                            Case 0 'Normal
                                                StretchBlt hDc, X, 0, tSpec.BarWidth, Y, frmConfig.picTop.hDc, 0, 0, tSpec.BarWidth, Y, vbSrcCopy
                                            Case 1 'Fire
                                                StretchBlt hDc, X, 0, tSpec.BarWidth, Y, frmConfig.picTop.hDc, 0, (Height - Y), tSpec.BarWidth, Y, vbSrcCopy
                                            Case 2 'Col
                                                StretchBlt hDc, X, 0, tSpec.BarWidth, Y, frmConfig.picTop.hDc, 0, (Height - Y) - (Height - Y) / 2, tSpec.BarWidth, Y, vbSrcCopy
                                        End Select
                                    Case 1 'Middle
                                        Select Case tSpec.FillMode
                                            Case 0 'Normal
                                                StretchBlt hDc, (Width / 2 - X), 0, tSpec.BarWidth, Y, frmConfig.picTop.hDc, 0, 0, tSpec.BarWidth, Y, vbSrcCopy
                                                StretchBlt hDc, (Width / 2 + X), 0, tSpec.BarWidth, Y, frmConfig.picTop.hDc, 0, 0, tSpec.BarWidth, Y, vbSrcCopy
                                            Case 1 'Fire
                                                StretchBlt hDc, Width / 2 - X, 0, tSpec.BarWidth, Y, frmConfig.picTop.hDc, 0, (Height - Y), tSpec.BarWidth, Y, vbSrcCopy
                                                StretchBlt hDc, Width / 2 + X, 0, tSpec.BarWidth, Y, frmConfig.picTop.hDc, 0, (Height - Y), tSpec.BarWidth, Y, vbSrcCopy
                                            Case 2 'Col
                                                StretchBlt hDc, Width / 2 - X, 0, tSpec.BarWidth, Y, frmConfig.picTop.hDc, 0, (Height - Y) - (Height - Y) / 2, tSpec.BarWidth, Y, vbSrcCopy
                                                StretchBlt hDc, Width / 2 + X, 0, tSpec.BarWidth, Y, frmConfig.picTop.hDc, 0, (Height - Y) - (Height - Y) / 2, tSpec.BarWidth, Y, vbSrcCopy
                                        End Select
                                    Case 2 'Right
                                        Select Case tSpec.FillMode
                                            Case 0 'Normal
                                                StretchBlt hDc, Width - X, 0, tSpec.BarWidth, Y, frmConfig.picTop.hDc, 0, 0, tSpec.BarWidth, Y, vbSrcCopy
                                            Case 1 'Fire
                                                StretchBlt hDc, Width - X, 0, tSpec.BarWidth, Y, frmConfig.picTop.hDc, 0, (Height - Y), tSpec.BarWidth, Y, vbSrcCopy
                                            Case 2 'Col
                                                StretchBlt hDc, Width - X, 0, tSpec.BarWidth, Y, frmConfig.picTop.hDc, 0, (Height - Y) - (Height - Y) / 2, tSpec.BarWidth, Y, vbSrcCopy
                                        End Select
                                End Select
                            Case 1 'Middle
                                Select Case tSpec.PosX
                                    Case 0 'Left
                                        Select Case tSpec.FillMode
                                            Case 0 'Normal
                                                StretchBlt hDc, X, Round(Height / 2 - Y / 2), tSpec.BarWidth, Y / 2, frmConfig.picMiddle.hDc, 0, (Height - Y) / 2, tSpec.BarWidth, Y / 2, vbSrcCopy
                                                StretchBlt hDc, X, Round(Height / 2 - 1), tSpec.BarWidth, Y / 2, frmConfig.picMiddle.hDc, 0, (Height / 2), tSpec.BarWidth, Y / 2, vbSrcCopy
                                            Case 1 'Fire
                                                StretchBlt hDc, X, Round(Height / 2 - Y / 2), tSpec.BarWidth, Y / 2, frmConfig.picMiddle.hDc, 0, 0, tSpec.BarWidth, Y / 2, vbSrcCopy
                                                StretchBlt hDc, X, Round(Height / 2 - 1), tSpec.BarWidth, Y / 2, frmConfig.picMiddle.hDc, 0, Height - Y / 2, tSpec.BarWidth, Y / 2, vbSrcCopy
                                            Case 2 'Col
                                                StretchBlt hDc, X, Round(Height / 2 - Y / 2), tSpec.BarWidth, Y / 2, frmConfig.picMiddle.hDc, 0, Y / 2, tSpec.BarWidth, Y / 2, vbSrcCopy
                                                StretchBlt hDc, X, Round(Height / 2 - 1), tSpec.BarWidth, Y / 2, frmConfig.picMiddle.hDc, 0, Height - Y, tSpec.BarWidth, Y / 2, vbSrcCopy
                                        End Select
                                    Case 1 'Middle
                                        Select Case tSpec.FillMode
                                            Case 0 'Normal
                                                StretchBlt hDc, Width / 2 - X, Round(Height / 2 - Y / 2), tSpec.BarWidth, Y / 2, frmConfig.picMiddle.hDc, 0, (Height - Y) / 2, tSpec.BarWidth, Y / 2, vbSrcCopy
                                                StretchBlt hDc, Width / 2 - X, Round(Height / 2 - 1), tSpec.BarWidth, Y / 2, frmConfig.picMiddle.hDc, 0, (Height / 2), tSpec.BarWidth, Y / 2, vbSrcCopy
                                                StretchBlt hDc, Width / 2 + X, Round(Height / 2 - Y / 2), tSpec.BarWidth, Y / 2, frmConfig.picMiddle.hDc, 0, (Height - Y) / 2, tSpec.BarWidth, Y / 2, vbSrcCopy
                                                StretchBlt hDc, Width / 2 + X, Round(Height / 2 - 1), tSpec.BarWidth, Y / 2, frmConfig.picMiddle.hDc, 0, (Height / 2), tSpec.BarWidth, Y / 2, vbSrcCopy
                                            Case 1 'Fire
                                                StretchBlt hDc, Width / 2 - X, Round(Height / 2 - Y / 2), tSpec.BarWidth, Y / 2, frmConfig.picMiddle.hDc, 0, 0, tSpec.BarWidth, Y / 2, vbSrcCopy
                                                StretchBlt hDc, Width / 2 - X, Round(Height / 2 - 1), tSpec.BarWidth, Y / 2, frmConfig.picMiddle.hDc, 0, Height - Y / 2, tSpec.BarWidth, Y / 2, vbSrcCopy
                                                StretchBlt hDc, Width / 2 + X, Round(Height / 2 - Y / 2), tSpec.BarWidth, Y / 2, frmConfig.picMiddle.hDc, 0, 0, tSpec.BarWidth, Y / 2, vbSrcCopy
                                                StretchBlt hDc, Width / 2 + X, Round(Height / 2 - 1), tSpec.BarWidth, Y / 2, frmConfig.picMiddle.hDc, 0, Height - Y / 2, tSpec.BarWidth, Y / 2, vbSrcCopy
                                            Case 2 'Col
                                                StretchBlt hDc, Width / 2 - X, Round(Height / 2 - Y / 2), tSpec.BarWidth, Y / 2, frmConfig.picMiddle.hDc, 0, Y / 2, tSpec.BarWidth, Y / 2, vbSrcCopy
                                                StretchBlt hDc, Width / 2 - X, Round(Height / 2 - 1), tSpec.BarWidth, Y / 2, frmConfig.picMiddle.hDc, 0, Height - Y, tSpec.BarWidth, Y / 2, vbSrcCopy
                                                StretchBlt hDc, Width / 2 + X, Round(Height / 2 - Y / 2), tSpec.BarWidth, Y / 2, frmConfig.picMiddle.hDc, 0, Y / 2, tSpec.BarWidth, Y / 2, vbSrcCopy
                                                StretchBlt hDc, Width / 2 + X, Round(Height / 2 - 1), tSpec.BarWidth, Y / 2, frmConfig.picMiddle.hDc, 0, Height - Y, tSpec.BarWidth, Y / 2, vbSrcCopy
                                        End Select
                                    Case 2
                                        Select Case tSpec.FillMode
                                            Case 0 'Normal
                                                StretchBlt hDc, Width - X, Round(Height / 2 - Y / 2), tSpec.BarWidth, Y / 2, frmConfig.picMiddle.hDc, 0, (Height - Y) / 2, tSpec.BarWidth, Y / 2, vbSrcCopy
                                                StretchBlt hDc, Width - X, Round(Height / 2 - 1), tSpec.BarWidth, Y / 2, frmConfig.picMiddle.hDc, 0, (Height / 2), tSpec.BarWidth, Y / 2, vbSrcCopy
                                            Case 1 'Fire
                                                StretchBlt hDc, Width - X, Round(Height / 2 - Y / 2), tSpec.BarWidth, Y / 2, frmConfig.picMiddle.hDc, 0, 0, tSpec.BarWidth, Y / 2, vbSrcCopy
                                                StretchBlt hDc, Width - X, Round(Height / 2 - 1), tSpec.BarWidth, Y / 2, frmConfig.picMiddle.hDc, 0, Height - Y / 2, tSpec.BarWidth, Y / 2, vbSrcCopy
                                            Case 2 'Col
                                                StretchBlt hDc, Width - X, Round(Height / 2 - Y / 2), tSpec.BarWidth, Y / 2, frmConfig.picMiddle.hDc, 0, Y / 2, tSpec.BarWidth, Y / 2, vbSrcCopy
                                                StretchBlt hDc, Width - X, Round(Height / 2 - 1), tSpec.BarWidth, Y / 2, frmConfig.picMiddle.hDc, 0, Height - Y, tSpec.BarWidth, Y / 2, vbSrcCopy
                                        End Select
                                End Select
                            Case 2 'Bottom
                                Select Case tSpec.PosX
                                    Case 0
                                        Select Case tSpec.FillMode
                                            Case 0 'Normal
                                                StretchBlt hDc, X, Round(Height - Y), tSpec.BarWidth, Y, frmConfig.picBottom.hDc, 0, (Height - Y), tSpec.BarWidth, Y, vbSrcCopy
                                            Case 1 'Fire
                                                StretchBlt hDc, X, Round(Height - Y), tSpec.BarWidth, Y, frmConfig.picBottom.hDc, 0, 0, tSpec.BarWidth, Y, vbSrcCopy
                                            Case 2 'Col
                                                StretchBlt hDc, X, Round(Height - Y), tSpec.BarWidth, Y, frmConfig.picBottom.hDc, 0, (Height - Y) - (Height - Y) / 2, tSpec.BarWidth, Y, vbSrcCopy
                                        End Select
                                    Case 1
                                        Select Case tSpec.FillMode
                                            Case 0 'Normal
                                                StretchBlt hDc, Width / 2 - X, Round(Height - Y), tSpec.BarWidth, Y, frmConfig.picBottom.hDc, 0, (Height - Y), tSpec.BarWidth, Y, vbSrcCopy
                                                StretchBlt hDc, Width / 2 + X, Round(Height - Y), tSpec.BarWidth, Y, frmConfig.picBottom.hDc, 0, (Height - Y), tSpec.BarWidth, Y, vbSrcCopy
                                            Case 1 'Fire
                                                StretchBlt hDc, Width / 2 - X, Round(Height - Y), tSpec.BarWidth, Y, frmConfig.picBottom.hDc, 0, 0, tSpec.BarWidth, Y, vbSrcCopy
                                                StretchBlt hDc, Width / 2 + X, Round(Height - Y), tSpec.BarWidth, Y, frmConfig.picBottom.hDc, 0, 0, tSpec.BarWidth, Y, vbSrcCopy
                                            Case 2 'Col
                                                StretchBlt hDc, Width / 2 - X, Round(Height - Y), tSpec.BarWidth, Y, frmConfig.picBottom.hDc, 0, (Height - Y) - (Height - Y) / 2, tSpec.BarWidth, Y, vbSrcCopy
                                                StretchBlt hDc, Width / 2 + X, Round(Height - Y), tSpec.BarWidth, Y, frmConfig.picBottom.hDc, 0, (Height - Y) - (Height - Y) / 2, tSpec.BarWidth, Y, vbSrcCopy
                                        End Select
                                    Case 2
                                        Select Case tSpec.FillMode
                                            Case 0 'Normal
                                                StretchBlt hDc, Width - X, Round(Height - Y), tSpec.BarWidth, Y, frmConfig.picBottom.hDc, 0, (Height - Y), tSpec.BarWidth, Y, vbSrcCopy
                                            Case 1 'Fire
                                                StretchBlt hDc, Width - X, Round(Height - Y), tSpec.BarWidth, Y, frmConfig.picBottom.hDc, 0, 0, tSpec.BarWidth, Y, vbSrcCopy
                                            Case 2 'Col
                                                StretchBlt hDc, Width - X, Round(Height - Y), tSpec.BarWidth, Y, frmConfig.picBottom.hDc, 0, (Height - Y) - (Height - Y) / 2, tSpec.BarWidth, Y, vbSrcCopy
                                        End Select
                                End Select
                            End Select
                        End If
                        '----------------------Peak----------------------
                        If tSpec.PeakDraw = 0 Then
                            If bolDelay(r) = tSpec.PeakInteval Then
                                Data3(r) = Data3(r) - tSpec.PeakDrop / 10
                                bolStart(r) = True
                            End If
                            bolDelay(r) = bolDelay(r) + 1
                            If Data3(r) = Y Then
                                bolDelay(r) = 0
                                bolStart(r) = False
                            End If
                            If Data3(r) > Y And bolStart(r) Then
                                bolDelay(r) = tSpec.PeakInteval
                            End If
                        End If
                        If tSpec.PeakDraw = 1 Then
                            If bolDelay(r) = tSpec.PeakInteval Then
                                Data3(r) = Data3(r) + tSpec.PeakDrop / 10
                                bolStart(r) = True
                            End If
                            bolDelay(r) = bolDelay(r) + 1
                            If Data3(r) = Y Then
                                bolDelay(r) = 0
                                bolStart(r) = False
                            End If
                            If Data3(r) > Y And bolStart(r) Then
                                bolDelay(r) = tSpec.PeakInteval
                            End If
                            If Data3(r) >= Height Then Data3(r) = Y
                            If Data3(r) = Y Then
                                bolDelay(r) = 0
                                bolStart(r) = False
                            End If
                        End If
                        If tSpec.PeakDraw = 2 Then
                            If Data3(r) <= Height And Not (bolStart(r)) Then
                                Data3(r) = Data3(r) + 5 / 10
                                If Data3(r) = Height Then
                                    bolStart(r) = True
                                End If
                            End If
                            If bolStart(r) Then
                                Data3(r) = Data3(r) - 5 / 10
                                bolStart(r) = True
                            End If
                            If Data3(r) = Y Then
                                bolStart(r) = False
                            End If
                        End If
                        If Data3(r) < Y Then
                            Data3(r) = (Y)
                        End If
                        If tSpec.ShowPeak Then
                            Select Case tSpec.PosY
                                Case 0  'Top
                                    Select Case tSpec.PosX
                                        Case 0 'Left
                                            If tSpec.PeakGradient = False Then
                                                BitBlt hDc, X, Round(Data3(r)), tSpec.BarWidth, tSpec.PeakHeight, frmConfig.picPeak.hDc, 0, 0, vbSrcCopy
                                            Else
                                                Select Case tSpec.FillMode
                                                    Case 0 'Normal
                                                        BitBlt hDc, X, Round(Data3(r)), tSpec.BarWidth, tSpec.PeakHeight, frmConfig.picTop.hDc, 0, Y, vbSrcCopy
                                                    Case 1 'Fire
                                                        BitBlt hDc, X, Round(Data3(r)), tSpec.BarWidth, tSpec.PeakHeight, frmConfig.picTop.hDc, 0, Height - 1, vbSrcCopy
                                                    Case 2 'Col
                                                        BitBlt hDc, X, Round(Data3(r)), tSpec.BarWidth, tSpec.PeakHeight, frmConfig.picTop.hDc, 0, (Height - Y) - (Height - Y) / 2, vbSrcCopy
                                                    End Select
                                            End If
                                        Case 1 'Middle
                                            If tSpec.PeakGradient = False Then
                                                BitBlt hDc, Width / 2 - X, Round(Data3(r)), tSpec.BarWidth, tSpec.PeakHeight, frmConfig.picPeak.hDc, 0, 0, vbSrcCopy
                                                BitBlt hDc, Width / 2 + X, Round(Data3(r)), tSpec.BarWidth, tSpec.PeakHeight, frmConfig.picPeak.hDc, 0, 0, vbSrcCopy
                                            Else
                                                Select Case tSpec.FillMode
                                                    Case 0 'Normal
                                                        BitBlt hDc, Width / 2 - X, Round(Data3(r)), tSpec.BarWidth, tSpec.PeakHeight, frmConfig.picTop.hDc, 0, Y, vbSrcCopy
                                                        BitBlt hDc, Width / 2 + X, Round(Data3(r)), tSpec.BarWidth, tSpec.PeakHeight, frmConfig.picTop.hDc, 0, Y, vbSrcCopy
                                                    Case 1 'Fire
                                                        BitBlt hDc, Width / 2 - X, Round(Data3(r)), tSpec.BarWidth, tSpec.PeakHeight, frmConfig.picTop.hDc, 0, Height - 1, vbSrcCopy
                                                        BitBlt hDc, Width / 2 + X, Round(Data3(r)), tSpec.BarWidth, tSpec.PeakHeight, frmConfig.picTop.hDc, 0, Height - 1, vbSrcCopy
                                                    Case 2 'Col
                                                        BitBlt hDc, Width / 2 - X, Round(Data3(r)), tSpec.BarWidth, tSpec.PeakHeight, frmConfig.picTop.hDc, 0, (Height - Y) - (Height - Y) / 2, vbSrcCopy
                                                        BitBlt hDc, Width / 2 + X, Round(Data3(r)), tSpec.BarWidth, tSpec.PeakHeight, frmConfig.picTop.hDc, 0, (Height - Y) - (Height - Y) / 2, vbSrcCopy
                                                End Select
                                            End If
                                        Case 2 'Right
                                            If tSpec.PeakGradient = False Then
                                                BitBlt hDc, Width - X, Round(Data3(r)), tSpec.BarWidth, tSpec.PeakHeight, frmConfig.picPeak.hDc, 0, 0, vbSrcCopy
                                            Else
                                                Select Case tSpec.FillMode
                                                    Case 0 'Normal
                                                        BitBlt hDc, Width - X, Round(Data3(r)), tSpec.BarWidth, tSpec.PeakHeight, frmConfig.picTop.hDc, 0, Y, vbSrcCopy
                                                    Case 1 'Fire
                                                        BitBlt hDc, Width - X, Round(Data3(r)), tSpec.BarWidth, tSpec.PeakHeight, frmConfig.picTop.hDc, 0, Height - 1, vbSrcCopy
                                                    Case 2 'Col
                                                        BitBlt hDc, Width - X, Round(Data3(r)), tSpec.BarWidth, tSpec.PeakHeight, frmConfig.picTop.hDc, 0, (Height - Y) - (Height - Y) / 2, vbSrcCopy
                                                End Select
                                            End If
                                    End Select
                                Case 1 'Middle
                                    Select Case tSpec.PosX
                                        Case 0 'Left
                                            If tSpec.PeakGradient = False Then
                                                BitBlt hDc, X, Round((Height - Data3(r)) / 2), tSpec.BarWidth, tSpec.PeakHeight, frmConfig.picPeak.hDc, 0, 0, vbSrcCopy
                                                BitBlt hDc, X, Round((Height + Data3(r)) / 2), tSpec.BarWidth, tSpec.PeakHeight, frmConfig.picPeak.hDc, 0, 0, vbSrcCopy
                                            Else
                                                Select Case tSpec.FillMode
                                                    Case 0 'Normal
                                                        BitBlt hDc, X, Round((Height - Data3(r)) / 2), tSpec.BarWidth, tSpec.PeakHeight, frmConfig.picMiddle.hDc, 0, (Height - Data3(r)) / 2, vbSrcCopy
                                                        BitBlt hDc, X, Round((Height + Data3(r)) / 2), tSpec.BarWidth, tSpec.PeakHeight, frmConfig.picMiddle.hDc, 0, (Height - Data3(r)) / 2, vbSrcCopy
                                                    Case 1 'Fire
                                                        BitBlt hDc, X, Round((Height - Data3(r)) / 2), tSpec.BarWidth, tSpec.PeakHeight, frmConfig.picMiddle.hDc, 0, Height - Data3(r) / 2, vbSrcCopy
                                                        BitBlt hDc, X, Round((Height + Data3(r)) / 2), tSpec.BarWidth, tSpec.PeakHeight, frmConfig.picMiddle.hDc, 0, Height - Data3(r) / 2, vbSrcCopy
                                                    Case 2 'Col
                                                        BitBlt hDc, X, Round((Height - Data3(r)) / 2), tSpec.BarWidth, tSpec.PeakHeight, frmConfig.picMiddle.hDc, 0, Height - Data3(r), vbSrcCopy
                                                        BitBlt hDc, X, Round((Height + Data3(r)) / 2), tSpec.BarWidth, tSpec.PeakHeight, frmConfig.picMiddle.hDc, 0, Height - Data3(r), vbSrcCopy
                                                End Select
                                            End If
                                        Case 1 'Middle
                                            If tSpec.PeakGradient = False Then
                                                BitBlt hDc, Width / 2 - X, Round((Height - Data3(r)) / 2), tSpec.BarWidth, tSpec.PeakHeight, frmConfig.picPeak.hDc, 0, 0, vbSrcCopy
                                                BitBlt hDc, Width / 2 - X, Round((Height + Data3(r)) / 2), tSpec.BarWidth, tSpec.PeakHeight, frmConfig.picPeak.hDc, 0, 0, vbSrcCopy
                                                BitBlt hDc, Width / 2 + X, Round((Height - Data3(r)) / 2), tSpec.BarWidth, tSpec.PeakHeight, frmConfig.picPeak.hDc, 0, 0, vbSrcCopy
                                                BitBlt hDc, Width / 2 + X, Round((Height + Data3(r)) / 2), tSpec.BarWidth, tSpec.PeakHeight, frmConfig.picPeak.hDc, 0, 0, vbSrcCopy
                                            Else
                                                Select Case tSpec.FillMode
                                                    Case 0 'Normal
                                                        BitBlt hDc, Width / 2 - X, Round((Height - Data3(r)) / 2), tSpec.BarWidth, tSpec.PeakHeight, frmConfig.picMiddle.hDc, 0, (Height - Data3(r)) / 2, vbSrcCopy
                                                        BitBlt hDc, Width / 2 - X, Round((Height + Data3(r)) / 2), tSpec.BarWidth, tSpec.PeakHeight, frmConfig.picMiddle.hDc, 0, (Height - Data3(r)) / 2, vbSrcCopy
                                                        BitBlt hDc, Width / 2 + X, Round((Height - Data3(r)) / 2), tSpec.BarWidth, tSpec.PeakHeight, frmConfig.picMiddle.hDc, 0, (Height - Data3(r)) / 2, vbSrcCopy
                                                        BitBlt hDc, Width / 2 + X, Round((Height + Data3(r)) / 2), tSpec.BarWidth, tSpec.PeakHeight, frmConfig.picMiddle.hDc, 0, (Height - Data3(r)) / 2, vbSrcCopy
                                                    Case 1 'Fire
                                                        BitBlt hDc, Width / 2 - X, Round((Height - Data3(r)) / 2), tSpec.BarWidth, tSpec.PeakHeight, frmConfig.picMiddle.hDc, 0, Height - Data3(r) / 2, vbSrcCopy
                                                        BitBlt hDc, Width / 2 - X, Round((Height + Data3(r)) / 2), tSpec.BarWidth, tSpec.PeakHeight, frmConfig.picMiddle.hDc, 0, Height - Data3(r) / 2, vbSrcCopy
                                                        BitBlt hDc, Width / 2 + X, Round((Height - Data3(r)) / 2), tSpec.BarWidth, tSpec.PeakHeight, frmConfig.picMiddle.hDc, 0, Height - Data3(r) / 2, vbSrcCopy
                                                        BitBlt hDc, Width / 2 + X, Round((Height + Data3(r)) / 2), tSpec.BarWidth, tSpec.PeakHeight, frmConfig.picMiddle.hDc, 0, Height - Data3(r) / 2, vbSrcCopy
                                                    Case 2 'Col
                                                        BitBlt hDc, Width / 2 - X, Round((Height - Data3(r)) / 2), tSpec.BarWidth, tSpec.PeakHeight, frmConfig.picMiddle.hDc, 0, Height - Data3(r), vbSrcCopy
                                                        BitBlt hDc, Width / 2 - X, Round((Height + Data3(r)) / 2), tSpec.BarWidth, tSpec.PeakHeight, frmConfig.picMiddle.hDc, 0, Height - Data3(r), vbSrcCopy
                                                        BitBlt hDc, Width / 2 + X, Round((Height - Data3(r)) / 2), tSpec.BarWidth, tSpec.PeakHeight, frmConfig.picMiddle.hDc, 0, Height - Data3(r), vbSrcCopy
                                                        BitBlt hDc, Width / 2 + X, Round((Height + Data3(r)) / 2), tSpec.BarWidth, tSpec.PeakHeight, frmConfig.picMiddle.hDc, 0, Height - Data3(r), vbSrcCopy
                                                End Select
                                            End If
                                        Case 2 'Right
                                            If tSpec.PeakGradient = False Then
                                                BitBlt hDc, Width - X, Round((Height - Data3(r)) / 2), tSpec.BarWidth, tSpec.PeakHeight, frmConfig.picPeak.hDc, 0, 0, vbSrcCopy
                                                BitBlt hDc, Width - X, Round((Height + Data3(r)) / 2), tSpec.BarWidth, tSpec.PeakHeight, frmConfig.picPeak.hDc, 0, 0, vbSrcCopy
                                            Else
                                                Select Case tSpec.FillMode
                                                    Case 0 'Normal
                                                        BitBlt hDc, Width - X, Round((Height - Data3(r)) / 2), tSpec.BarWidth, tSpec.PeakHeight, frmConfig.picMiddle.hDc, 0, (Height - Data3(r)) / 2, vbSrcCopy
                                                        BitBlt hDc, Width - X, Round((Height + Data3(r)) / 2), tSpec.BarWidth, tSpec.PeakHeight, frmConfig.picMiddle.hDc, 0, (Height - Data3(r)) / 2, vbSrcCopy
                                                    Case 1 'Fire
                                                        BitBlt hDc, Width - X, Round((Height - Data3(r)) / 2), tSpec.BarWidth, tSpec.PeakHeight, frmConfig.picMiddle.hDc, 0, Height - Data3(r) / 2, vbSrcCopy
                                                        BitBlt hDc, Width - X, Round((Height + Data3(r)) / 2), tSpec.BarWidth, tSpec.PeakHeight, frmConfig.picMiddle.hDc, 0, Height - Data3(r) / 2, vbSrcCopy
                                                    Case 2 'Col
                                                        BitBlt hDc, Width - X, Round((Height - Data3(r)) / 2), tSpec.BarWidth, tSpec.PeakHeight, frmConfig.picMiddle.hDc, 0, Height - Data3(r), vbSrcCopy
                                                        BitBlt hDc, Width - X, Round((Height + Data3(r)) / 2), tSpec.BarWidth, tSpec.PeakHeight, frmConfig.picMiddle.hDc, 0, Height - Data3(r), vbSrcCopy
                                                End Select
                                            End If
                                    End Select
                                Case 2 'Bottom
                                    Select Case tSpec.PosX
                                        Case 0
                                            If tSpec.PeakGradient = False Then
                                                BitBlt hDc, X, Round(Height - Data3(r)), tSpec.BarWidth, tSpec.PeakHeight, frmConfig.picPeak.hDc, 0, 0, vbSrcCopy
                                            Else
                                                Select Case tSpec.FillMode
                                                    Case 0 'Normal
                                                        BitBlt hDc, X, Round(Height - Data3(r)), tSpec.BarWidth, tSpec.PeakHeight, frmConfig.picBottom.hDc, 0, Height - Y, vbSrcCopy
                                                    Case 1 'Fire
                                                        BitBlt hDc, X, Round(Height - Data3(r)), tSpec.BarWidth, tSpec.PeakHeight, frmConfig.picBottom.hDc, 0, 0, vbSrcCopy
                                                    Case 2 'Col
                                                        BitBlt hDc, X, Round(Height - Data3(r)), tSpec.BarWidth, tSpec.PeakHeight, frmConfig.picBottom.hDc, 0, (Height - Y) - (Height - Y) / 2, vbSrcCopy
                                                End Select
                                            End If
                                        Case 1
                                            If tSpec.PeakGradient = False Then
                                                BitBlt hDc, Width / 2 - X, Round(Height - Data3(r)), tSpec.BarWidth, tSpec.PeakHeight, frmConfig.picPeak.hDc, 0, 0, vbSrcCopy
                                                BitBlt hDc, Width / 2 + X, Round(Height - Data3(r)), tSpec.BarWidth, tSpec.PeakHeight, frmConfig.picPeak.hDc, 0, 0, vbSrcCopy
                                            Else
                                                Select Case tSpec.FillMode
                                                    Case 0 'Normal
                                                        BitBlt hDc, Width / 2 - X, Round(Height - Data3(r)), tSpec.BarWidth, tSpec.PeakHeight, frmConfig.picBottom.hDc, 0, Height - Y, vbSrcCopy
                                                        BitBlt hDc, Width / 2 + X, Round(Height - Data3(r)), tSpec.BarWidth, tSpec.PeakHeight, frmConfig.picBottom.hDc, 0, Height - Y, vbSrcCopy
                                                    Case 1 'Fire
                                                        BitBlt hDc, Width / 2 - X, Round(Height - Data3(r)), tSpec.BarWidth, tSpec.PeakHeight, frmConfig.picBottom.hDc, 0, 0, vbSrcCopy
                                                        BitBlt hDc, Width / 2 + X, Round(Height - Data3(r)), tSpec.BarWidth, tSpec.PeakHeight, frmConfig.picBottom.hDc, 0, 0, vbSrcCopy
                                                    Case 2 'Col
                                                        BitBlt hDc, Width / 2 - X, Round(Height - Data3(r)), tSpec.BarWidth, tSpec.PeakHeight, frmConfig.picBottom.hDc, 0, (Height - Y) - (Height - Y) / 2, vbSrcCopy
                                                        BitBlt hDc, Width / 2 + X, Round(Height - Data3(r)), tSpec.BarWidth, tSpec.PeakHeight, frmConfig.picBottom.hDc, 0, (Height - Y) - (Height - Y) / 2, vbSrcCopy
                                                End Select
                                            End If
                                        Case 2
                                            If tSpec.PeakGradient = False Then
                                                BitBlt hDc, Width - X, Round(Height - Data3(r)), tSpec.BarWidth, tSpec.PeakHeight, frmConfig.picPeak.hDc, 0, 0, vbSrcCopy
                                            Else
                                                Select Case tSpec.FillMode
                                                    Case 0 'Normal
                                                        BitBlt hDc, Width - X, Round(Height - Data3(r)), tSpec.BarWidth, tSpec.PeakHeight, frmConfig.picBottom.hDc, 0, Height - Y, vbSrcCopy
                                                    Case 1 'Fire
                                                        BitBlt hDc, Width - X, Round(Height - Data3(r)), tSpec.BarWidth, tSpec.PeakHeight, frmConfig.picBottom.hDc, 0, 0, vbSrcCopy
                                                    Case 2 'Col
                                                        BitBlt hDc, Width - X, Round(Height - Data3(r)), tSpec.BarWidth, tSpec.PeakHeight, frmConfig.picBottom.hDc, 0, (Height - Y) - (Height - Y) / 2, vbSrcCopy
                                                End Select
                                            End If
                                    End Select
                            End Select
                        End If
                        Select Case tSpec.PosX
                            Case 0, 2
                                r = r + Int(256 / tSpec.BarNumber)
                            Case 1
                                r = r + Int(128 / tSpec.BarNumber)
                        End Select
                        X = X + tSpec.BarWidth + tSpec.ScaleW
                Next r
            '****************Line***************
            Case 1
                Dim W As Long
                W = IIf(Width < 1023, Width, 1023)
                For i = 0 To W
                    H = (Sqrt(Indata(i + 1)) * Height) * tSpec.Zoom
                    If H > Height Then H = Height
                    X = X + 1
                    'draw line
                    Select Case tSpec.PosY
                        Case 0
                            Select Case tSpec.PosX
                                Case 0
                                    Select Case tSpec.FillMode
                                        Case 0 'Normal
                                            BitBlt hDc, X, 0, 1, Round(H), frmConfig.picTop.hDc, 0, 0, vbSrcCopy
                                        Case 1
                                            BitBlt hDc, X, 0, 1, Round(H), frmConfig.picTop.hDc, 0, Height - H, vbSrcCopy
                                        Case 2
                                            BitBlt hDc, X, 0, 1, Round(H), frmConfig.picTop.hDc, 0, (Height - H) / 2, vbSrcCopy
                                    End Select
                                Case 1
                                    Select Case tSpec.FillMode
                                        Case 0 'Normal
                                            BitBlt hDc, Width / 2 - X, 0, 1, Round(H), frmConfig.picTop.hDc, 0, 0, vbSrcCopy
                                            BitBlt hDc, Width / 2 + X, 0, 1, Round(H), frmConfig.picTop.hDc, 0, 0, vbSrcCopy
                                        Case 1
                                            BitBlt hDc, Width / 2 - X, 0, 1, Round(H), frmConfig.picTop.hDc, 0, Height - H, vbSrcCopy
                                            BitBlt hDc, Width / 2 + X, 0, 1, Round(H), frmConfig.picTop.hDc, 0, Height - H, vbSrcCopy
                                        Case 2
                                            BitBlt hDc, Width / 2 - X, 0, 1, Round(H), frmConfig.picTop.hDc, 0, (Height - H) / 2, vbSrcCopy
                                            BitBlt hDc, Width / 2 + X, 0, 1, Round(H), frmConfig.picTop.hDc, 0, (Height - H) / 2, vbSrcCopy
                                        End Select
                                Case 2
                                    Select Case tSpec.FillMode
                                        Case 0 'Normal
                                            BitBlt hDc, Width - X, 0, 1, Round(H), frmConfig.picTop.hDc, 0, 0, vbSrcCopy
                                        Case 1
                                            BitBlt hDc, Width - X, 0, 1, Round(H), frmConfig.picTop.hDc, 0, Height - H, vbSrcCopy
                                        Case 2
                                            BitBlt hDc, Width - X, 0, 1, Round(H), frmConfig.picTop.hDc, 0, (Height - H) / 2, vbSrcCopy
                                    End Select
                            End Select
                        Case 1
                            Select Case tSpec.PosX
                                Case 0
                                    Select Case tSpec.FillMode
                                        Case 0 'Normal
                                            BitBlt hDc, X, Round((Height - H) / 2), 1, Round(H / 2), frmConfig.picMiddle.hDc, 0, (Height - H) / 2, vbSrcCopy
                                            BitBlt hDc, X, Round(Height / 2 - 1), 1, Round(H / 2), frmConfig.picMiddle.hDc, 0, (Height) / 2, vbSrcCopy
                                        Case 1
                                            BitBlt hDc, X, Round((Height - H) / 2), 1, Round(H / 2), frmConfig.picMiddle.hDc, 0, 0, vbSrcCopy
                                            BitBlt hDc, X, Round(Height / 2 - 1), 1, Round(H / 2), frmConfig.picMiddle.hDc, 0, Height - H / 2, vbSrcCopy
                                        Case 2
                                            BitBlt hDc, X, Round((Height - H) / 2), 1, Round(H / 2), frmConfig.picMiddle.hDc, 0, (Height - H) / 4, vbSrcCopy
                                            BitBlt hDc, X, Round(Height / 2 - 1), 1, Round(H / 2), frmConfig.picMiddle.hDc, 0, Height / 2 + (Height - H) / 4, vbSrcCopy
                                    End Select
                                Case 1
                                    Select Case tSpec.FillMode
                                        Case 0 'Normal
                                            BitBlt hDc, Width / 2 - X, Round((Height - H) / 2), 1, Round(H / 2), frmConfig.picMiddle.hDc, 0, (Height - H) / 2, vbSrcCopy
                                            BitBlt hDc, Width / 2 - X, Round(Height / 2 - 1), 1, Round(H / 2), frmConfig.picMiddle.hDc, 0, (Height) / 2, vbSrcCopy
                                            BitBlt hDc, Width / 2 + X, Round((Height - H) / 2), 1, Round(H / 2), frmConfig.picMiddle.hDc, 0, (Height - H) / 2, vbSrcCopy
                                            BitBlt hDc, Width / 2 + X, Round(Height / 2 - 1), 1, Round(H / 2), frmConfig.picMiddle.hDc, 0, (Height) / 2, vbSrcCopy
                                        Case 1
                                            BitBlt hDc, Width / 2 - X, Round((Height - H) / 2), 1, Round(H / 2), frmConfig.picMiddle.hDc, 0, 0, vbSrcCopy
                                            BitBlt hDc, Width / 2 - X, Round(Height / 2 - 1), 1, Round(H / 2), frmConfig.picMiddle.hDc, 0, Height - H / 2, vbSrcCopy
                                            BitBlt hDc, Width / 2 + X, Round((Height - H) / 2), 1, Round(H / 2), frmConfig.picMiddle.hDc, 0, 0, vbSrcCopy
                                            BitBlt hDc, Width / 2 + X, Round(Height / 2 - 1), 1, Round(H / 2), frmConfig.picMiddle.hDc, 0, Height - H / 2, vbSrcCopy
                                        Case 2
                                            BitBlt hDc, Width / 2 - X, Round((Height - H) / 2), 1, Round(H / 2), frmConfig.picMiddle.hDc, 0, (Height - H) / 4, vbSrcCopy
                                            BitBlt hDc, Width / 2 + X, Round((Height - H) / 2), 1, Round(H / 2), frmConfig.picMiddle.hDc, 0, (Height - H) / 4, vbSrcCopy
                                            BitBlt hDc, Width / 2 - X, Round(Height / 2 - 1), 1, Round(H / 2), frmConfig.picMiddle.hDc, 0, Height / 2 + (Height - H) / 4, vbSrcCopy
                                            BitBlt hDc, Width / 2 + X, Round(Height / 2 - 1), 1, Round(H / 2), frmConfig.picMiddle.hDc, 0, Height / 2 + (Height - H) / 4, vbSrcCopy
                                    End Select
                                Case 2
                                    Select Case tSpec.FillMode
                                        Case 0 'Normal
                                            BitBlt hDc, Width - X, Round((Height - H) / 2), 1, Round(H / 2), frmConfig.picMiddle.hDc, 0, (Height - H) / 2, vbSrcCopy
                                            BitBlt hDc, Width - X, Round(Height / 2 - 1), 1, Round(H / 2), frmConfig.picMiddle.hDc, 0, (Height) / 2, vbSrcCopy
                                        Case 1
                                            BitBlt hDc, Width - X, Round((Height - H) / 2), 1, Round(H / 2), frmConfig.picMiddle.hDc, 0, 0, vbSrcCopy
                                            BitBlt hDc, Width - X, Round(Height / 2 - 1), 1, Round(H / 2), frmConfig.picMiddle.hDc, 0, Height - H / 2, vbSrcCopy
                                        Case 2
                                            BitBlt hDc, Width - X, Round((Height - H) / 2), 1, Round(H / 2), frmConfig.picMiddle.hDc, 0, (Height - H) / 4, vbSrcCopy
                                            BitBlt hDc, Width - X, Round(Height / 2 - 1), 1, Round(H / 2), frmConfig.picMiddle.hDc, 0, Height / 2 + (Height - H) / 4, vbSrcCopy
                                    End Select
                            End Select
                        Case 2
                            Select Case tSpec.PosX
                                Case 0
                                    Select Case tSpec.FillMode
                                        Case 0 'Normal
                                            BitBlt hDc, X, Round(Height - H), 1, Round(H), frmConfig.picBottom.hDc, 0, (Height - H), vbSrcCopy
                                        Case 1
                                            BitBlt hDc, X, Round(Height - H), 1, Round(H), frmConfig.picBottom.hDc, 0, 0, vbSrcCopy
                                        Case 2
                                            BitBlt hDc, X, Round(Height - H), 1, Round(H), frmConfig.picBottom.hDc, 0, (Height - H) / 2, vbSrcCopy
                                    End Select
                                Case 1
                                    Select Case tSpec.FillMode
                                        Case 0 'Normal
                                            BitBlt hDc, Width / 2 - X, Round(Height - H), 1, Round(H), frmConfig.picBottom.hDc, 0, (Height - H), vbSrcCopy
                                            BitBlt hDc, Width / 2 + X, Round(Height - H), 1, Round(H), frmConfig.picBottom.hDc, 0, (Height - H), vbSrcCopy
                                        Case 1
                                            BitBlt hDc, Width / 2 - X, Round(Height - H), 1, Round(H), frmConfig.picBottom.hDc, 0, 0, vbSrcCopy
                                            BitBlt hDc, Width / 2 + X, Round(Height - H), 1, Round(H), frmConfig.picBottom.hDc, 0, 0, vbSrcCopy
                                        Case 2
                                            BitBlt hDc, Width / 2 - X, Round(Height - H), 1, Round(H), frmConfig.picBottom.hDc, 0, (Height - H) / 2, vbSrcCopy
                                            BitBlt hDc, Width / 2 + X, Round(Height - H), 1, Round(H), frmConfig.picBottom.hDc, 0, (Height - H) / 2, vbSrcCopy
                                    End Select
                                Case 2
                                    Select Case tSpec.FillMode
                                        Case 0 'Normal
                                            BitBlt hDc, Width - X, Round(Height - H), 1, Round(H), frmConfig.picBottom.hDc, 0, (Height - H), vbSrcCopy
                                        Case 1
                                            BitBlt hDc, Width - X, Round(Height - H), 1, Round(H), frmConfig.picBottom.hDc, 0, 0, vbSrcCopy
                                        Case 2
                                            BitBlt hDc, Width - X, Round(Height - H), 1, Round(H), frmConfig.picBottom.hDc, 0, (Height - H) / 2, vbSrcCopy
                                    End Select
                            End Select
                    End Select
                Next i
            Case 2
                '******************Draw Led Bar********************
                For r = 0 To 248
                    Y = (Data(r) + Data(r + 1) + Data(r + 2) + Data(r + 3) + Data(r + 4) + Data(r + 5) + Data(r + 6) + Data(r + 7)) / 8 * tSpec.Zoom
                    If Y > Height Then Y = Height
                    '-------------------Bar------------------
                        Dim z As Integer
                        If tSpec.ShowBar Then
                            Select Case tSpec.PosY
                                Case 0 'Top
                                    Select Case tSpec.PosX
                                        Case 0 'Left
                                            For z = 0 To Y Step 4
                                                StretchBlt hDc, X, z, tSpec.BarWidth, 2, frmConfig.picTop.hDc, 0, z, tSpec.BarWidth, 1, vbSrcCopy
                                            Next z
                                        Case 1 'Middle
                                            For z = 0 To Y Step 4
                                                StretchBlt hDc, (Width / 2 - X), z, tSpec.BarWidth, 2, frmConfig.picTop.hDc, 0, z, tSpec.BarWidth, 1, vbSrcCopy
                                                StretchBlt hDc, (Width / 2 + X), z, tSpec.BarWidth, 2, frmConfig.picTop.hDc, 0, z, tSpec.BarWidth, 1, vbSrcCopy
                                            Next z
                                        Case 2 'Right
                                            For z = 0 To Y Step 4
                                                StretchBlt hDc, Width - X, z, tSpec.BarWidth, 2, frmConfig.picTop.hDc, 0, z, tSpec.BarWidth, 1, vbSrcCopy
                                            Next z
                                    End Select
                                Case 1 'Middle
                                    Select Case tSpec.PosX
                                        Case 0 'Left
                                            For z = Height / 2 To (Height / 2 - Y / 2) Step -4
                                                StretchBlt hDc, X, z, tSpec.BarWidth, 2, frmConfig.picMiddle.hDc, 0, z, tSpec.BarWidth, 1, vbSrcCopy
                                                StretchBlt hDc, X, Height + 1 - z, tSpec.BarWidth, 2, frmConfig.picMiddle.hDc, 0, (Height + 1 - z), tSpec.BarWidth, 1, vbSrcCopy
                                            Next z
                                        Case 1 'Middle
                                            For z = Height / 2 To (Height / 2 - Y / 2) Step -4
                                                StretchBlt hDc, Width / 2 - X, z, tSpec.BarWidth, 2, frmConfig.picMiddle.hDc, 0, z, tSpec.BarWidth, 1, vbSrcCopy
                                                StretchBlt hDc, Width / 2 - X, Height + 1 - z, tSpec.BarWidth, 2, frmConfig.picMiddle.hDc, 0, (Height + 1 - z), tSpec.BarWidth, 1, vbSrcCopy
                                                StretchBlt hDc, Width / 2 + X, z, tSpec.BarWidth, 2, frmConfig.picMiddle.hDc, 0, z, tSpec.BarWidth, 1, vbSrcCopy
                                                StretchBlt hDc, Width / 2 + X, Height + 1 - z, tSpec.BarWidth, 2, frmConfig.picMiddle.hDc, 0, (Height + 1 - z), tSpec.BarWidth, 1, vbSrcCopy
                                            Next z
                                        Case 2
                                            For z = Height / 2 To (Height / 2 - Y / 2) Step -4
                                                StretchBlt hDc, Width - X, z, tSpec.BarWidth, 2, frmConfig.picMiddle.hDc, 0, z, tSpec.BarWidth, 1, vbSrcCopy
                                                StretchBlt hDc, Width - X, Height + 1 - z, tSpec.BarWidth, 2, frmConfig.picMiddle.hDc, 0, (Height + 1 - z), tSpec.BarWidth, 1, vbSrcCopy
                                            Next z
                                    End Select
                                Case 2 'Bottom
                                    Select Case tSpec.PosX
                                        Case 0
                                            For z = Height To (Height - Y) Step -4
                                                StretchBlt hDc, X, z, tSpec.BarWidth, 2, frmConfig.picBottom.hDc, 0, z, tSpec.BarWidth, 1, vbSrcCopy
                                            Next z
                                        Case 1
                                            For z = Height To (Height - Y) Step -4
                                                StretchBlt hDc, Width / 2 - X, z, tSpec.BarWidth, 2, frmConfig.picBottom.hDc, 0, z, tSpec.BarWidth, 1, vbSrcCopy
                                                StretchBlt hDc, Width / 2 + X, z, tSpec.BarWidth, 2, frmConfig.picBottom.hDc, 0, z, tSpec.BarWidth, 1, vbSrcCopy
                                            Next z
                                        Case 2
                                            For z = Height To (Height - Y) Step -4
                                                StretchBlt hDc, Width - X, z, tSpec.BarWidth, 2, frmConfig.picBottom.hDc, 0, z, tSpec.BarWidth, 1, vbSrcCopy
                                            Next z
                                    End Select
                            End Select
                        End If
                        '----------------------Peak----------------------
                        If tSpec.PeakDraw = 0 Then
                            If bolDelay(r) = tSpec.PeakInteval Then
                                Data3(r) = Data3(r) - tSpec.PeakDrop / 10
                                bolStart(r) = True
                            End If
                            bolDelay(r) = bolDelay(r) + 1
                            If Data3(r) = Y Then
                                bolDelay(r) = 0
                                bolStart(r) = False
                            End If
                            If Data3(r) > Y And bolStart(r) Then
                                bolDelay(r) = tSpec.PeakInteval
                            End If
                            If Data3(r) = Y Then
                                bolDelay(r) = 0
                                bolStart(r) = False
                            End If
                        End If
                        If tSpec.PeakDraw = 1 Then
                            If bolDelay(r) = tSpec.PeakInteval Then
                                Data3(r) = Data3(r) + tSpec.PeakDrop / 10
                                bolStart(r) = True
                            End If
                            bolDelay(r) = bolDelay(r) + 1
                            If Data3(r) = Y Then
                                bolDelay(r) = 0
                                bolStart(r) = False
                            End If
                            If Data3(r) > Y And bolStart(r) Then
                                bolDelay(r) = tSpec.PeakInteval
                            End If
                            If Data3(r) >= Height Then Data3(r) = Y
                            If Data3(r) = Y Then
                                bolDelay(r) = 0
                                bolStart(r) = False
                            End If
                        End If
                        If tSpec.PeakDraw = 2 Then
                            If Data3(r) <= Height And Not (bolStart(r)) Then
                                Data3(r) = Data3(r) + 5 / 10
                                If Data3(r) = Height Then
                                    bolStart(r) = True
                                End If
                            End If
                            If bolStart(r) Then
                                Data3(r) = Data3(r) - 5 / 10
                                bolStart(r) = True
                            End If
                            If Data3(r) = Y Then
                                bolStart(r) = False
                            End If
                        End If
                        If Data3(r) < Y Then
                            Data3(r) = (Y)
                        End If
                        If tSpec.ShowPeak Then
                            Select Case tSpec.PosY
                                Case 0  'Top
                                    Select Case tSpec.PosX
                                        Case 0 'Left
                                            BitBlt hDc, X, Round(Data3(r)), tSpec.BarWidth, tSpec.PeakHeight, frmConfig.picPeak.hDc, 0, 0, vbSrcCopy
                                        Case 1 'Middle
                                            BitBlt hDc, Width / 2 - X, Round(Data3(r)), tSpec.BarWidth, tSpec.PeakHeight, frmConfig.picPeak.hDc, 0, 0, vbSrcCopy
                                            BitBlt hDc, Width / 2 + X, Round(Data3(r)), tSpec.BarWidth, tSpec.PeakHeight, frmConfig.picPeak.hDc, 0, 0, vbSrcCopy
                                        Case 2 'Right
                                            BitBlt hDc, Width - X, Round(Data3(r)), tSpec.BarWidth, tSpec.PeakHeight, frmConfig.picPeak.hDc, 0, 0, vbSrcCopy
                                    End Select
                                Case 1 'Middle
                                    Select Case tSpec.PosX
                                        Case 0 'Left
                                            BitBlt hDc, X, Round((Height - Data3(r)) / 2), tSpec.BarWidth, tSpec.PeakHeight, frmConfig.picPeak.hDc, 0, 0, vbSrcCopy
                                            BitBlt hDc, X, Round((Height + Data3(r)) / 2), tSpec.BarWidth, tSpec.PeakHeight, frmConfig.picPeak.hDc, 0, 0, vbSrcCopy
                                        Case 1 'Middle
                                            BitBlt hDc, Width / 2 - X, Round((Height - Data3(r)) / 2), tSpec.BarWidth, tSpec.PeakHeight, frmConfig.picPeak.hDc, 0, 0, vbSrcCopy
                                            BitBlt hDc, Width / 2 - X, Round((Height + Data3(r)) / 2), tSpec.BarWidth, tSpec.PeakHeight, frmConfig.picPeak.hDc, 0, 0, vbSrcCopy
                                            BitBlt hDc, Width / 2 + X, Round((Height - Data3(r)) / 2), tSpec.BarWidth, tSpec.PeakHeight, frmConfig.picPeak.hDc, 0, 0, vbSrcCopy
                                            BitBlt hDc, Width / 2 + X, Round((Height + Data3(r)) / 2), tSpec.BarWidth, tSpec.PeakHeight, frmConfig.picPeak.hDc, 0, 0, vbSrcCopy
                                        Case 2 'Right
                                            BitBlt hDc, Width - X, Round((Height - Data3(r)) / 2), tSpec.BarWidth, tSpec.PeakHeight, frmConfig.picPeak.hDc, 0, 0, vbSrcCopy
                                            BitBlt hDc, Width - X, Round((Height + Data3(r)) / 2), tSpec.BarWidth, tSpec.PeakHeight, frmConfig.picPeak.hDc, 0, 0, vbSrcCopy
                                    End Select
                                Case 2 'Bottom
                                    Select Case tSpec.PosX
                                        Case 0
                                            BitBlt hDc, X, Round(Height - Data3(r)), tSpec.BarWidth, tSpec.PeakHeight, frmConfig.picPeak.hDc, 0, 0, vbSrcCopy
                                        Case 1
                                            BitBlt hDc, Width / 2 - X, Round(Height - Data3(r)), tSpec.BarWidth, tSpec.PeakHeight, frmConfig.picPeak.hDc, 0, 0, vbSrcCopy
                                            BitBlt hDc, Width / 2 + X, Round(Height - Data3(r)), tSpec.BarWidth, tSpec.PeakHeight, frmConfig.picPeak.hDc, 0, 0, vbSrcCopy
                                        Case 2
                                            BitBlt hDc, Width - X, Round(Height - Data3(r)), tSpec.BarWidth, tSpec.PeakHeight, frmConfig.picPeak.hDc, 0, 0, vbSrcCopy
                                    End Select
                            End Select
                        End If
                        Select Case tSpec.PosX
                            Case 0, 2
                                    r = r + Int(256 / tSpec.BarNumber)
                            Case 1
                                    r = r + Int(128 / tSpec.BarNumber)
                        End Select
                        X = X + tSpec.BarWidth + tSpec.ScaleW
                Next r
        End Select
End Sub
Public Sub doConfig()
    frmConfig.Visible = True
End Sub
Public Function Sqrt(ByVal num As Double) As Double
    Sqrt = num ^ 0.5
End Function


Private Sub Class_Initialize()
    Dim strFileconfig As String
    strFileconfig = App.Path & "\M3P_vis.ini"

    tSpec.Zoom = ReadINI("visSpectrum", "Zoom", strFileconfig, 3)
    tSpec.DrawStyle = ReadINI("visSpectrum", "DrawStyle", strFileconfig, 0)
    tSpec.FillMode = ReadINI("visSpectrum", "FillMode", strFileconfig, 0)
    tSpec.PosX = ReadINI("visSpectrum", "PosX", strFileconfig, 0)
    tSpec.PosY = ReadINI("visSpectrum", "PosY", strFileconfig, 0)
    tSpec.ShowBar = ReadINI("visSpectrum", "ShowBar", strFileconfig, True)
    tSpec.BarNumber = ReadINI("visSpectrum", "BarNumber", strFileconfig, 32)
    tSpec.ScaleW = ReadINI("visSpectrum", "ScaleW", strFileconfig, 1)
    tSpec.LowColor = CLng(ReadINI("visSpectrum", "LowColor", strFileconfig, RGB(0, 255, 0)))
    tSpec.MidColor = CLng(ReadINI("visSpectrum", "MidColor", strFileconfig, RGB(255, 255, 0)))
    tSpec.HightColor = CLng(ReadINI("visSpectrum", "HightColor", strFileconfig, RGB(255, 0, 0)))
    tSpec.ShowPeak = ReadINI("visSpectrum", "ShowPeak", strFileconfig, True)
    tSpec.PeakHeight = ReadINI("visSpectrum", "PeakHeight", strFileconfig, 1)
    tSpec.PeakColor = CLng(ReadINI("visSpectrum", "PeakColor", strFileconfig, &HFFFFFF))
    tSpec.PeakGradient = ReadINI("visSpectrum", "PeakGradient", strFileconfig, False)
    tSpec.PeakInteval = ReadINI("visSpectrum", "PeakInteval", strFileconfig, 25)
    tSpec.PeakDrop = ReadINI("visSpectrum", "PeakDrop", strFileconfig, 4)
    tSpec.PeakDraw = ReadINI("visSpectrum", "PeakDraw", strFileconfig, 0)
    Load frmConfig
    frmConfig.Visible = False
End Sub

Private Sub Class_Terminate()
    Unload frmConfig
End Sub


