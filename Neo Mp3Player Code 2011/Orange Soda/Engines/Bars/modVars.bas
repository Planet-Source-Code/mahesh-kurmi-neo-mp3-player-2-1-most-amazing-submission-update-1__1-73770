Attribute VB_Name = "modVars"
Public BarRed As Long
Public BarGreen As Long
Public BarBlue As Long

Public PeakRed As Long
Public PeakGreen As Long
Public PeakBlue As Long
Option Explicit
'XPTheme
Public Declare Function InitCommonControls Lib "comctl32.dll" () As Long

Private Declare Function StretchBlt Lib "gdi32" (ByVal hDc As Long, _
                                        ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, _
                                        ByVal nHeight As Long, ByVal hSrcDC As Long, _
                                        ByVal xSrc As Long, ByVal ySrc As Long, _
                                        ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, _
                                        ByVal dwRop As Long) As Long

Private Declare Function StretchDIBits Lib "gdi32" (ByVal hDc As Long, _
                                        ByVal X As Long, ByVal Y As Long, ByVal dx As Long, _
                                        ByVal dy As Long, ByVal SrcX As Long, ByVal SrcY As Long, _
                                        ByVal wSrcWidth As Long, ByVal wSrcHeight As Long, _
                                        lpBits As Any, lpBitsInfo As Any, ByVal wUsage As Long, _
                                        ByVal dwRop As Long) As Long
Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" _
                                        (ByVal lpApplicationName As String, ByVal lpKeyName As String, _
                                        ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long

Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" _
                                        (ByVal lpApplicationName As String, ByVal lpKeyName As Any, _
                                        ByVal lpString As Any, ByVal lplFileName As String) As Long
Type Spectrum
    Zoom As Double
    DrawStyle As Integer 'Bar / Line
    FillMode As Integer 'Normal / Fire / col
    HightColor As Long
    MidColor As Long
    LowColor As Long
    ShowBar As Boolean
    BarNumber As Integer
    BarWidth As Long
    ScaleW As Integer
    ShowPeak As Boolean
    PeakColor As Long
    PeakDrop As Integer
    PeakGradient As Boolean
    PeakFaster As Boolean
    PeakHeight As Integer
    PeakInteval As Long
    PeakDraw As Integer
    PosX As Integer
    PosY As Integer
End Type
Public tSpec As Spectrum

Const DIB_RGB_ColS      As Long = 0

Private Type BITMAPINFOHEADER    '40 bytes
        biSize As Long
        biWidth As Long
        biHeight As Long
        biPlanes As Integer
        biBitCount As Integer
        biCompression As Long
        biSizeImage As Long
        biXPelsPerMeter As Long
        biYPelsPerMeter As Long
        biClrUsed As Long
        biClrImportant As Long
End Type

Private Type RGBQUAD
        rgbBlue As Byte
        rgbGreen As Byte
        rgbRed As Byte
        rgbReserved As Byte
End Type

Private Type BITMAPINFO
        bmiHeader As BITMAPINFOHEADER
        bmiColors(255) As RGBQUAD
End Type

Public Enum GradientDirectionEnum
    [Fill_None] = 0
    [Fill_Horizontal] = 1
    [Fill_HorizontalMiddleOut] = 2
    [Fill_Vertical] = 3
    [Fill_VerticalMiddleOut] = 4
    [Fill_DownwardDiagonal] = 5
    [Fill_UpwardDiagonal] = 6
End Enum

Public Function ReadINI(strSection As String, strKey As String, strFileINI As String, Optional strDefault As Variant) As String
    On Error GoTo beep
        Dim StrTemp As String * 255
        GetPrivateProfileString strSection, strKey, vbNull, StrTemp, Len(StrTemp), strFileINI
        
        ReadINI = Mid(StrTemp, 1, InStr(1, StrTemp, vbNullChar) - 1)
        Exit Function
beep:
        ReadINI = strDefault
End Function
Public Function WriteINI(strSection As String, strKey As String, KeyValue As Variant, strFileINI As String) As Boolean
    Dim ret As Long
        ret = WritePrivateProfileString(strSection, strKey, CStr(KeyValue), strFileINI)
        If ret = 0 Then
            WriteINI = True
        Else
            WriteINI = False
        End If
End Function



Private Sub DIBGradient(ByVal hDc As Long, _
                         ByVal X As Long, _
                         ByVal Y As Long, _
                         ByVal Width As Long, _
                         ByVal Height As Long, _
                         ByVal Col1 As Long, _
                         ByVal Col2 As Long, _
                         ByVal GradientDirection As GradientDirectionEnum)

  Dim uBIH    As BITMAPINFOHEADER
  Dim lBits() As Long
  Dim lGrad() As Long
  
  Dim R1      As Long
  Dim G1      As Long
  Dim B1      As Long
  Dim R2      As Long
  Dim G2      As Long
  Dim B2      As Long
  Dim dR      As Long
  Dim dG      As Long
  Dim dB      As Long
  
  Dim Scan    As Long
  Dim i       As Long
  Dim iEnd    As Long
  Dim iOffset As Long
  Dim J       As Long
  Dim jEnd    As Long
  Dim iGrad   As Long
  
    '-- A minor check
    If (Width < 1 Or Height < 1) Then Exit Sub
    
    '-- Decompose Cols
    Col1 = Col1 And &HFFFFFF
    R1 = Col1 Mod &H100&
    Col1 = Col1 \ &H100&
    G1 = Col1 Mod &H100&
    Col1 = Col1 \ &H100&
    B1 = Col1 Mod &H100&
    Col2 = Col2 And &HFFFFFF
    R2 = Col2 Mod &H100&
    Col2 = Col2 \ &H100&
    G2 = Col2 Mod &H100&
    Col2 = Col2 \ &H100&
    B2 = Col2 Mod &H100&
    
    '-- Get Col distances
    dR = R2 - R1
    dG = G2 - G1
    dB = B2 - B1
    
    '-- Size gradient-Cols array
    Select Case GradientDirection
        Case [Fill_Horizontal]
            ReDim lGrad(0 To Width - 1)
        Case [Fill_Vertical]
            ReDim lGrad(0 To Height - 1)
        Case Else
            ReDim lGrad(0 To Width + Height - 2)
    End Select
    
    '-- Calculate gradient-Cols
    iEnd = UBound(lGrad())
    If (iEnd = 0) Then
        '-- Special case (1-pixel wide gradient)
        lGrad(0) = (B1 \ 2 + B2 \ 2) + 256 * (G1 \ 2 + G2 \ 2) + 65536 * (R1 \ 2 + R2 \ 2)
      Else
        For i = 0 To iEnd
            lGrad(i) = B1 + (dB * i) \ iEnd + 256 * (G1 + (dG * i) \ iEnd) + 65536 * (R1 + (dR * i) \ iEnd)
        Next i
    End If
    
    '-- Size DIB array
    ReDim lBits(Width * Height - 1) As Long
    iEnd = Width - 1
    jEnd = Height - 1
    Scan = Width
    
    '-- Render gradient DIB
    Select Case GradientDirection
        
        Case [Fill_Horizontal]
        
            For J = 0 To jEnd
                For i = iOffset To iEnd + iOffset
                    lBits(i) = lGrad(i - iOffset)
                Next i
                iOffset = iOffset + Scan
            Next J
        
        Case [Fill_Vertical]
        
            For J = jEnd To 0 Step -1
                For i = iOffset To iEnd + iOffset
                    lBits(i) = lGrad(J)
                Next i
                iOffset = iOffset + Scan
            Next J
            
        Case [Fill_DownwardDiagonal]
            
            iOffset = jEnd * Scan
            For J = 1 To jEnd + 1
                For i = iOffset To iEnd + iOffset
                    lBits(i) = lGrad(iGrad)
                    iGrad = iGrad + 1
                Next i
                iOffset = iOffset - Scan
                iGrad = J
            Next J
            
        Case [Fill_UpwardDiagonal]
            
            iOffset = 0
            For J = 1 To jEnd + 1
                For i = iOffset To iEnd + iOffset
                    lBits(i) = lGrad(iGrad)
                    iGrad = iGrad + 1
                Next i
                iOffset = iOffset + Scan
                iGrad = J
            Next J
    End Select
    
    '-- Define DIB header
    With uBIH
        .biSize = 40
        .biPlanes = 1
        .biBitCount = 32
        .biWidth = Width
        .biHeight = Height
    End With
    
    '-- Paint it!
    Call StretchDIBits(hDc, X, Y, Width, Height, 0, 0, Width, Height, lBits(0), uBIH, DIB_RGB_ColS, vbSrcCopy)

End Sub
Public Sub FillGradient(ByVal hDc As Long, _
                         ByVal X As Long, _
                         ByVal Y As Long, _
                         ByVal Width As Long, _
                         ByVal Height As Long, _
                         ByVal Col1 As Long, _
                         ByVal Col2 As Long, _
                         ByVal GradientDirection As GradientDirectionEnum, _
                         Optional Right2Left As Boolean = True)
                         
    Dim tmpCol  As Long
  
    ' Exit if needed
    If GradientDirection = Fill_None Then Exit Sub
    
    ' Right-To-Left
    If Right2Left Then
        tmpCol = Col1
        Col1 = Col2
        Col2 = tmpCol
    End If
    
    Select Case GradientDirection
        Case Fill_HorizontalMiddleOut
            DIBGradient hDc, X, Y, Width / 2, Height, Col1, Col2, Fill_Horizontal
            DIBGradient hDc, X + Width / 2 - 1, Y, Width / 2, Height, Col2, Col1, Fill_Horizontal

        Case Fill_VerticalMiddleOut
            DIBGradient hDc, X, Y, Width, Height / 2, Col1, Col2, Fill_Vertical
            DIBGradient hDc, X, Y + Height / 2 - 1, Width, Height / 2, Col2, Col1, Fill_Vertical

        Case Else
            DIBGradient hDc, X, Y, Width, Height, Col1, Col2, GradientDirection
    End Select
    
End Sub




