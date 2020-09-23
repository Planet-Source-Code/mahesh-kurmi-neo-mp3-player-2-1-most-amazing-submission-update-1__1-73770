Attribute VB_Name = "ModSpectrum"
Option Explicit

' some constants shared by frmVisualisation and clsDraw
' (actually only FFT_SAMPLES...)
Public boolRec_ON As Boolean

Public Const Pi As Single = 3.14159265358979
Public Const FFT_MAXAMPLITUDE As Double = 0.2
Public Const FFT_BANDLOWER As Double = 0.07
Public Const FFT_BANDSPACE As Long = 1
Public Const FFT_BANDWIDTH As Long = 5
Public Const FFT_STARTINDEX As Long = 1
Public Const FFT_SAMPLES As Long = 1024

Public Const DRW_BARXOFF As Long = 2
Public Const DRW_BARYOFF As Long = 1

Public DRW_BARWIDTH As Long
Public visDRW_BARWIDTH As Long
Public DRW_BARSPACE As Long
Public FFT_BANDS As Long
Public DRW_PEAKFALL As Long
Public COLOR_1 As Long
Public COLOR_2 As Long
Public COLOR_3 As Long
Public ColorBackGraph As Long

Public Menu_foreColor As Long
Public Menu_BackColor As Long
Public Menu_AciveColor As Long
Public Menu_GradStColor As Long
Public Menu_GradEndColor As Long
Public Menu_SepBackColor As Long

Public Type vbcolor
    Red As Single
    Green As Single
    Blue As Single
End Type


'//////FUNCTION FOR GETTING INTERMEDIATE COLOR
Public Function GetGradColor( _
       ByVal Max As Single, _
       ByVal Value As Single, _
       ByVal colstart As Long, _
       ByVal colmiddle As Long, _
       ByVal colend As Long _
       ) As Long

    Dim udtCol1 As vbcolor
    Dim udtCol2 As vbcolor
    Dim udtCol3 As vbcolor
    Dim udtColS As vbcolor
    Dim udtColE As vbcolor
    Dim udtDrw As vbcolor
    Dim X As Long

    udtCol1 = translate_color(colstart)
    udtCol2 = translate_color(colmiddle)
    udtCol3 = translate_color(colend)

    If Value < Max / 2 Then
        udtDrw = udtCol1
        udtColS = udtCol1
        udtColE = udtCol2
    Else
        Value = Value - (Max / 2)
        udtDrw = udtCol2
        udtColS = udtCol2
        udtColE = udtCol3
    End If

    Max = Max / 2

    With udtDrw    'INTERPOATE COLOR TO GET INTERMEDIATE GRADIENT COLOR
        .Red = .Red + (((udtColE.Red - udtColS.Red) / Max) * Value)
        .Green = .Green + (((udtColE.Green - udtColS.Green) / Max) * Value)
        .Blue = .Blue + (((udtColE.Blue - udtColS.Blue) / Max) * Value)
        GetGradColor = RGB(.Red, .Green, .Blue)
    End With
End Function


Public Function translate_color( _
       ByVal olecol As Long _
       ) As vbcolor

    With translate_color
        .Blue = (olecol \ &H10000) And &HFF
        .Green = (olecol \ &H100) And &HFF
        .Red = olecol And &HFF
    End With
End Function


' get the greatest absolute value in an array of samples
Private Function GetArrayMaxAbs( _
        intArray() As Integer, _
        Optional ByVal offStart As Long = 0, _
        Optional ByVal steps As Long = 1 _
        ) As Long

    Dim lngTemp As Long
    Dim lngMax As Long
    Dim i As Long

    For i = offStart To UBound(intArray) Step steps
        lngTemp = Abs(CLng(intArray(i)))
        If lngTemp > lngMax Then
            lngMax = lngTemp
        End If
    Next

    GetArrayMaxAbs = lngMax
End Function

'///// EQUALISE FREQUENCIES
Public Function Hanning( _
       ByVal X As Single, _
       ByVal Length As Long _
       ) As Single

    Hanning = 0.5 * (1 - Cos((2 * Pi * X) / Length))
End Function


