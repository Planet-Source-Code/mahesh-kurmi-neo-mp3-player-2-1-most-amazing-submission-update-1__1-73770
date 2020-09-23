Attribute VB_Name = "modPi"
' module containing plug-in loader
Option Explicit

Public Type PiType
    PiObject As String
    isConfigurable As Boolean
End Type

Public oPlugIn As Object

'function to read pi properties from file
Public Function getPlugInObject(PiFile As String) As PiType

    Dim Fsys As New FileSystemObject
    Dim Fin As TextStream, readLine As String
    Dim p As PiType

    Set Fin = Fsys.OpenTextFile(PiFile)

    While Not Fin.AtEndOfStream
        readLine = Fin.readLine
        ' found an object
        If InStr(1, readLine, "OBJECT") Then
            p.PiObject = Trim(Split(readLine, ":")(1))
        End If
        ' found a config setting
        If InStr(1, readLine, "CONFIG") Then
            p.isConfigurable = CBool(Split(readLine, ":")(1))
        End If

    Wend
    Fin.Close
    Set Fin = Nothing
    Set Fsys = Nothing

    getPlugInObject = p

End Function


Public Sub importXPIList(File As String)
    Dim inStream As TextStream

    Dim A As Boolean, str As String
    Dim i As Integer, ext As String, c As Integer

    Let A = True
    If Trim(File) = "" Or Dir(File) = "" Then GoTo e
    Set inStream = Fsys.OpenTextFile(Trim(File), ForReading, False)


    If Not StrComp(Replace(inStream.Read(5), "<?", " "), "Soda") Then
        Dim e As ErrStruct
        e.errNum = 10
        e.errShortDesc = "This does not appear to be a FireAMP! Orange Soda list"
        e.errLongDesc = "The file recently opened did not have the FireAMP! Orange Soda header in it. The File is either corrupt or invalid"
        logError e
        Exit Sub
    End If

    inStream.SkipLine    ' Skip header
    inStream.SkipLine    ' Skip comment
    inStream.SkipLine    ' Skip main tag
    i = frmPopUp.mnuObjectName.Count - 1

    While inStream.AtEndOfStream = False

        str = inStream.readLine
        If str = "</list>" Then GoTo JMP

        If A = True Then
            On Error Resume Next
            Load frmPopUp.mnuObjectName(i)

            i = i + 1
            frmPopUp.mnuObjectName(i - 1).Caption = Split(parseString(str, 9, 9), ",")(1)    ' load name
            frmPopUp.mnuObjectName(i - 1).Tag = Split(parseString(str, 9, 9), ",")(0)    ' load object
        Else
            frmPopUp.mnuObjectName(i - 1).Tag = frmPopUp.mnuObjectName(i - 1).Tag & "," & Val(parseString(str, 8, 8))
        End If
        A = Not A
    Wend

JMP:
    For i = 0 To frmPopUp.mnuObjectName.Count - 1
        frmPopUp.mnuObjectName(i).Checked = False
    Next i
    Set inStream = Nothing    ' destroy object
e:
End Sub

