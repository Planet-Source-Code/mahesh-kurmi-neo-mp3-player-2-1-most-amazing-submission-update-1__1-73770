Attribute VB_Name = "modPi"
' module containing plug-in loader
Public Type PiType
 PiObject As String
 isConfigurable As Boolean
End Type

Public tPlugin As PiType
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
