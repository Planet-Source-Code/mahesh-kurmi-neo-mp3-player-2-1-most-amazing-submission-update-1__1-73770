VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Option Base 0

Private StrListArray() As String
Private StrEXListArray() As String
Private Tracklength() As String
Private Selected() As Boolean
Private List_is_Empty As Boolean

Public Sub Clear()
    ReDim StrListArray(0 To 0)
    ReDim StrEXListArray(0 To 0)
    ReDim Tracklength(0 To 0)
    ReDim Selected(0 To 0)
    List_is_Empty = True
End Sub
Public Sub RemoveItem(Index As Long)
    Const Err_Subscript_Out_Of_Range = 9
    Dim IndexID As Long
    If UBound(StrListArray) = 0 Then
        StrListArray(0) = ""
        StrEXListArray(0) = ""
        Tracklength(0) = ""
        Selected(0) = False
        List_is_Empty = True
    Else
        Select Case Index
        Case Is > ItemCount - 1, Is < 0
            err.Raise Err_Subscript_Out_Of_Range
            Exit Sub
        Case Else

            For IndexID = Index To UBound(StrListArray) - 1
                StrListArray(IndexID) = StrListArray(IndexID + 1)
                StrEXListArray(IndexID) = StrEXListArray(IndexID + 1)
                Tracklength(IndexID) = Tracklength(IndexID + 1)
                Selected(IndexID) = Selected(IndexID + 1)
            Next IndexID

            ReDim Preserve StrListArray(LBound(StrListArray) To UBound(StrListArray) - 1)
            ReDim Preserve StrEXListArray(LBound(StrEXListArray) To UBound(StrEXListArray) - 1)
            ReDim Preserve Tracklength(LBound(Tracklength) To UBound(Tracklength) - 1)
            ReDim Preserve Selected(LBound(Selected) To UBound(Selected) - 1)
        End Select
    End If
End Sub

Public Sub AddItem(TextLine As String, exTextLine As String, Optional EXSelection As String = " ", Optional bSelection As Boolean = False, Optional Index As Long = -1)
    Dim sFullpath As String, sFilename As String, sFileEx As String
    'If EXSelection = "" Then EXSelection = "00:00 "
    Dim i As Long

    If List_is_Empty = False Then
        ReDim Preserve StrListArray(LBound(StrListArray) To UBound(StrListArray) + 1)
        ReDim Preserve StrEXListArray(LBound(StrEXListArray) To UBound(StrEXListArray) + 1)
        ReDim Preserve Tracklength(LBound(Tracklength) To UBound(Tracklength) + 1)
        ReDim Preserve Selected(LBound(Selected) To UBound(Selected) + 1)

    End If
    StrEXListArray(UBound(StrListArray)) = exTextLine
    StrListArray(UBound(StrListArray)) = TextLine
    Tracklength(UBound(Tracklength)) = EXSelection
    Selected(UBound(Selected)) = bSelection
    List_is_Empty = False
    If Index <> -1 And Index < UBound(StrListArray) Then
        For i = UBound(StrListArray) To Index + 1 Step -1
            SwapItems i, i - 1
        Next
    End If
End Sub


Public Sub Sort(Optional iSortParam As Integer = 1)
'The fastets sort algorithm!
'iSortParam decides which parameter should be compared in sorting
'isortParam=3 sort by tracklength
'isortParam=2 sort by filename
'isortParam=1 or else sort by playlist entry
    Dim sVal1 As String, sVal2 As String

    Dim Row As Long
    Dim MaxRow As Long
    Dim MinRow As Long
    Dim Swtch As Long
    Dim Limit As Long
    Dim Offset As Long

    MaxRow = UBound(StrListArray)
    MinRow = LBound(StrListArray)
    Offset = MaxRow \ 2

    Do While Offset > 0
        Limit = MaxRow - Offset
        Do
            Swtch = False         ' Assume no switches at this offset.

            ' Compare elements and switch ones out of order:

            For Row = MinRow To Limit
                If iSortParam = 3 Then
                    sVal1 = LCase(Tracklength(Row))
                    sVal2 = LCase(Tracklength(Row + Offset))
                ElseIf iSortParam = 2 Then
                    sVal1 = LCase(StrEXListArray(Row))
                    sVal2 = LCase(StrEXListArray(Row + Offset))
                Else
                    sVal1 = LCase(StrListArray(Row))
                    sVal2 = LCase(StrListArray(Row + Offset))
                End If
                If sVal1 > sVal2 Then
                    Call SwapItems(Row, (Row + Offset))
                    Swtch = Row
                End If
            Next Row

            ' Sort on next pass only to where last switch was made:
            Limit = Swtch - Offset
        Loop While Swtch

        ' No switches at last offset, try one half as big:
        Offset = Offset \ 2
    Loop
End Sub

Public Property Get ItemCount() As Long
    If List_is_Empty Then
        ItemCount = 0
    Else
        ItemCount = UBound(StrListArray) + 1
    End If
End Property

Public Function Item(ByVal IndexID As Long) As String
    If IndexID > UBound(StrListArray) Then Item = "": Exit Function
    Item = StrListArray(IndexID)
End Function

Public Function exItem(ByVal IndexID As Long) As String
    If IndexID > UBound(StrListArray) Then exItem = "": Exit Function
    exItem = StrEXListArray(IndexID)
End Function
Public Function bSelected(ByVal IndexID As Long) As Boolean
    On Error Resume Next
    bSelected = Selected(IndexID)
End Function

Public Function ChangeItem(ByVal IndexID As Long, newdata As String)
    On Error Resume Next
    StrListArray(IndexID) = newdata
End Function

Public Function ChangeEXTracklength(ByVal IndexID As Long, newdata As String)
    On Error Resume Next
    Tracklength(IndexID) = newdata
End Function

Public Function ChangeSelection(ByVal IndexID As Long, newdata As Boolean)
    Selected(IndexID) = newdata
End Function

Public Function ChangeEXItem(ByVal IndexID As Long, newdata As String)
    StrEXListArray(IndexID) = newdata
End Function

Public Function exTracklength(ByVal IndexID As Long) As String
    exTracklength = Tracklength(IndexID)
End Function

Public Function SetTracklength(ByVal IndexID As Long, nEXSelection As String) As String
    Tracklength(IndexID) = nEXSelection
End Function

Private Sub Class_Initialize()
    ReDim StrListArray(0 To 0)
    ReDim StrEXListArray(0 To 0)
    ReDim Tracklength(0 To 0)
    ReDim Selected(0 To 0)
    List_is_Empty = True
End Sub

Private Sub strSwap(ByRef var1 As String, ByRef var2 As String)
    Dim X As String
    X = var1
    var1 = var2
    var2 = X
End Sub

Private Sub intSwap(ByRef var1 As Integer, ByRef var2 As Integer)
    Dim X As Integer
    X = var1
    var1 = var2
    var2 = X
End Sub


Public Sub SwapItems(index1 As Long, index2 As Long)
    If index1 < 0 Or index1 > UBound(StrListArray) Then Exit Sub
    If index2 < 0 Or index2 > UBound(StrListArray) Then Exit Sub
    strSwap StrListArray(index1), StrListArray(index2)
    strSwap StrEXListArray(index1), StrEXListArray(index2)
    strSwap Tracklength(index1), Tracklength(index2)
    Dim bTemp As Boolean
    bTemp = Selected(index1)
    Selected(index1) = Selected(index2)
    Selected(index2) = bTemp
End Sub

Public Sub Reverse()
    Dim i As Long
    For i = 0 To (UBound(StrListArray) + 1) / 2
        Call SwapItems(i, UBound(StrListArray) - i)
    Next
End Sub

