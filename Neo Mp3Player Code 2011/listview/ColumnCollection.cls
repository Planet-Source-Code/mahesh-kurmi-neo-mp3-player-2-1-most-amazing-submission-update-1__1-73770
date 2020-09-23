VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ColumnCollection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_clsItems() As Column
Private m_lngColCount As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (lpDest As Any, lpSource As Any, ByVal cBytes As Long)
Private Declare Sub ZeroMemory Lib "kernel32" Alias "RtlZeroMemory" (lpDest As Any, ByVal cBytes As Long)


Public Property Get Count() As Long
    Count = m_lngColCount
End Property

Public Property Get Item(Index) As Column
    Dim i As Long
    Dim strKey As String

    If IsNumeric(Index) Then
        Set Item = m_clsItems(CLng(Index))
    Else
        strKey = Index
        For i = 0 To m_lngColCount - 1
            If StrComp(m_clsItems(i).Key, strKey, vbTextCompare) = 0 Then
                Set Item = m_clsItems(i)
                Exit Property
            End If
        Next

        err.Raise 9
    End If
End Property

Public Sub Remove(Index)
    Dim i As Long
    Dim strKey As String

    If IsNumeric(Index) Then
        i = CLng(Index)
    Else
        strKey = Index
        For i = 0 To m_lngColCount - 1
            If StrComp(m_clsItems(i).Key, strKey, vbTextCompare) = 0 Then
                Exit For
            End If
        Next
    End If

    If i > m_lngColCount - 1 Then
        err.Raise 9
    End If

    Set m_clsItems(i) = Nothing
    CopyMemory m_clsItems(i), m_clsItems(i + 1), 4 * (m_lngColCount - (i + 1))
    ZeroMemory m_clsItems(m_lngColCount - 1), 4
    m_lngColCount = m_lngColCount - 1
End Sub

Public Function Add(Optional ByVal Key As String = "") As Column
    Dim i As Long

    For i = 0 To m_lngColCount - 1
        If StrComp(m_clsItems(i).Key, Key, vbTextCompare) = 0 Then
            err.Raise 300, , "Objekt nicht eindeutig in Aufzählung!"
        End If
    Next

    ReDim Preserve m_clsItems(m_lngColCount) As Column
    Set m_clsItems(m_lngColCount) = New Column
    m_clsItems(m_lngColCount).Key = Key
    Set Add = m_clsItems(m_lngColCount)
    m_lngColCount = m_lngColCount + 1
End Function
