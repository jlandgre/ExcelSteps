'dictionary_cls.vb
'High-performance cross-platform dictionary class (Claude Sonnet generated providing a simple,
'cross-platform dictionary data type Mac and Windows Excel compatible)
'Version 9/19/25
Option Explicit

Private keys() As String
Private values() As Variant
Private itemCount As Long
Private capacity As Long
'-----------------------------------------------------------------------------------------------
' Initialize dictionary with default capacity
'
Public Sub Class_Initialize()
    capacity = 16
    itemCount = 0
    ReDim keys(0 To capacity - 1)
    ReDim values(0 To capacity - 1)
End Sub
'-----------------------------------------------------------------------------------------------
' Add or update key-value pair
' JDL 8/6/25
'
Public Sub Add(ByVal key As String, ByVal value As Variant)
    Dim index As Long
    
    ' Check if key already exists
    index = FindIndex(key)
    If index >= 0 Then
        ' Update existing value
        If IsObject(value) Then
            Set values(index) = value
        Else
            values(index) = value
        End If
    Else
        ' Add new key-value pair
        If itemCount >= capacity Then ExpandCapacity
        
        keys(itemCount) = key
        If IsObject(value) Then
            Set values(itemCount) = value
        Else
            values(itemCount) = value
        End If
        itemCount = itemCount + 1
    End If
End Sub
'-----------------------------------------------------------------------------------------------
' Remove key-value pair by key
' JDL 8/6/25
'
Public Sub Remove(ByVal key As String)
    Dim index As Long, i As Long
    
    index = FindIndex(key)
    If index >= 0 Then
        ' Shift remaining items left
        For i = index To itemCount - 2
            keys(i) = keys(i + 1)
            If IsObject(values(i + 1)) Then
                Set values(i) = values(i + 1)
            Else
                values(i) = values(i + 1)
            End If
        Next i
        itemCount = itemCount - 1
    End If
End Sub
'-----------------------------------------------------------------------------------------------
' Get value by key
' JDL 8/6/25
'
Public Function Item(ByVal key As String) As Variant
    Dim index As Long
    
    index = FindIndex(key)
    If index >= 0 Then
        If IsObject(values(index)) Then
            Set Item = values(index)
        Else
            Item = values(index)
        End If
    Else
        Item = Empty
    End If
End Function
'-----------------------------------------------------------------------------------------------
' Check if key exists
' JDL 8/6/25
'
Public Function Exists(ByVal key As String) As Boolean
    Exists = (FindIndex(key) >= 0)
End Function
'-----------------------------------------------------------------------------------------------
' Get number of items
' JDL 8/6/25
'
Public Property Get Size() As Long
    Size = itemCount
End Property
'-----------------------------------------------------------------------------------------------
' Get all keys as array
' JDL 8/6/25
'
Public Function GetKeys() As String()
    Dim result() As String
    Dim i As Long
    
    If itemCount = 0 Then
        GetKeys = result
        Exit Function
    End If
    
    ReDim result(0 To itemCount - 1)
    For i = 0 To itemCount - 1
        result(i) = keys(i)
    Next i
    GetKeys = result
End Function
'-----------------------------------------------------------------------------------------------
' Clear all items
' JDL 8/6/25
'
Public Sub Clear()
    itemCount = 0
    capacity = 16
    ReDim keys(0 To capacity - 1)
    ReDim values(0 To capacity - 1)
End Sub
'-----------------------------------------------------------------------------------------------
' Find index of key (returns -1 if not found)
' JDL 8/6/25
'
Private Function FindIndex(ByVal key As String) As Long
    Dim i As Long
    
    For i = 0 To itemCount - 1
        If keys(i) = key Then
            FindIndex = i
            Exit Function
        End If
    Next i
    FindIndex = -1
End Function
'-----------------------------------------------------------------------------------------------
' Expand capacity when needed
' JDL 8/6/25
'
Private Sub ExpandCapacity()
    capacity = capacity * 2
    ReDim Preserve keys(0 To capacity - 1)
    ReDim Preserve values(0 To capacity - 1)
End Sub

