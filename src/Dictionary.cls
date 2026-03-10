VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Dictionary"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Version 1/29/26
'High-performance cross-platform dictionary class (Claude Sonnet generated to have
'cross-platform dictionary accessible in client-facing project)

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

'-----------------------------------------------------------------------------------------------
'-----------------------------------------------------------------------------------------------
' ParseStringToDictProcedure and related methods
'-----------------------------------------------------------------------------------------------
'-----------------------------------------------------------------------------------------------
'-----------------------------------------------------------------------------------------------
' Parse JSON-like string and add key-value pairs to Dictionary
' JDL 3/10/26
'
Public Function ParseStringToDictProcedure(ByVal jsonStr As String) As Boolean
    Dim innerStr As String, pairs As Variant, i As Long
    Dim key As String, valParsed As Variant
    Dim lb As Long, ub As Long

    If Not ValidateAndStripBraces(jsonStr, innerStr) Then Exit Function
    If Not SplitIntoPairs(innerStr, pairs) Then Exit Function

    If IsArray(pairs) Then
        On Error Resume Next
        lb = LBound(pairs)
        ub = UBound(pairs)
        If Err.Number <> 0 Then
            Err.Clear
            On Error GoTo 0
            ParseStringToDictProcedure = True
            Exit Function
        End If
        On Error GoTo 0

        For i = lb To ub
            If Not ParsePair(CStr(pairs(i)), key, valParsed) Then Exit Function
            If Not AddParsedValue(key, valParsed) Then Exit Function
        Next i
    End If

    ParseStringToDictProcedure = True
End Function
'-----------------------------------------------------------------------------------------------
' Validate outer braces and return inner string
' JDL 3/10/26
'
Public Function ValidateAndStripBraces(ByVal jsonStr As String, ByRef innerStr As String) As Boolean
    Dim sTrim As String

    sTrim = Trim$(jsonStr)
    If Len(sTrim) < 2 Then Exit Function
    If Left$(sTrim, 1) <> "{" Then Exit Function
    If Right$(sTrim, 1) <> "}" Then Exit Function

    If Len(sTrim) = 2 Then
        innerStr = ""
    Else
        innerStr = Mid$(sTrim, 2, Len(sTrim) - 2)
    End If

    ValidateAndStripBraces = True
End Function
'-----------------------------------------------------------------------------------------------
' Split inner string into key:value tokens while respecting quoted strings
' JDL 3/10/26
'
Public Function SplitIntoPairs(ByVal innerStr As String, ByRef pairs As Variant) As Boolean
    Dim colPairs As Collection
    Dim i As Long, ch As String, cur As String
    Dim inQuotes As Boolean

    If Len(innerStr) = 0 Then
        pairs = Array()
        SplitIntoPairs = True
        Exit Function
    End If

    Set colPairs = New Collection
    cur = ""

    For i = 1 To Len(innerStr)
        ch = Mid$(innerStr, i, 1)

        If ch = Chr$(34) Then
            inQuotes = Not inQuotes
            cur = cur & ch
        ElseIf ch = "," And Not inQuotes Then
            cur = Trim$(cur)
            If Len(cur) = 0 Then Exit Function
            colPairs.Add cur
            cur = ""
        Else
            cur = cur & ch
        End If
    Next i

    If inQuotes Then Exit Function

    cur = Trim$(cur)
    If Len(cur) = 0 Then Exit Function
    colPairs.Add cur

    ReDim pairs(0 To colPairs.Count - 1)
    For i = 1 To colPairs.Count
        pairs(i - 1) = colPairs(i)
    Next i

    SplitIntoPairs = True
End Function
'-----------------------------------------------------------------------------------------------
' Parse single key:value token into key and typed value
' JDL 3/10/26
'
Public Function ParsePair(ByVal pairStr As String, ByRef key As String, ByRef valParsed As Variant) As Boolean
    Dim i As Long, ch As String, inQuotes As Boolean
    Dim colonPos As Long, keyRaw As String, valRaw As String

    For i = 1 To Len(pairStr)
        ch = Mid$(pairStr, i, 1)
        If ch = Chr$(34) Then
            inQuotes = Not inQuotes
        ElseIf ch = ":" And Not inQuotes Then
            colonPos = i
            Exit For
        End If
    Next i

    If colonPos = 0 Then Exit Function

    keyRaw = Trim$(Left$(pairStr, colonPos - 1))
    valRaw = Trim$(Mid$(pairStr, colonPos + 1))
    If Len(keyRaw) = 0 Then Exit Function

    If Left$(keyRaw, 1) = Chr$(34) And Right$(keyRaw, 1) = Chr$(34) Then
        If Len(keyRaw) < 2 Then Exit Function
        key = Mid$(keyRaw, 2, Len(keyRaw) - 2)
    Else
        key = keyRaw
    End If

    If Len(key) = 0 Then Exit Function
    If Not DetectValueType(valRaw, valParsed) Then Exit Function

    ParsePair = True
End Function
'-----------------------------------------------------------------------------------------------
' Detect String, Boolean, Double or Empty from value text
' JDL 3/10/26
'
Public Function DetectValueType(ByVal valueStr As String, ByRef valParsed As Variant) As Boolean
    Dim sTrim As String

    sTrim = Trim$(valueStr)

    If Len(sTrim) = 0 Then
        valParsed = Empty
        DetectValueType = True
        Exit Function
    End If

    If Len(sTrim) >= 2 And Left$(sTrim, 1) = Chr$(34) And Right$(sTrim, 1) = Chr$(34) Then
        valParsed = Mid$(sTrim, 2, Len(sTrim) - 2)
        DetectValueType = True
        Exit Function
    End If

    Select Case UCase$(sTrim)
        Case "TRUE"
            valParsed = True
            DetectValueType = True
            Exit Function
        Case "FALSE"
            valParsed = False
            DetectValueType = True
            Exit Function
    End Select

    If IsNumeric(sTrim) Then
        valParsed = CDbl(sTrim)
        DetectValueType = True
        Exit Function
    End If
End Function
'-----------------------------------------------------------------------------------------------
' Add parsed key/value to Dictionary using Add
' JDL 3/10/26
'
Public Function AddParsedValue(ByVal key As String, ByVal valParsed As Variant) As Boolean
    On Error GoTo ErrorExit

    Add key, valParsed
    AddParsedValue = True
    Exit Function

ErrorExit:
    AddParsedValue = False
End Function

