VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Dictionary"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Version 3/11/26
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
    SetErrs ParseStringToDictProcedure: If errs.IsHandle Then On Error GoTo ErrorExit
    Dim innerStr As String, pairs As Variant, i As Long
    Dim key As String, valParsed As Variant
    Dim lb As Long, ub As Long

    If Not ValidateAndStripBraces(jsonStr, innerStr) Then GoTo ErrorExit
    If Not SplitIntoPairs(innerStr, pairs) Then GoTo ErrorExit

    If IsArray(pairs) Then
        On Error Resume Next
        lb = LBound(pairs)
        ub = UBound(pairs)
        If err.Number <> 0 Then
            err.Clear
            On Error GoTo 0
            Exit Function
        End If
        On Error GoTo 0

        For i = lb To ub
            If Not ParsePair(CStr(pairs(i)), key, valParsed) Then GoTo ErrorExit
            If Not AddParsedValue(key, valParsed) Then GoTo ErrorExit
        Next i
    End If

    Exit Function

ErrorExit:
    errs.RecordErr "ParseStringToDictProcedure", ParseStringToDictProcedure
End Function
'-----------------------------------------------------------------------------------------------
' Validate outer braces and return inner string
' JDL 3/10/26
'
Public Function ValidateAndStripBraces(ByVal jsonStr As String, ByRef innerStr As String) As Boolean
    SetErrs ValidateAndStripBraces: If errs.IsHandle Then On Error GoTo ErrorExit
    Dim sTrim As String

    sTrim = Trim$(jsonStr)

    ' Check for missing or misplaced braces
    If errs.IsFail(Left$(sTrim, 1) <> "{", 1) Then GoTo ErrorExit
    If errs.IsFail(Right$(sTrim, 1) <> "}", 2) Then GoTo ErrorExit

    If Len(sTrim) = 2 Then
        innerStr = ""
    Else
        innerStr = Mid$(sTrim, 2, Len(sTrim) - 2)
    End If

    Exit Function

ErrorExit:
    errs.RecordErr "ValidateAndStripBraces", ValidateAndStripBraces
End Function
'-----------------------------------------------------------------------------------------------
' Split inner string into key:value tokens while respecting quoted strings
' JDL 3/10/26
'
Public Function SplitIntoPairs(ByVal innerStr As String, ByRef pairs As Variant) As Boolean
    SetErrs SplitIntoPairs: If errs.IsHandle Then On Error GoTo ErrorExit
    Dim colPairs As Collection
    Dim i As Long, ch As String, cur As String
    Dim inQuotes As Boolean

    If Len(innerStr) = 0 Then
        pairs = Array()
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
            
            ' Missing key:value pair between commas (for example, double commas).
            If errs.IsFail(Len(cur) = 0, 1, innerStr) Then GoTo ErrorExit
            colPairs.Add cur
            cur = ""
        Else
            cur = cur & ch
        End If
    Next i

    ' Unterminated quote: pair splitting by commas may be incorrect.
    If errs.IsFail(inQuotes, 2, innerStr) Then GoTo ErrorExit

    cur = Trim$(cur)
    ' Blank final token: string ended with a comma (missing last key:value pair).
    If errs.IsFail(Len(cur) = 0, 3, innerStr) Then GoTo ErrorExit
    colPairs.Add cur

    ReDim pairs(0 To colPairs.Count - 1)
    For i = 1 To colPairs.Count
        pairs(i - 1) = colPairs(i)
    Next i

    Exit Function

ErrorExit:
    errs.RecordErr "SplitIntoPairs", SplitIntoPairs
End Function
'-----------------------------------------------------------------------------------------------
' Parse single key:value token into key and typed value
' JDL 3/10/26
'
Public Function ParsePair(ByVal pairStr As String, ByRef key As String, ByRef valParsed As Variant) As Boolean
    SetErrs ParsePair: If errs.IsHandle Then On Error GoTo ErrorExit
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

    ' Missing colon/invalid key:value pair
    If errs.IsFail(colonPos = 0, 1, pairStr) Then GoTo ErrorExit

    keyRaw = Trim$(Left$(pairStr, colonPos - 1))
    valRaw = Trim$(Mid$(pairStr, colonPos + 1))

    ' key:value pair starts with a colon
    If errs.IsFail(Len(keyRaw) = 0, 2, pairStr) Then GoTo ErrorExit

    If Left$(keyRaw, 1) = Chr$(34) And Right$(keyRaw, 1) = Chr$(34) Then

        ' Quoted key must include opening and closing quote characters
        If errs.IsFail(Len(keyRaw) < 2, 3, pairStr) Then GoTo ErrorExit
        key = Mid$(keyRaw, 2, Len(keyRaw) - 2)
    Else
        key = keyRaw
    End If

    ' Key cannot be blank after quote removal
    If errs.IsFail(Len(key) = 0, 4, pairStr) Then GoTo ErrorExit
    If Not DetectValueType(valRaw, valParsed) Then GoTo ErrorExit

    Exit Function

ErrorExit:
    errs.RecordErr "ParsePair", ParsePair
End Function
'-----------------------------------------------------------------------------------------------
' Detect String, Boolean, Double or Empty from value text
' JDL 3/10/26
'
Public Function DetectValueType(ByVal valueStr As String, ByRef valParsed As Variant) As Boolean
    SetErrs DetectValueType: If errs.IsHandle Then On Error GoTo ErrorExit
    Dim sTrim As String

    sTrim = Trim$(valueStr)

    If Len(sTrim) = 0 Then
        valParsed = Empty
        Exit Function
    End If

    ' valueStr is a quoted string
    If Len(sTrim) >= 2 And Left$(sTrim, 1) = Chr$(34) And Right$(sTrim, 1) = Chr$(34) Then
        valParsed = Mid$(sTrim, 2, Len(sTrim) - 2)
        Exit Function
    End If

    ' Boolean value (case-insensitive)
    Select Case UCase$(sTrim)
        Case "TRUE"
            valParsed = True
            Exit Function
        Case "FALSE"
            valParsed = False
            Exit Function
    End Select

    ' Numeric value
    If IsNumeric(sTrim) Then
        valParsed = CDbl(sTrim)
        Exit Function
    End If

    ' valueStr is invalid: not a recognized type (string, boolean, double or empty)
    If errs.IsFail(True, 1) Then GoTo ErrorExit
    Exit Function

ErrorExit:
    errs.RecordErr "DetectValueType", DetectValueType
End Function
'-----------------------------------------------------------------------------------------------
' Add parsed key/value to Dictionary using Add
' JDL 3/10/26
'
Public Function AddParsedValue(ByVal key As String, ByVal valParsed As Variant) As Boolean
    SetErrs AddParsedValue: If errs.IsHandle Then On Error GoTo ErrorExit

    Add key, valParsed
    Exit Function

ErrorExit:
    errs.RecordErr "AddParsedValue", AddParsedValue
End Function


