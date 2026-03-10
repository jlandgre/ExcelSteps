VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Tests"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Version '2/13/25 Mac compatibility
'Version of 7/18/23 - add testing utility functions TestAryVals etc.
'Utility subs/functions used: AddSheet, SheetExists, ColWidthAutoFit, ShadeYellow
Option Explicit
Public wkbk As Workbook 'Testing workbook
Public wkbkTest As Workbook 'Workbook with results being tested (may be same as wkbk)
Public sFileTest As String 'Wkbktest file reference
Public shtT As String
Public valExpected As Variant
Public valTest As Variant
Public Result As Boolean
Public Name As String
'-----------------------------------------------------------------------------------------------
'Purpose:   Initialize sheet for reporting results (clears previous results)
'Inputs: wkbk [Workbook] contains tests sheet
'        shtT [String] tests sheet name for reporting test results
'
'Created:   11/10/21 JDL
'
Sub InitTestsSheet(wkbk, shtT)
    Dim c As Range
    
    'Add tests sheet if it does not exist
    If Not SheetExists(wkbk, shtT) Then AddSheet wkbk, shtT, wkbk.Sheets(wkbk.Sheets.Count).Name

    'Clear the results columns and refresh header strings
    With wkbk.Sheets(shtT)
        Range(.Columns(3), .Columns(4)).Clear
        Range(.Cells(1, 1), .Cells(1, 4)) = Split("Test Name,valExpected,valActual,TestResult", ",")
        .Columns(2).Interior.Pattern = xlNone
        ColWidthAutofit Range(.Cells(1, 1), .Cells(1, 4))
    End With
End Sub
'-----------------------------------------------------------------------------------------------
'Purpose:   Populate clsT properties
'Inputs:    clsT [Tests Class]
'           wkbk [Workbook] contains tests sheet
'           shtT [String] tests sheet name for reporting test results
'           sName [String] test name
'
'Created:   11/10/21 JDL    Modified: 6/7/22
'                                     7/18/23 default wkbktest to be wkbk if not specified
'
Sub Populate(clsT, wkbk, Optional shtT, Optional sName, Optional wkbkTest, Optional sFileTest)
    With clsT
        Set .wkbk = wkbk
        If Not IsMissing(wkbkTest) Then
            Set .wkbkTest = wkbkTest
        Else
            Set .wkbkTest = .wkbk
        End If
        If Not IsMissing(sFileTest) Then .sFileTest = sFileTest
        If Not IsMissing(shtT) Then .shtT = shtT
        If Not IsMissing(sName) Then .Name = sName
        .valExpected = Expected(clsT)
    End With
End Sub
'-----------------------------------------------------------------------------------------------
'Purpose:   Write test result to appropriate row in tests sheet
'
'Created:   11/10/21 JDL      Modified: 11/22/21
'
Sub Update(clsT)
    Dim c As Range, rng As Range
    With clsT
        If Not SheetExists(.wkbk, .shtT) Then InitTestsSheet .wkbk, .shtT
    End With
        
    'Find the existing test row or add a new if  not found
    With clsT.wkbk.Sheets(shtT)
        Set c = .Columns(1).Find(clsT.Name, LookAt:=xlWhole)
        If c Is Nothing Then
            Set c = .Cells(.Rows.Count, 1).End(xlUp).Offset(1, 0)
            c.Value = clsT.Name
        End If
    End With
    
    'Write the result
    With clsT
        .Result = False
    
        'Fail if expected value not populated
        If IsEmpty(.valExpected) Then
            ShadeYellow c.Offset(0, 1)
            
        'Otherwise compare Expected and Test values
        Else
            .Result = (.valExpected = .valTest)
        End If
        Set rng = Range(c.Offset(0, 1), c.Offset(0, 3))
        rng = Array(.valExpected, .valTest, .Result)
        If Not .Result Then ShadeYellow c.Offset(0, 3)
    End With
End Sub
'-----------------------------------------------------------------------------------------------
'Purpose: Return expected value from shtT or nothing if not found
'
'   JDL 11/10/21 Modified:
'
Function Expected(clsT)
    Dim c As Range
    
    With clsT
        If Not SheetExists(.wkbk, .shtT) Then Exit Function
        Set c = .wkbk.Sheets(.shtT).Columns(1).Find(.Name, LookAt:=xlWhole)
    End With
    
    If c Is Nothing Then Exit Function
    Expected = c.Offset(0, 1)
End Function
'-----------------------------------------------------------------------------------------------
'Purpose: Assert that specified Boolean test is True
'
'   JDL 11/17/21
'
Sub Assert(clsT, IsTrue)
    With clsT
        If IsEmpty(.valTest) Then .valTest = True
        .valTest = .valTest And IsTrue
    End With
End Sub
'-----------------------------------------------------------------------------------------------
'Purpose: Evaluate overall test results
'
'Inputs:  wkbk [Workbook] workbook object containing test results table
'         shtT [String] worksheet name containing test results table
'         sTestSuite [String] name of test suite
'
'   JDL 11/18/21
'
Sub EvalOverall(wkbk, shtT, sTestSuite)
    Dim w As Variant, IsResults As Boolean, curCell As Range, sMsg As String
    IsResults = True
    Set curCell = wkbk.Sheets(shtT).Cells(2, 1)
    
    'Loop through rows with test names; ignore if blank results
    Do While Len(curCell.Value) > 0
            
        w = curCell.Offset(0, 3).Value
        If Not IsEmpty(w) Then IsResults = IsResults And w
        Set curCell = curCell.Offset(1, 0)
    Loop
    
    'Report overall result to user
    sMsg = sTestSuite & " Testing Failed"
    If IsResults Then sMsg = sTestSuite & " Testing Passed"
    MsgBox sMsg
End Sub
'-----------------------------------------------------------------------------------------------
' Utility methods for checking wkbktest values, addresses and formats
'-----------------------------------------------------------------------------------------------
' Check specified table cell value versus expected
'
' Inputs: colrngKey [Range] key column range to search to identify row to check
'         colrngCk [Range] column value to check at intersection with found row
'         valKey [Variant] key value used to identify row to check
'         valExpected [Variant] expected val at intersection of valKey row and colrngCk
'
' JDL 2/10/22 Modified 3/6/22 debug.prints; 7/18/23 update docstring
'
Function CkTableLocVal(colrngKey, colrngCk, valKey, valExpected) As Boolean
    Dim r As Range, valCompare As Variant
    CkTableLocVal = False
    
    Set r = FindInRange(colrngKey, valKey)
    If r Is Nothing Then
        'Debug.Print "test.CkTableLocVal, ", valKey, "not found in colrngkey"
        Exit Function
    End If
    
    CkTableLocVal = TableLoc(r, colrngCk) = CStr(valExpected)
    'Debug.Print valKey, TableLoc(R, colrngCk) = CStr(valExpected)
End Function
'-----------------------------------------------------------------------------------------------
'Check array of tblRowsCols values based on array of key column values to locate rows
' JDL 3/6/22
Sub TestAryTblVals(Test, colrngKey, colrngTest, aryKeyVals, aryExpect)
    Dim i As Integer
    For i = 0 To UBound(aryKeyVals)
        Test.Assert Test, CkTableLocVal(colrngKey, colrngTest, aryKeyVals(i), aryExpect(i))
    Next i
End Sub
'-----------------------------------------------------------------------------------------------
' Check array of values versus expected (i, j indices allow for aryVals and aryExpect to
' have different lbound 0 or 1
'
' JDL 3/6/23    Modified 7/26/23 Add numeric conversion
'                       2/13/25 CDbl forMac compatibility
'
Sub TestAryVals(Test, aryVals, aryExpect)
    Dim i As Integer, j As Integer, valExpect As Variant
    
    j = LBound(aryExpect)
    For i = LBound(aryVals) To UBound(aryVals)
        valExpect = aryExpect(j)
        
        'Convert expected to number if it is numeric (allows for aryExpect from split list)
        If IsNumeric(valExpect) And Len(valExpect) > 0 Then valExpect = CDbl(valExpect)
        Test.Assert Test, aryVals(i) = valExpect
        j = j + 1
    Next i
End Sub
'-----------------------------------------------------------------------------------------------
' Check Range values versus expected
'
' JDL 7/26/23   Modified 8/21/23 Add ConvertNumerics argument
'
Sub TestRngVals(Test, rng, aryExpect, Optional ConvertNumerics = True)
    Dim i As Integer, c As Variant, valExpect As Variant
    
    'Convert expected to number if it is numeric (e.g. if aryExpect from split list)
    i = LBound(aryExpect)
    For Each c In rng.Cells
    
        'Optionally convert non-zero numeric strings to numbers
        valExpect = aryExpect(i)
        If ConvertNumerics Then valExpect = ConvertValToNumeric(aryExpect(i))
        
        'Debug.Print i, (c.Value = valExpect), c.Value, valExpect
        Test.Assert Test, c.Value = valExpect
        'Debug.Print test.valTest
        i = i + 1
    Next c
    'Debug.Print " "
End Sub
'-----------------------------------------------------------------------------------------------
' Convert type of a string if it is numeric or boolean
' JDL 7/26/23
'
Function ConvertValToNumeric(ByVal val As Variant) As Variant
    If IsNumeric(val) And Len(val) > 0 Then
        val = CDbl(val)
    ElseIf LCase(val) = "true" Or LCase(val) = "false" Then
        val = CBool(val)
    End If
    ConvertValToNumeric = val
End Function
'-----------------------------------------------------------------------------------------------
'Check array of named range address versus array of expected vals
' JDL 3/6/22
Sub CkRngNameAddresses(Test, aryRngNames, aryAddrExpected)
    Dim i As Integer, isok As Boolean
    For i = 0 To UBound(aryRngNames)
        isok = Test.IsNameAddrMatch(Test, aryRngNames(i), aryAddrExpected(i))
        Test.Assert Test, Test.IsNameAddrMatch(Test, aryRngNames(i), aryAddrExpected(i))
    Next i
End Sub
'-----------------------------------------------------------------------------------------------
' Check whether named range's address matches a specified string
'
'Inputs:  sName [String] range name in test.wkbktest
'         sAddress [String] address string to check
'
'   JDL 11/18/21    Modified: 1/6/22 remove test argument; 7/18/23 use .wkbktest
'
Function IsNameAddrMatch(Test, sName, sAddress) As Boolean
    IsNameAddrMatch = False
    With Test
        If Not NameExists(.wkbkTest, sName) Then Exit Function
        If .wkbkTest.Names(sName).RefersToRange.Address = sAddress Then IsNameAddrMatch = True
    End With
End Function
'-----------------------------------------------------------------------------------------------
' Check whether range has specified built-in Style (and that all cells in rng have same style)
'
' Inputs:  rng [String] range in test.wkbktest
'          style [String] Excel Style name
'
' JDL 7/18/23
'
Sub CkStyleMatch(Test, rng, style_expected)
    Dim s As String
    
    'Check that cells in range have same style (Induce error if not same)
    On Error Resume Next
    s = rng.Style
    Test.Assert Test, Err.Number = 0
    If Not Test.valTest Then Exit Sub
    On Error GoTo -1
    
    'If same style in range, check for match with expected
    Test.Assert Test, (rng.Style = style_expected)
End Sub
'-----------------------------------------------------------------------------------------------
' Check whether range has specified NumberFormat (and that all rng cells have same style)
'
' Inputs:  rng [String] range in test.wkbktest
'          style [String] Excel Style name
'
' JDL 7/26/23
'
Sub CkNumFmtMatch(Test, rng, NumFmt_expected)
    Dim s As String
    
    'Check that cells in range have same style (Induce error if not same)
    On Error Resume Next
    s = rng.NumberFormat
    Test.Assert Test, Err.Number = 0
    If Not Test.valTest Then Exit Sub
    On Error GoTo -1
    
    'If same style in range, check for match with expected
    Test.Assert Test, (rng.NumberFormat = NumFmt_expected)
End Sub


'-----------------------------------------------------------------------------------------------
'Test if specified range is empty
' JDL 3/6/22 Modified 7/18/23 add Exit Function

Sub TestRangeIsEmpty(Test, rng)
    Dim c As Variant
    For Each c In rng
        If Not IsEmpty(c) Then
            Test.valTest = False
            Exit Sub
        End If
    Next c
End Sub
'-----------------------------------------------------------------------------------------------
' Create a delimited list of repeated values
' JDL 8/23/23
'
Function CreateLstRepeatVals(val, n_vals, sDelim)
    Dim i As Integer
    CreateLstRepeatVals = ""
    For i = 1 To n_vals
        CreateLstRepeatVals = CreateLstRepeatVals & val & sDelim
    Next i
    
    'Return the list without trailing delimiter after last value
    CreateLstRepeatVals = Left(CreateLstRepeatVals, Len(CreateLstRepeatVals) - 1)
End Function
