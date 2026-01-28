VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Test"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Version 5/27/24 Add writing procedure name in results table
'Version 5/10/24 Refactor
'Version of 7/18/23 - add testing utility functions TestAryVals etc.
'Version 2/13/25 - change for Mac compatibility
'Utility subs/functions used: AddSheet, SheetExists, ColWidthAutoFit, ShadeYellow
Option Explicit
Public wkbk As Workbook 'Testing workbook
Public wkbkTest As Workbook 'Workbook with results being tested (may be same as wkbk)
Public sFileTest As String 'Wkbktest file reference
Public shtT As String
Public valExpected As Variant
Public valTest As Boolean
Public Name As String
'-----------------------------------------------------------------------------------------------
'Populate test properties
'Inputs:    test [Tests Class]
'           sName [String] test name
'
'Created:   5/10/24 JDL to replace .Populate method
'
Sub Init(Test, sName)
    With Test
        .valTest = True
        .Name = sName
        Set .wkbkTest = ThisWorkbook
    End With
End Sub
'-----------------------------------------------------------------------------------------------
'Purpose:   Populate test properties
'Inputs:    test [Tests Class]
'           wkbk [Workbook] contains tests sheet
'           shtT [String] tests sheet name for reporting test results
'           sName [String] test name
'
'Created:   11/10/21 JDL    Modified: 6/7/22
'                                     7/18/23 default wkbktest to be wkbk if not specified
'
Sub Populate(Test, wkbk, Optional shtT, Optional sName, Optional wkbkTest, Optional sFileTest)
    With Test
        Set .wkbk = wkbk
        If Not IsMissing(wkbkTest) Then
            Set .wkbkTest = wkbkTest
        Else
            Set .wkbkTest = .wkbk
        End If
        If Not IsMissing(sFileTest) Then .sFileTest = sFileTest
'        If Not IsMissing(shtT) Then .shtT = shtT
        If Not IsMissing(sName) Then .Name = sName
        '.valExpected = Expected(test)
    End With
End Sub
'-----------------------------------------------------------------------------------------------
' Write individual test result to row in tests sheet
'
'Created:   11/10/21 JDL      Modified: 5/27/24
'
Sub Update(tst, procs)
    Dim c As Range, rng As Range
        
    'Find the existing test row or add a new if first time running this test
    With procs.wksht_results
        Set c = procs.colrngTestName.Find(tst.Name, LookAt:=xlWhole)
        If c Is Nothing Then
            Set c = .Cells(.Rows.Count, procs.colrngTestName.Column).End(xlUp).Offset(1, 0)
            c.Value = tst.Name
            procs.colrngTestName.AutoFit
        End If
    End With
    
    'Write the current procedure name and test result in the test's results row
    With tst
        Intersect(c.EntireRow, procs.colrngProcedure) = procs.curProcedure
        Intersect(c.EntireRow, procs.colrngTestResult) = .valTest
        If Not .valTest Then ShadeYellow Intersect(c.EntireRow, procs.colrngTestResult)
    End With
End Sub
'-----------------------------------------------------------------------------------------------
'Purpose: Assert that specified Boolean test is True
'
'   JDL 11/17/21; Modified 5/10/24
'
Sub Assert(tst, IsCheck)
    With tst
        If IsEmpty(.valTest) Then .valTest = True
        .valTest = .valTest And IsCheck
    End With
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
'     4/29/24 remove  TableLoc call - need to verify change
'
Function CkTableLocVal(colrngKey, colrngCk, valKey, valExpected) As Boolean
    Dim r As Range, valCompare As Variant
    CkTableLocVal = False
    
    Set r = FindInRange(colrngKey, valKey)
    If r Is Nothing Then
        'Debug.Print "tst.CkTableLocVal, ", valKey, "not found in colrngkey"
        Exit Function
    End If
    
    CkTableLocVal = Intersect(r.EntireRow, colrngCk) = CStr(valExpected)
    'CkTableLocVal = TableLoc(r, colrngCk) = CStr(valExpected)
    'Debug.Print valKey, TableLoc(R, colrngCk) = CStr(valExpected)
End Function
'-----------------------------------------------------------------------------------------------
' Lookup and return value at intersection of table cell row and column range
' Inputs:   rngCell [Range] single row range (usually single cell from key column)
'           rngCol [Range] table column range
'           iShift [Integer] optional shift row value
'
'Created: 6/20 JDL
'
Function TableLoc(rngCell, rngcol, Optional ishift = 0) As Variant
    TableLoc = Intersect(rngCell.Offset(ishift, 0).EntireRow, rngcol)
End Function
'-----------------------------------------------------------------------------------------------
'Check array of tblRowsCols values based on array of key column values to locate rows
' JDL 3/6/22
Sub TestAryTblVals(tst, colrngKey, colrngTest, aryKeyVals, aryExpect)
    Dim i As Integer
    For i = 0 To UBound(aryKeyVals)
        tst.Assert tst, CkTableLocVal(colrngKey, colrngTest, aryKeyVals(i), aryExpect(i))
    Next i
End Sub
'-----------------------------------------------------------------------------------------------
' Check array of values versus expected (i, j indices allow for aryVals and aryExpect to
' have different lbound 0 or 1
'
' JDL 3/6/23    Modified 7/26/23 Add numeric conversion
'                   2/13/25 change CDec to CDbl for Mac compatibility
'
Sub TestAryVals(tst, aryVals, aryExpect)
    Dim i As Integer, j As Integer, valExpect As Variant
    
    j = LBound(aryExpect)
    For i = LBound(aryVals) To UBound(aryVals)
        valExpect = aryExpect(j)
        
        'Convert expected to number if it is numeric (allows for aryExpect from split list)
        If IsNumeric(valExpect) And Len(valExpect) > 0 Then valExpect = CDbl(valExpect)
        tst.Assert tst, aryVals(i) = valExpect
        j = j + 1
    Next i
End Sub
'-----------------------------------------------------------------------------------------------
' Check Range values versus expected
'
' JDL 7/26/23   Modified 8/21/23 Add ConvertNumerics argument
'
Sub TestRngVals(tst, rng, aryExpect, Optional ConvertNumerics = True)
    Dim i As Integer, c As Variant, valExpect As Variant
    
    'Convert expected to number if it is numeric (e.g. if aryExpect from split list)
    i = LBound(aryExpect)
    For Each c In rng.Cells
    
        'Optionally convert non-zero numeric strings to numbers
        valExpect = aryExpect(i)
        If ConvertNumerics Then valExpect = ConvertValToNumeric(aryExpect(i))
        
        'Debug.Print i, (c.Value = valExpect), c.Value, valExpect
        tst.Assert tst, c.Value = valExpect
        'Debug.Print tst.valTest
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
Sub CkRngNameAddresses(tst, aryRngNames, aryAddrExpected)
    Dim i As Integer, isok As Boolean
    For i = 0 To UBound(aryRngNames)
        isok = tst.IsNameAddrMatch(tst, aryRngNames(i), aryAddrExpected(i))
        tst.Assert tst, tst.IsNameAddrMatch(tst, aryRngNames(i), aryAddrExpected(i))
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
'-------------------------------------------------------------------------------------
' Compare values in row or column range array to expected values array
' JDL 11/17/25
'
Sub CompareRangeToExpected(tst, rngArray As Variant, aryExpected As Variant)
    Dim i As Integer, j As Integer, nRows As Integer, nCols As Integer
    Dim idx As Integer
    
    nRows = UBound(rngArray, 1)
    nCols = UBound(rngArray, 2)
    
    'Check if row array (single row, multiple columns)
    If nRows = 1 Then
        For i = 1 To nCols
            idx = i - 1
            tst.Assert tst, CStr(rngArray(1, i)) = aryExpected(idx)
        Next i
    
    'Check if column array (multiple rows, single column)
    ElseIf nCols = 1 Then
        For i = 1 To nRows
            idx = i - 1
            tst.Assert tst, CStr(rngArray(i, 1)) = aryExpected(idx)
        Next i
    End If
End Sub

