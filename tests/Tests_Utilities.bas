Attribute VB_Name = "Tests_Utilities"
'Version 1/28/26
Option Explicit
Const shtTesting As String = "SMdl"
'-----------------------------------------------------------------------------------------------
' Utilities Testing
Sub TestDriver_Utilities()
    Dim procs As New Procedures, AllEnabled As Boolean
    With procs
        .Init procs, ThisWorkbook, "tests_Utilities", "tests_Utilities"
        SetApplEnvir False, False, xlCalculationManual
        
        'Enable testing of all or individual procedures
        AllEnabled = True
        .Utilities.Enabled = False
        .LiteModelSpeed.Enabled = True
    End With
    
    'Tests of ParseBRSales for parsing Better Reports sales data
    With procs.Utilities
        If .Enabled Or AllEnabled Then
            procs.curProcedure = .Name
            test_LastPopulatedCell1 procs
            test_LastPopulatedCell2 procs
            test_LastPopulatedCell3 procs
            test_LastPopulatedCell4 procs
            test_LastPopulatedCell5 procs
            test_BuildMultiCellRange1 procs
            test_BuildMultiCellRange2 procs
            test_BuildMultiCellRange3 procs
            test_BuildMultiCellRange4 procs
            test_IsShtCaseErr procs
            test_iCountAndDeleteStraySheets procs
            test_PopRngMultiKeyTbl procs
            test_RngMultiKey_One procs
            test_RngMultiKey_Two procs
            test_RngMultiKey_Four procs
            test_AryUniqueValsInRange procs
            test_TestAryVals procs
            test_CkNumFmtMatch procs
            test_TestRngVals procs
            test_ConvertValToNumeric procs
        End If
    End With
    
    With procs.LiteModelSpeed
        If .Enabled Or AllEnabled Then
            procs.curProcedure = .Name
            test_Search1 procs
            test_Search2 procs
            test_Search3 procs
            test_Search4 procs
            test_Search5 procs
            test_Search6 procs
            test_Search7 procs
        End If
    End With
    
    procs.EvalOverall procs
    SetApplEnvir True, True, xlCalculationAutomatic
End Sub
'-----------------------------------------------------------------------------------------------
'-----------------------------------------------------------------------------------------------
'This section tests Lite Model Speedup JDL 10/24/25
'-----------------------------------------------------------------------------------------------
'-----------------------------------------------------------------------------------------------
'These tests were used to optimize performance of the ExcelSteps mdlScenario.SetRngFormulaRows, which
'was a drag on refresh speed. The tests show that Excel's range.Find is a significant speedup compared to
'the custom ExcelSteps.FindInRange, which is safe with sheets that contain hidden cells or outlines. Other
'speedups are to not individually reset recipe formulas to Text format and to switch to using .Value2 in place
' of .Value
'
' Mac Excel Benchmarks:
'test_Search2: 1.043 seconds for 500 iterations (original SetRngFormulaRows logic)
'test_Search3: 0.441 seconds for 500 iterations
'test_Search4: 0.465 seconds for 500 iterations
'test_Search5: 0.484 seconds for 500 iterations
'test_Search6: 0.355 seconds for 500 iterations
'test_Search7: 0.355 seconds for 500 iterations (best, optimized version)
'
'Windows Excel Benchmarks
'test_Search2: 1.758 seconds for 500 iterations
'test_Search3: 0.844 seconds for 500 iterations
'test_Search4: 0.887 seconds for 500 iterations
'test_Search5: 0.965 seconds for 500 iterations
'test_Search6: 0.707 seconds for 500 iterations
'test_Search7: 0.703 seconds for 500 iterations
'-----------------------------------------------------------------------------------------------

' like test_Search3 but with .Value2 and best approach to reset formula text
' JDL 10/25/25
'
Sub test_Search7(procs)
    Dim tst As New Test: tst.Init tst, "test_Search7"
    Dim tbl As Object, rngStepsVars As Range, colrngVarNames As Range
    Dim rngFormulaRows As Range, r As Range, r_formula As Range, sVar As String
    Dim i As Long, j As Long, n As Long, timeStart As Double, timeEnd As Double
    
    n = 500
    helper_SetSimRecipeTbl tst, tbl, rngStepsVars, colrngVarNames
    
    With tst
        'Start timer and run loop
        timeStart = Timer
        
        For j = 1 To n
            'Fix all formula cells once outside of loop
            With Intersect(tbl.rngrows, tbl.wksht.Columns(4))
                .Value2 = .Formula
            End With
            
            'Initialize to dummy range two rows below tbl.rngRows (Avoid If/Else in loop)
            Set rngFormulaRows = tbl.rngrows.Rows(tbl.rngrows.Rows.Count).Offset(2, 0).Cells(1, 1)
            
            'Iterate on var_1 through var_15
            For i = 1 To 15
                sVar = "var_" & i
                
                'Search for variable in Column B within tgt_sheet rows using VBA Find
                Set r = colrngVarNames.Find(sVar, LookAt:=xlWhole)
                If Not r Is Nothing Then
                    Set r_formula = Intersect(r.EntireRow, tbl.wksht.Columns(4))
                    If (Left(r_formula.Value2, 1) = "=") Then Set rngFormulaRows = Union(rngFormulaRows, r)
                End If
            Next i
            
            'Remove dummy range from result
            Set rngFormulaRows = Intersect(rngFormulaRows, tbl.rngrows)

        Next j
        timeEnd = Timer
        
        'Report timing results
        Debug.Print "test_Search7: " & Format((timeEnd - timeStart), "0.000") & " seconds for " & n & " iterations"
        .Assert tst, rngFormulaRows.Address = "$B$34:$B$36,$B$42:$B$44"
        .Update tst, procs
    End With
End Sub
'-----------------------------------------------------------------------------------------------
' Test based on test_Search4 with formula optimization
' * Move formula fix outside of loop
' * Use .Value2
' JDL 10/25/25
'
Sub test_Search6(procs)
    Dim tst As New Test: tst.Init tst, "test_Search6"
    Dim tbl As Object, rngStepsVars As Range, colrngVarNames As Range
    Dim rngFormulaRows As Range, rngFormulaCol As Range, r As Range, r_formula As Range
    Dim sVar As String, i As Long, j As Long, n As Long, timeStart As Double, timeEnd As Double
    Dim IsFormula As Boolean
    
    n = 500
    helper_SetSimRecipeTbl tst, tbl, rngStepsVars, colrngVarNames
    
    With tst
        
        'Start timer and run loop
        timeStart = Timer

        For j = 1 To n
            'Fix all formula cells once outside of loop
            With Intersect(tbl.rngrows, tbl.wksht.Columns(4))
                .Value2 = .Formula
            End With

            'Initialize to dummy range two rows below tbl.rngRows (Avoid If/Else in loop)
            Set rngFormulaRows = tbl.rngrows.Rows(tbl.rngrows.Rows.Count).Offset(2, 0).Cells(1, 1)
            
            'Iterate on var_1 through var_15
            For i = 1 To 15
                sVar = "var_" & i
                
                'Search for variable using VBA Find
                Set r = colrngVarNames.Find(sVar, LookAt:=xlWhole)
                
                If Not r Is Nothing Then
                    'Check if formula using .Value2
                    Set r_formula = Intersect(r.EntireRow, tbl.wksht.Columns(4))
                    IsFormula = (Left(r_formula.Value2, 1) = "=")
                    If IsFormula Then Set rngFormulaRows = Union(rngFormulaRows, r)
                End If
            Next i
            
            'Remove dummy range from result
            Set rngFormulaRows = Intersect(rngFormulaRows, tbl.rngrows)
        Next j
        timeEnd = Timer
        
        'Report timing results (benchmark: test_Search6: 0.496 seconds for 500 iterations)
        Debug.Print "test_Search6: " & Format((timeEnd - timeStart), "0.000") & " seconds for " & n & " iterations"
        .Assert tst, rngFormulaRows.Address = "$B$34:$B$36,$B$42:$B$44"

        .Update tst, procs
    End With
End Sub
'-----------------------------------------------------------------------------------------------
' Test using contiguous rngSearch (not as fast as multirange ala test_Search4)
' JDL 10/25/25
'
Sub test_Search5(procs)
    Dim tst As New Test: tst.Init tst, "test_Search5"
    Dim tbl As Object, rngStepsVars As Range, colrngVarNames As Range
    Dim rngFormulaRows As Range, rngSearch As Range, r As Range, rFirst As Range, r_formula As Range, sVar As String
    Dim i As Long, j As Long, n As Long, timeStart As Double, timeEnd As Double
    Dim IsFormula As Boolean, iRowFirst As Long, iRowLast As Long
    
    n = 500
    helper_SetSimRecipeTbl tst, tbl, rngStepsVars, colrngVarNames
    
    With tst
        i = rngStepsVars.Areas.Count
        With tbl.wksht
            Set rngSearch = .Range(rngStepsVars.Rows(1), rngStepsVars.Areas(i).Rows(rngStepsVars.Areas(i).Rows.Count))
            Set rngSearch = Intersect(rngSearch, .Columns(2))
        End With
        
        'Start timer and run loop
        timeStart = Timer
        For j = 1 To n
            'Initialize to dummy range two rows below tbl.rngRows
            Set rngFormulaRows = tbl.rngrows.Rows(tbl.rngrows.Rows.Count).Offset(2, 0).Cells(1, 1)
            
            'Iterate on var_1 through var_15
            For i = 1 To 15
                sVar = "var_" & i
                
                'Search contiguous range for occurrence
                Set r = rngSearch.Find(sVar, LookAt:=xlWhole)
                
                'Check if r intersects rngStepsVars (validates it's in target rows)
                If Not r Is Nothing Then
                    If Not Intersect(r, rngStepsVars) Is Nothing Then
                        'Ensure formula displayed as text (not evaluated formula/error etc.)
                        Set r_formula = Intersect(r.EntireRow, tbl.wksht.Columns(4))
                        r_formula.Value = r_formula.Formula
                        
                        'if formula, add to rngFormulaRows
                        IsFormula = (Left(r_formula.Value, 1) = "=")
                        If IsFormula Then Set rngFormulaRows = Union(rngFormulaRows, r)
                    End If
                End If
            Next i
            
            'Remove dummy range from result
            Set rngFormulaRows = Intersect(rngFormulaRows, tbl.rngrows)
        Next j
        timeEnd = Timer
        
        'Report timing results (Benchmark: test_Search5: 0.992 seconds for 500 iterations)
        Debug.Print "test_Search5: " & Format((timeEnd - timeStart), "0.000") & " seconds for " & n & " iterations"
        .Assert tst, rngFormulaRows.Address = "$B$34:$B$36,$B$42:$B$44"

        .Update tst, procs
    End With
End Sub
'-----------------------------------------------------------------------------------------------
' Test using dummy range to avoid If/Else check within loop
' JDL 10/25/25
'
Sub test_Search4(procs)
    Dim tst As New Test: tst.Init tst, "test_Search4"
    Dim tbl As Object, rngStepsVars As Range, colrngVarNames As Range
    Dim rngFormulaRows As Range, r As Range, r_formula As Range, sVar As String
    Dim i As Long, j As Long, n As Long, timeStart As Double, timeEnd As Double
    Dim IsFormula As Boolean
    
    n = 500
    helper_SetSimRecipeTbl tst, tbl, rngStepsVars, colrngVarNames
    
    With tst
        'Start timer and run loop
        timeStart = Timer
        For j = 1 To n
            'Initialize to dummy range two rows below tbl.rngRows
            Set rngFormulaRows = tbl.rngrows.Rows(tbl.rngrows.Rows.Count).Offset(2, 0).Cells(1, 1)
            
            'Iterate on var_1 through var_15
            For i = 1 To 15
                sVar = "var_" & i
                
                'Search for variable  using VBA Find
                Set r = colrngVarNames.Find(sVar, LookAt:=xlWhole)
                
                If Not r Is Nothing Then
                    'Ensure formula displayed as text (not evaluated formula/error etc.)
                    Set r_formula = Intersect(r.EntireRow, tbl.wksht.Columns(4))
                    r_formula.Value = r_formula.Formula
                    
                    'if formula, add to rngFormulaRows
                    IsFormula = (Left(r_formula.Value, 1) = "=")
                    If IsFormula Then Set rngFormulaRows = Union(rngFormulaRows, r)
                End If
            Next i
            
            'Remove dummy range from result
            Set rngFormulaRows = Intersect(rngFormulaRows, tbl.rngrows)
        Next j
        timeEnd = Timer
        
        'Report timing results (Benchmark: test_Search4: 0.773 seconds for 500 iterations)
        'Slightly slower than test_Search3 without dummy range
        Debug.Print "test_Search4: " & Format((timeEnd - timeStart), "0.000") & " seconds for " & n & " iterations"

        .Update tst, procs
    End With
End Sub
'-----------------------------------------------------------------------------------------------
' Test search algorithm using VBA Range.Find instead of FindInRange
' JDL 10/25/25
'
Sub test_Search3(procs)
    Dim tst As New Test: tst.Init tst, "test_Search3"
    Dim tbl As Object, rngStepsVars As Range, colrngVarNames As Range
    Dim rngFormulaRows As Range, r As Range, r_formula As Range, sVar As String
    Dim i As Long, j As Long, n As Long, timeStart As Double, timeEnd As Double
    Dim IsFormula As Boolean
    
    n = 500
    helper_SetSimRecipeTbl tst, tbl, rngStepsVars, colrngVarNames
    
    With tst
        'Start timer and run loop
        timeStart = Timer
        For j = 1 To n
            Set rngFormulaRows = Nothing
            
            'Iterate on var_1 through var_15
            For i = 1 To 15
                sVar = "var_" & i
                
                'Search for variable in Column B within tgt_sheet rows using VBA Find
                Set r = colrngVarNames.Find(sVar, LookAt:=xlWhole)
                
                If Not r Is Nothing Then
                    'Ensure formula displayed as text (not evaluated formula/error etc.)
                    Set r_formula = Intersect(r.EntireRow, tbl.wksht.Columns(4))
                    r_formula.Value = r_formula.Formula
                    
                    'Check if formula exists
                    IsFormula = (Left(r_formula.Value, 1) = "=")
                    
                    'Add variable to multirange for calculated variables
                    If IsFormula Then
                        If rngFormulaRows Is Nothing Then
                            Set rngFormulaRows = r
                        Else
                            Set rngFormulaRows = Union(rngFormulaRows, r)
                        End If
                    End If
                End If
            Next i
        Next j
        timeEnd = Timer
        
        'Report timing results (Benchmark: test_Search3: 0.699 seconds for 500 iterations)
        Debug.Print "test_Search3: " & Format((timeEnd - timeStart), "0.000") & " seconds for " & n & " iterations"

        .Update tst, procs
    End With
End Sub
'-----------------------------------------------------------------------------------------------
' Test alternate search algorithm iterating on var_i values
' JDL 10/25/25
'
Sub test_Search2(procs)
    Dim tst As New Test: tst.Init tst, "test_Search2"
    Dim tbl As Object, rngStepsVars As Range, colrngVarNames As Range
    Dim rngFormulaRows As Range, r As Range, r_formula As Range, sVar As String
    Dim i As Long, j As Long, n As Long, timeStart As Double, timeEnd As Double
    Dim IsFormula As Boolean
    
    n = 500
    helper_SetSimRecipeTbl tst, tbl, rngStepsVars, colrngVarNames
    
    With tst
        'Start timer and run loop
        timeStart = Timer
        For j = 1 To n
            Set rngFormulaRows = Nothing
            
            'Iterate on var_1 through var_15
            For i = 1 To 15
                sVar = "var_" & i
                
                'Search for variable in Column B within tgt_sheet rows
                Set r = Excelsteps.FindInRange(colrngVarNames, sVar)
                
                If Not r Is Nothing Then
                    'Ensure formula displayed as text (not evaluated formula/error etc.)
                    Set r_formula = Intersect(r.EntireRow, tbl.wksht.Columns(4))
                    r_formula.Value = r_formula.Formula
                    
                    'Check if formula exists
                    IsFormula = (Left(r_formula.Value, 1) = "=")
                    
                    'Add variable to multirange for calculated variables
                    If IsFormula Then
                        If rngFormulaRows Is Nothing Then
                            Set rngFormulaRows = r
                        Else
                            Set rngFormulaRows = Union(rngFormulaRows, r)
                        End If
                    End If
                End If
            Next i
        Next j
        timeEnd = Timer
        
        'Report timing results (Benchmark: test_Search2: 1.672 seconds for 500 iterations)
        'Report timing results (Benchmark: test_Search2: 1.453 seconds for 500 iterations)
        Debug.Print "test_Search2: " & Format((timeEnd - timeStart), "0.000") & " seconds for " & n & " iterations"

        .Update tst, procs
    End With
End Sub
'Simulate current (10/24/25) algorithm
'
Sub test_Search1(procs)
    Dim tst As New Test: tst.Init tst, "test_Search1"
    Dim tbl As Object, rngStepsVars As Range, colrngVarNames As Range
    Dim i As Long, n As Long, timeStart As Double, timeEnd As Double
    
    n = 1
    'n = 500
    helper_SetSimRecipeTbl tst, tbl, rngStepsVars, colrngVarNames
    
    With tst
        'Start timer and run loop
        timeStart = Timer
        For i = 1 To n
            Set rngStepsVars = KeyColRng(tbl, Array(tbl.wksht.Columns(1)), Array("tgt_sheet"))
        Next i
        timeEnd = Timer
        
        'Report timing results (Benchmark: test_Search1: 0.406 seconds for 500 iterations)
        'Debug.Print "test_Search1: " & Format((timeEnd - timeStart), "0.000") & " seconds for " & n & " iterations"

        .Update tst, procs
    End With
End Sub

Sub helper_SetSimRecipeTbl(tst, tbl, rngStepsVars, colrngVarNames)
    With tst
        
        'Populate sheet with simulated recipe values
        PopulateSimRecipe .wkbkTest, shtTesting
        
        'Initialize and provision tbl
        Set tbl = Excelsteps.New_tbl
        Set tbl.wksht = .wkbkTest.Sheets(shtTesting)
        tbl.sht = shtTesting
        
        'Set table rows range
        Set tbl.rngrows = Range(tbl.wksht.Cells(2, 1), tbl.wksht.Columns(1).Cells(tbl.wksht.Rows.Count, 1).End(xlUp)).EntireRow
        
        'Get range of rows with "tgt_sheet"
        Set rngStepsVars = KeyColRng(tbl, Array(tbl.wksht.Columns(1)), Array("tgt_sheet"))
        
        'Set search column range
        Set colrngVarNames = Intersect(rngStepsVars, tbl.wksht.Columns(2))
    End With
End Sub

'<<Additional code not shown for brevity>>

'-------------------------------------------------------------------------------------
'This section tests Function LastPopulatedCell - JDL 12/3/21
'-------------------------------------------------------------------------------------
'Last Populated Cell in row (no outline etc.)
'JDL 12/3/21; Refactored 10/23/25 for procs
'
Sub test_LastPopulatedCell1(procs)
    Dim tst As New Test: tst.Init tst, "test_LastPopulatedCell1"
    Dim rng As Range, wksht As Worksheet
    
    With tst
        Set wksht = .wkbkTest.Sheets(shtTesting)
    
        'Populate row 1 cells
        PopulateRow .wkbkTest, shtTesting, 2
        Set rng = Range(wksht.Cells(1, 2), wksht.Cells(1, 7))
        .Assert tst, ListFromArray(rng) = "a,b,,c,,d"
                           
        'Check last populated cell
        Set rng = Excelsteps.rngLastPopCell(wksht.Rows(1), xlToRight)
        .Assert tst, (rng.Column = 7)
        
        .Update tst, procs
    End With
End Sub
'-------------------------------------------------------------------------------------
'Last Populated Cell in row (with outline and a hidden column)
'JDL 12/3/21; Refactored 10/23/25 for procs
'
Sub test_LastPopulatedCell2(procs)
    Dim tst As New Test: tst.Init tst, "test_LastPopulatedCell2"
    Dim rng As Range, wksht As Worksheet
    
    With tst
        Set wksht = .wkbkTest.Sheets(shtTesting)
    
        'Populate row 1 cells
        PopulateRow .wkbkTest, shtTesting, 2
        Set rng = Range(wksht.Cells(1, 2), wksht.Cells(1, 7))
        .Assert tst, ListFromArray(rng) = "a,b,,c,,d"
                           
        'Add a column outline and hide a column
        With wksht
            .Columns(7).Columns.Group
            .Outline.ShowLevels ColumnLevels:=1
            .Columns(2).EntireColumn.Hidden = True
        End With
                           
        'Check last populated cell
        Set rng = Excelsteps.rngLastPopCell(wksht.Rows(1), xlToRight)
        .Assert tst, (rng.Column = 7)
        .Update tst, procs
    End With
End Sub
'-------------------------------------------------------------------------------------
'Last Populated Cell in row (with outline and a hidden column; start first column)
'JDL 12/3/21; Refactored 10/23/25 for procs
'
Sub test_LastPopulatedCell3(procs)
    Dim tst As New Test: tst.Init tst, "test_LastPopulatedCell3"
    Dim rng As Range, wksht As Worksheet
    
    With tst
        Set wksht = .wkbkTest.Sheets(shtTesting)
    
        'Populate row 1 cells
        PopulateRow .wkbkTest, shtTesting, 1
        Set rng = Range(wksht.Cells(1, 1), wksht.Cells(1, 6))
        .Assert tst, ListFromArray(rng) = "a,b,,c,,d"
                           
        'Add a column outline and hide a column
        With wksht
            .Columns(7).Columns.Group
            .Outline.ShowLevels ColumnLevels:=1
            .Columns(2).EntireColumn.Hidden = True
        End With
                           
        'Check last populated cell
        Set rng = Excelsteps.rngLastPopCell(wksht.Rows(1), xlToRight)
        .Assert tst, (rng.Column = 6)
        
        .Update tst, procs
    End With
End Sub
'-------------------------------------------------------------------------------------
'Last Populated Cell in column (with hidden rows)
'JDL 12/3/21; Refactored 10/23/25 for procs
'
Sub test_LastPopulatedCell4(procs)
    Dim tst As New Test: tst.Init tst, "test_LastPopulatedCell4"
    Dim rng As Range, wksht As Worksheet
    
    With tst
        Set wksht = .wkbkTest.Sheets(shtTesting)
    
        'Populate column 2 cells
        PopulateCol .wkbkTest, shtTesting, 2
        Set rng = Range(wksht.Cells(2, 2), wksht.Cells(7, 2))
        .Assert tst, ListFromArray(rng) = "a,b,,c,,d"
                           
        'Hide some rows
        Range(wksht.Rows(6), wksht.Rows(8)).EntireRow.Hidden = True
                           
        'Check last populated cell
        Set rng = Excelsteps.rngLastPopCell(wksht.Columns(2), xlDown)
        .Assert tst, (rng.Row = 7)
        
        .Update tst, procs
    End With
End Sub
'-------------------------------------------------------------------------------------
'Last Populated Cell in column (no hidden rows)
'JDL 12/5/24 While troubleshoot mdl.Provision issue; Refactored 10/23/25 for procs
'
Sub test_LastPopulatedCell5(procs)
    Dim tst As New Test: tst.Init tst, "test_LastPopulatedCell5"
    Dim rng As Range, wksht As Worksheet
    
    With tst
        Set wksht = .wkbkTest.Sheets(shtTesting)
    
        'Populate column 2 cells
        PopulateCol2 .wkbkTest, shtTesting
        Set rng = Range(wksht.Cells(2, 2), wksht.Cells(7, 2))
        .Assert tst, ListFromArray(rng) = "a,b,,c,,d"
                           
        'Check last populated cell
        Set rng = Excelsteps.rngLastPopCell(wksht.Columns(2), xlDown)
        .Assert tst, (rng.Row = 7)
        .Update tst, procs
    End With
End Sub
'-------------------------------------------------------------------------------------
'This section tests Function BuildMultiCellRange - JDL 12/3/21
'-------------------------------------------------------------------------------------
'Multi-cell range (without outline or hidden columns)
'Refactored 10/23/25 for procs
'
Sub test_BuildMultiCellRange1(procs)
    Dim tst As New Test: tst.Init tst, "test_BuildMultiCellRange1"
    Dim rng As Range, wksht As Worksheet
    
    With tst
        Set wksht = .wkbkTest.Sheets(shtTesting)
    
        'Populate row 1 cells
        PopulateRow .wkbkTest, shtTesting, 2
        Set rng = Range(wksht.Cells(1, 2), wksht.Cells(1, 7))
        .Assert tst, ListFromArray(rng) = "a,b,,c,,d"
           
        'Build/check multicell range
        Set rng = BuildMultiCellRange(wksht.Rows(1), Range(wksht.Columns(1), wksht.Columns(20)))
        .Assert tst, (rng.Address = Union(wksht.Cells(1, 2), _
            wksht.Cells(1, 3), wksht.Cells(1, 5), _
            wksht.Cells(1, 7)).Address)
        .Update tst, procs
    End With
End Sub
'-------------------------------------------------------------------------------------
'Multi-cell range in columns (with outline)
'Refactored 10/23/25 for procs
'
Sub test_BuildMultiCellRange2(procs)
    Dim tst As New Test: tst.Init tst, "test_BuildMultiCellRange2"
    Dim rng As Range, colFirst As Range, colLast As Range, wksht As Worksheet
    
    With tst
        Set wksht = .wkbkTest.Sheets(shtTesting)
    
        'Populate row 1 cells
        PopulateRow .wkbkTest, shtTesting, 2
        Set rng = Range(wksht.Cells(1, 2), wksht.Cells(1, 7))
        .Assert tst, ListFromArray(rng) = "a,b,,c,,d"
            
        With wksht
        
            'Add a column outline and hide a column
            .Columns(7).Columns.Group
            .Outline.ShowLevels ColumnLevels:=1
            .Columns(2).EntireColumn.Hidden = True
            
             'Build/check multicell range - Start to right of .Columns(1)
             Set colFirst = .Columns(2)
             Set colLast = .Columns(.Columns.Count)
             Set rng = BuildMultiCellRange(.Rows(1), Range(colFirst, colLast))
        End With
        
        .Assert tst, (rng.Address = Union(wksht.Cells(1, 2), _
            wksht.Cells(1, 3), wksht.Cells(1, 5), _
            wksht.Cells(1, 7)).Address)
        .Update tst, procs
    End With
End Sub
'-------------------------------------------------------------------------------------
'Multi-cell range in rows (with hidden rows)
'Refactored 10/23/25 for procs
'
Sub test_BuildMultiCellRange3(procs)
    Dim tst As New Test: tst.Init tst, "test_BuildMultiCellRange3"
    Dim rng As Range, rowfirst As Range, rowlast As Range, wksht As Worksheet
    
    With tst
        Set wksht = .wkbkTest.Sheets(shtTesting)
        
        'Populate row 1 cells
        PopulateCol .wkbkTest, shtTesting, 2
        Set rng = Range(wksht.Cells(2, 2), wksht.Cells(7, 2))
        .Assert tst, ListFromArray(rng) = "a,b,,c,,d"
            
        'Hide some rows
        Range(wksht.Rows(6), wksht.Rows(8)).EntireRow.Hidden = True
           
        'Build/check multicell range - Start below .Rows(1)
        Set rowfirst = wksht.Rows(2)
        Set rowlast = wksht.Rows(wksht.Rows.Count)
        Set rng = BuildMultiCellRange(wksht.Columns(2), Range(rowfirst, rowlast))
        .Assert tst, (rng.Address = Union(wksht.Cells(2, 2), _
            wksht.Cells(3, 2), wksht.Cells(5, 2), _
            wksht.Cells(7, 2)).Address)
        .Update tst, procs
    End With
End Sub
'-------------------------------------------------------------------------------------
'Multi-cell range in rows (with stray text below defined row range)
'Refactored 10/23/25 for procs
'
Sub test_BuildMultiCellRange4(procs)
    Dim tst As New Test: tst.Init tst, "test_BuildMultiCellRange4"
    Dim rng As Range, rng2 As Range, rowfirst As Range, rowlast As Range, wksht As Worksheet
    
    With tst
        Set wksht = .wkbkTest.Sheets(shtTesting)
    
        'Populate row 1 cells
        PopulateCol .wkbkTest, shtTesting, 2
        Set rng = Range(wksht.Cells(2, 2), wksht.Cells(7, 2))
        .Assert tst, ListFromArray(rng) = "a,b,,c,,d"
                           
        'Build/check multicell range - Start below .Rows(1)
        Set rowfirst = wksht.Rows(2)
        Set rowlast = wksht.Rows(6)
        Set rng = BuildMultiCellRange(wksht.Columns(2), Range(rowfirst, rowlast))
        .Assert tst, (rng.Address = Union(wksht.Cells(2, 2), _
            wksht.Cells(3, 2), wksht.Cells(5, 2)).Address)
        .Update tst, procs
    End With
End Sub
'-------------------------------------------------------------------------------------
'This section tests utility to resolve Worksheet Name case sensitivity
' JDL 12/20/21
'-------------------------------------------------------------------------------------
'Validate IsShtCaseErr function to handle Excel bug where sheet names case sensitive
'
'JDL 1/20/22; Refactored 10/23/25 for procs
'
Sub test_IsShtCaseErr(procs)
    Dim tst As New Test: tst.Init tst, "test_IsShtCaseErr"
    Dim shtTLower As String, sht
    
    With tst
        'shtTests exists; make a lower case version of its name
        shtTLower = LCase(shtTesting)
        sht = shtTLower
        .Assert tst, (shtTLower <> shtTesting)
        .Assert tst, (SheetExists(.wkbkTest, shtTesting))
        .Assert tst, (Not SheetExists(.wkbkTest, shtTLower))
        
        'Boolean True signifies changed byRef sht arg (sht exists but w/ case diff)
        .Assert tst, IsShtCaseErr(.wkbkTest, sht)
        
        'IsShtCaseErr fixes sht to be identical to shtTesting
        .Assert tst, (sht = shtTesting)
        .Update tst, procs
    End With
End Sub
'---------------------------------------------------------------------------------------------
'---------------------------------------------------------------------------------------------
'Validate iCountAndDeleteStraySheets function
'JDL 1/20/22; Refactored 10/23/25 for procs
'
Sub test_iCountAndDeleteStraySheets(procs)
    Dim tst As New Test: tst.Init tst, "test_iCountAndDeleteStraySheets"
    Dim Test As New Tests
    Dim sht As String, i As Integer
    
    With tst
    
        'Delete any pre-existing stray sheets
        i = iCountAndDeleteStraySheets(.wkbkTest)
        
        'Add an empty sheet
        sht = "Sheet_xxx" 'Mimic Excel default
        If SheetExists(.wkbkTest, sht) Then DeleteSheet .wkbkTest, sht
        AddSheet .wkbkTest, sht
    
        .Assert tst, (iCountAndDeleteStraySheets(.wkbkTest) = 1)
        .Assert tst, (Not SheetExists(.wkbkTest, sht))
        
        'Test on non-empty sheets
        AddSheet .wkbkTest, sht
        .wkbkTest.Sheets(sht).Cells(1, 1) = "xxx"
        .Assert tst, (iCountAndDeleteStraySheets(.wkbkTest) = 0)
        DeleteSheet .wkbkTest, sht
        
        AddSheet .wkbkTest, sht
        .wkbkTest.Sheets(sht).Cells(5, 5) = "xxx"
        .Assert tst, (iCountAndDeleteStraySheets(.wkbkTest) = 0)
        DeleteSheet .wkbkTest, sht
        
        .Update tst, procs
    End With
End Sub
'-------------------------------------------------------------------------------------
'Populate a table
'Refactored 10/23/25 for procs
'
Sub test_PopRngMultiKeyTbl(procs)
    Dim tst As New Test: tst.Init tst, "test_PopRngMultiKeyTbl"
    Dim wksht As Worksheet, cellHome As Range, tbl As Object
    
    SetApplEnvir False, False, xlCalculationAutomatic
    
    Helper_SetTestSht tst, tbl, wksht, cellHome
    
    'Populate the table
    PopulateRngMultiKeyLookupTbl wksht, cellHome
    
    'Provision as table to check dimensions
    With tst
        tbl.Provision tbl, .wkbkTest, True, shtTesting
        .Assert tst, (tbl.nCols = 5)
        .Assert tst, (tbl.nRows = 9)
        .Update tst, procs
    End With

End Sub
'-------------------------------------------------------------------------------------
'This section tests Function rngMultiKey - JDL 1/27/22
'-------------------------------------------------------------------------------------
'One Key lookup
'Refactored 10/23/25 for procs
'
Sub test_RngMultiKey_One(procs)
    Dim tst As New Test: tst.Init tst, "test_RngMultiKey_One"
    Dim wksht As Worksheet, cellHome As Range, r As Range
    Dim aryKeyCols As Variant, aryKeyVals As Variant, tbl As Object
    
    SetApplEnvir False, False, xlCalculationAutomatic
    
    Helper_SetTestSht tst, tbl, wksht, cellHome
    
    'Populate and provision the table
    PopulateRngMultiKeyLookupTbl wksht, cellHome
    
    With tst
        tbl.Provision tbl, .wkbkTest, True, shtTesting
    
        'Set Key Col (array)
        Set r = tbl.rngTblHeaderVal(tbl, "Key_1")
        .Assert tst, (r.Address = "$A$1")
        aryKeyCols = Array(Intersect(r.EntireColumn, tbl.rngrows))
        aryKeyVals = Array("C")

        'Lookup 'C' from Key_1
        Set r = rngMultiKey(tbl, aryKeyCols, aryKeyVals)
        .Assert tst, (Not r Is Nothing)
        If Not r Is Nothing Then .Assert tst, (r.Address = "$3:$3")
        .Update tst, procs
    End With

End Sub

'-------------------------------------------------------------------------------------
'Two Key lookup
'Refactored 10/23/25 for procs
'
Sub test_RngMultiKey_Two(procs)
    Dim tst As New Test: tst.Init tst, "test_RngMultiKey_Two"
    Dim wksht As Worksheet, cellHome As Range
    Dim aryKeyCols As Variant, aryKeyVals As Variant, tbl As Object
    Dim R1 As Range, R2 As Range
    
    SetApplEnvir False, False, xlCalculationAutomatic
    
    Helper_SetTestSht tst, tbl, wksht, cellHome
    
    'Populate and provision the table
    PopulateRngMultiKeyLookupTbl wksht, cellHome
    
    With tst
        tbl.Provision tbl, .wkbkTest, True, shtTesting
    
        'Set Key Col (array)
        Set R1 = tbl.rngTblHeaderVal(tbl, "Key_1")
        Set R1 = Intersect(R1.EntireColumn, tbl.rngrows)
        Set R2 = tbl.rngTblHeaderVal(tbl, "Key_2")
        Set R2 = Intersect(R2.EntireColumn, tbl.rngrows)
        
        aryKeyCols = Array(R1, R2)
        aryKeyVals = Array("A", "AA")

        'Lookup two-row range from Key_1 and Key_2
        Set R1 = rngMultiKey(tbl, aryKeyCols, aryKeyVals)
        .Assert tst, (Not R1 Is Nothing)
        If Not R1 Is Nothing Then .Assert tst, (R1.Address = "$2:$2,$4:$4")
        .Update tst, procs
    End With
End Sub
'-------------------------------------------------------------------------------------
'Four Key lookup
'Refactored 10/23/25 for procs
'
Sub test_RngMultiKey_Four(procs)
    Dim tst As New Test: tst.Init tst, "test_RngMultiKey_Four"
    Dim wksht As Worksheet, cellHome As Range
    Dim aryKeyCols As Variant, aryKeyVals As Variant, tbl As Object
    Dim R1 As Range, R2 As Range, R3 As Range, R4 As Range, r As Range
    
    SetApplEnvir False, False, xlCalculationAutomatic
    
    Helper_SetTestSht tst, tbl, wksht, cellHome
    
    'Populate and provision the table
    PopulateRngMultiKeyLookupTbl wksht, cellHome
    
    With tst
        tbl.Provision tbl, .wkbkTest, True, shtTesting
    
        'Set Key Col array
        Set R1 = tbl.rngTblHeaderVal(tbl, "Key_1")
        Set R1 = Intersect(R1.EntireColumn, tbl.rngrows)
        Set R2 = tbl.rngTblHeaderVal(tbl, "Key_2")
        Set R2 = Intersect(R2.EntireColumn, tbl.rngrows)
        Set R3 = tbl.rngTblHeaderVal(tbl, "Key_3")
        Set R3 = Intersect(R3.EntireColumn, tbl.rngrows)
        Set R4 = tbl.rngTblHeaderVal(tbl, "Key_4")
        Set R4 = Intersect(R4.EntireColumn, tbl.rngrows)
        
        aryKeyCols = Array(R1, R2, R3, R4)
        aryKeyVals = Array("B", "BC", "X", 2)

        'Lookup one-row range from Key_1 to Key_4
        Set r = rngMultiKey(tbl, aryKeyCols, aryKeyVals)
        .Assert tst, (Not r Is Nothing)
        If Not r Is Nothing Then .Assert tst, (r.Address = "$8:$8")
        .Update tst, procs
    End With
End Sub
'---------------------------------------------------------------------------------------------
'This section validates the aryUniqueValsInRange function - JDL 3/12/22
' Refactored 10/23/25
'---------------------------------------------------------------------------------------------
'Contiguous column range and Two Key lookup test_RngMultiKey_Two non-contig.
'Refactored 10/23/25 for procs
Sub test_AryUniqueValsInRange(procs)
    Dim tst As New Test: tst.Init tst, "test_AryUniqueValsInRange"
    Dim wksht As Worksheet, cellHome As Range, tbl As Object
    Dim R1 As Range, R2 As Range, R3 As Range, aryUnique As Variant
    
    SetApplEnvir False, False, xlCalculationAutomatic
    
    Helper_SetTestSht tst, tbl, wksht, cellHome
    
    'Populate and provision the table
    PopulateRngMultiKeyLookupTbl wksht, cellHome
    
    With tst
    
        tbl.Provision tbl, .wkbkTest, True, shtTesting
    
        'Set Key Col (array)
        Set R1 = tbl.rngTblHeaderVal(tbl, "Key_1")
        Set R1 = Intersect(R1.EntireColumn, tbl.rngrows)
        Set R2 = tbl.rngTblHeaderVal(tbl, "Key_2")
        Set R2 = Intersect(R2.EntireColumn, tbl.rngrows)
        Set R3 = tbl.rngTblHeaderVal(tbl, "Key_4")
        Set R3 = Intersect(R3.EntireColumn, tbl.rngrows)
        
        'Get array of unique vals from Key_2 column (contiguous range)
        aryUnique = aryUniqueValsInRange(R3)
        .Assert tst, (UBound(aryUnique) = 1)
        .Assert tst, (ListFromArray(aryUnique) = "1,2")
        
        'Get array of unique vals from Key_4 column (contiguous; contains blanks)
        aryUnique = aryUniqueValsInRange(R2)
        .Assert tst, (UBound(aryUnique) = 3)
        .Assert tst, (ListFromArray(aryUnique) = "AA,AB,BC,BD")
        
        'Lookup two-row non-contiguous range from Key_1 and Key_2
        Set R1 = rngMultiKey(tbl, Array(R1, R2), Array("A", "AA"))
        .Assert tst, (R1.Address = "$2:$2,$4:$4")
        
        'Get array of unique vals in Key_3 column
        Set R2 = tbl.rngTblHeaderVal(tbl, "Key_3").EntireColumn
        aryUnique = aryUniqueValsInRange(Intersect(R1, R2))
        .Assert tst, (UBound(aryUnique) = 1)
        .Assert tst, (ListFromArray(aryUnique) = "X,Y")
        
        .Update tst, procs
    End With
End Sub
'-------------------------------------------------------------------------------------------------------
' Test Class method for comparing array of values versus expected
'JDL 7/26/23; Refactored 10/23/25 for procs
'
Sub test_TestAryVals(procs)
    Dim tst As New Test: tst.Init tst, "test_TestAryVals"
    Dim testDummy As New Test
    Dim aryVals As Variant, aryExpect As Variant, wksht As Worksheet, rng As Range, s As String
    
    With tst
        'Create a dummy Tests instance to use for running checks
        testDummy.Init tst, ""

        'A passing test with Strings in AryVals
        aryVals = Array("a", "b")
        aryExpect = Array("a", "b")
        testDummy.valTest = True
        testDummy.TestAryVals testDummy, aryVals, aryExpect
        .Assert tst, testDummy.valTest
    
        'A failing test with Strings in AryVals
        aryVals = Array("a", "b")
        aryExpect = Array("a", "bb")
        testDummy.valTest = True
        testDummy.TestAryVals testDummy, aryVals, aryExpect
        .Assert tst, Not testDummy.valTest
        
        Set wksht = .wkbkTest.Sheets("SMdl")
        
        'Vals array from row Range [strings] -
        ClearTestSheetAndNames wksht
        s = "a,,c"
        Set rng = Range(wksht.Cells(1, 1), wksht.Cells(1, 3))
        rng = Split(s, ",")
        .Assert tst, IsEmpty(wksht.Cells(1, 2))
        .Assert tst, wksht.Cells(1, 3) = "c"
        
        'Check the values in Row 1 range
        aryVals = AryFromRowRng(rng, iRow:=1)
        testDummy.valTest = True
        testDummy.TestAryVals testDummy, aryVals, Split("a,,c", ",")
        .Assert tst, testDummy.valTest
    
        'Vals array from col Range [strings] -
        ClearTestSheetAndNames wksht
        Set rng = Range(wksht.Cells(1, 1), wksht.Cells(3, 1))
        rng = WorksheetFunction.Transpose(Split(s, ","))
        .Assert tst, IsEmpty(wksht.Cells(2, 1))
        .Assert tst, wksht.Cells(3, 1) = "c"
        
        'Check the values in Column 1 range
        aryVals = AryFromColRng(rng, iCol:=1)
        testDummy.valTest = True
        testDummy.TestAryVals testDummy, aryVals, Split("a,,c", ",")
        .Assert tst, testDummy.valTest

        'Vals array from row Range [mix strings and numeric]
        ClearTestSheetAndNames wksht
        s = "a,,10"
        Set rng = Range(wksht.Cells(1, 1), wksht.Cells(1, 3))
        rng.Cells(1) = "a"
        rng.Cells(3) = 10
        .Assert tst, IsEmpty(wksht.Cells(1, 2))
        .Assert tst, wksht.Cells(1, 3) = 10
        
        'Check the values in Row 1 range (Tests.TestAryVals checks aryExpect for numeric)
        aryVals = AryFromRowRng(rng, iRow:=1)
        testDummy.valTest = True
        testDummy.TestAryVals testDummy, aryVals, Split("a,,10", ",")
        .Assert tst, testDummy.valTest

        .Update tst, procs
    End With
End Sub
'-------------------------------------------------------------------------------------------------------
' Tests Class method for comparing array of values versus expected
'JDL 7/26/23; Refactored 10/23/25 for procs
'
Sub test_CkNumFmtMatch(procs)
    Dim tst As New Test: tst.Init tst, "test_CkNumFmtMatch"
    Dim testDummy As New Test, rng As Range, wksht As Worksheet
    
    With tst
        testDummy.Init tst, ""
        Set wksht = .wkbkTest.Sheets(shtTesting)
        
        'set some values and a number format for a range
        ClearTestSheetAndNames wksht
        Set rng = Range(wksht.Cells(1, 1), wksht.Cells(1, 3))
        rng = Array(1, 2, 3)
        rng.NumberFormat = "0.00"
        
        'NumFmt match over entire range
        testDummy.valTest = True
        testDummy.CkNumFmtMatch testDummy, rng, "0.00"
        .Assert tst, testDummy.valTest
        
        'NumFmt mismatch over entire range
        testDummy.valTest = True
        testDummy.CkNumFmtMatch testDummy, rng, "0"
        .Assert tst, Not testDummy.valTest
        
        'Inconsistent format across range
        wksht.Cells(1, 1).NumberFormat = "0"
        testDummy.valTest = True
        testDummy.CkNumFmtMatch testDummy, rng, "0.00"
        .Assert tst, Not testDummy.valTest

        .Update tst, procs
    End With
End Sub
'-------------------------------------------------------------------------------------------------------
' Tests Class method for comparing Range values to expected
'JDL 7/26/23; Refactored 10/23/25 for procs
'
Sub test_TestRngVals(procs)
    Dim tst As New Test: tst.Init tst, "test_TestRngVals"
    Dim testDummy As New Test
    Dim aryVals As Variant, aryExpect As Variant, wksht As Worksheet, rng As Range, s As String
    
    With tst
        'Create a dummy Tests instance to use for running checks
        testDummy.Init tst, ""
        Set wksht = .wkbkTest.Sheets("SMdl")
        
        'Vals array from row Range [strings] -
        ClearTestSheetAndNames wksht
        s = "a,,c"
        Set rng = Range(wksht.Cells(1, 1), wksht.Cells(1, 3))
        rng = Split(s, ",")
        .Assert tst, IsEmpty(wksht.Cells(1, 2))
        .Assert tst, wksht.Cells(1, 3) = "c"
        
        'Check the values in Row 1 range
        testDummy.valTest = True
        testDummy.TestRngVals testDummy, rng, Split("a,,c", ",")
        .Assert tst, testDummy.valTest
    
        'Mismatch
        testDummy.valTest = True
        testDummy.TestRngVals testDummy, rng, Split("a,b,c", ",")
        .Assert tst, Not testDummy.valTest
    
        'Numeric in range
        wksht.Cells(1, 2) = 10
        testDummy.valTest = True
        testDummy.TestRngVals testDummy, rng, Split("a,10,c", ",")
        .Assert tst, testDummy.valTest

        .Update tst, procs
    End With
End Sub
'-------------------------------------------------------------------------------------------------------
' Convert type of a string if it is numeric or boolean
' JDL 7/26/23; Refactored 10/23/25 for procs
'
Sub test_ConvertValToNumeric(procs)
    Dim tst As New Test: tst.Init tst, "test_ConvertValToNumeric"
    Dim ary As Variant, s As String, val As Variant, valNew As Variant

    With tst
        'Initial value is numeric
        val = 10
        valNew = .ConvertValToNumeric(val)
        .Assert tst, VarType(val) = 2
        .Assert tst, VarType(valNew) = 5
        .Assert tst, val = 10
        
        'Initial value is string numeric
        val = "10"
        valNew = .ConvertValToNumeric(val)
        .Assert tst, VarType(val) = 8
        .Assert tst, VarType(valNew) = 5
        .Assert tst, val = 10
        
        'Initial value is string from list Split - check for VBA Class/array weirdness
        ary = Split("a,10", ",")
        valNew = .ConvertValToNumeric(ary(1))
        .Assert tst, VarType(ary(1)) = 8
        .Assert tst, VarType(valNew) = 5
        .Assert tst, val = 10
        
        'Initial value is string non-numeric
        val = "a"
        valNew = .ConvertValToNumeric(val)
        .Assert tst, VarType(val) = 8
        .Assert tst, VarType(valNew) = 8
        .Assert tst, val = "a"

        'Initial value is string Boolean
        val = "True"
        valNew = .ConvertValToNumeric(val)
        .Assert tst, VarType(val) = 8
        .Assert tst, VarType(valNew) = 11
        .Assert tst, valNew

        .Update tst, procs
    End With
End Sub
'-------------------------------------------------------------------------------------------------------
Sub Helper_SetTestSht(tst, tbl, wksht, cellHome)
    Set tbl = Excelsteps.New_tbl
    Set wksht = tst.wkbkTest.Sheets(shtTesting)
    Set cellHome = wksht.Cells(1, 1)
End Sub



