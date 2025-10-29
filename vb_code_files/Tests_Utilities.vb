'Tests_Utilities.vb
'Version 10/23/25
Option Explicit
Const shtTesting As String = "SMdl"
'-----------------------------------------------------------------------------------------------
' Utilities Testing
Sub TestDriverUtilities()
    Dim procs As New Procedures, AllEnabled As Boolean
    With procs
        .Init procs, ThisWorkbook, "ExcelSteps_Utilities", "ExcelSteps_Utilities"
        SetApplEnvir False, False, xlCalculationAutomatic
        
        'Enable testing of all or individual procedures
        AllEnabled = False
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
                Set r = ExcelSteps.FindInRange(colrngVarNames, sVar)
                
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
        Set tbl = ExcelSteps.New_tbl
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
