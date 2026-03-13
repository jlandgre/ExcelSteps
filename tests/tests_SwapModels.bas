Attribute VB_Name = "tests_SwapModels"
'Version 11/6/24
Option Explicit
'-----------------------------------------------------------------------------------------------
'Definition for testing parsing - non-default model with $I$4 cellHome
'defn Booleans: IsCalc, IsSuppHeader, IsRngNames, IsMdlNmPrefix, IsLiteModel
Public Const defn_dash As String = "SMdl:4,9:6:T:T:T:T:T:SMdlDash"
'Public Const defn_dest As String = "SMdl:10,9:0:T:T:T:F:T:SMdlDest"
Public Const defn_dest As String = "SMdl:10,9:0:T:T:T:F:T:SMdlType2"

'Constants related to test models used in validation
Public Const nrows_type1 As Integer = 4
Public Const nrows_type2 As Integer = 5
Public Const nrows_both As Integer = 9
Public Const sMdl1 As String = "SMdlType1"
Public Const sMdl2 As String = "SMdlType2"
'-----------------------------------------------------------------------------------------------
' Test suite for mdlScenario Class SwapModels Procedure
' JDL 1/4/22    Modified 8/23/23
'
Sub TestDriver_SwapModels()
    Dim wkbk As Workbook, shtT As String, testsetup As New Tests
    Set wkbk = ThisWorkbook: shtT = "Tests_SwapModels"
    
    'Turn off events and Screenupdataing; calculation Automatic
    SetApplEnvir False, False, xlCalculationAutomatic
    
    'Clear previous test results
    testsetup.InitTestsSheet wkbk, shtT
    
    'Set up and test the tblImport sheet table with its models
    test_FormatMdlImport wkbk, shtT
    test_PopulateMdlImport wkbk, shtT
    
    'Populate "Dashboard" (non-swapped) model
    test_PopulateDashMdl wkbk, shtT
    
    'SwapModels Procedure
    test_InitSwapModels wkbk, shtT
    test_SwapModels1 wkbk, shtT 'no ModelNew specified and No current mdlDest
    test_SwapModels2 wkbk, shtT 'no ModelNew specified and SMdlType2 initially in mdlDest
    test_SwapModels3 wkbk, shtT 'ModelNew specified and no current mdlDest
    test_SwapModels4 wkbk, shtT 'ModelNew SMdlType1 and SMdlType2 initially in mdlDest
    
    'TransferToMdlDest Procedure
    test_InitTransferToMdl wkbk, shtT
    test_TransferTblImportRows1 wkbk, shtT
    test_TransferTblImportRows2 wkbk, shtT
    test_ResetPostTransfer wkbk, shtT
    test_TransferToMdlDest wkbk, shtT

    'TransferToTblImport Procedure
    test_ReadModelName wkbk, shtT
    test_TblImportDeleteModel wkbk, shtT
    test_TransferMdlDestRows wkbk, shtT
    test_DeleteTblImpTrailingBlankRows wkbk, shtT
    test_TransferToTblImport wkbk, shtT 'Also checks for ClearModel and StepsDeleteMdl
    
    'report results
    testsetup.EvalOverall wkbk, shtT, shtT
    SetApplEnvir True, True, xlCalculationAutomatic
End Sub
'-----------------------------------------------------------------------------------------------
'-----------------------------------------------------------------------------------------------
' SwapModels Procedure
'-----------------------------------------------------------------------------------------------
'-----------------------------------------------------------------------------------------------
' SwapModels Procedure where ModelNew specified and no current mdlDest
' JDL 8/23/23
'
Function test_SwapModels4(wkbk, shtTests)
    Dim Test As New Tests
    Test.Populate Test, wkbk, shtTests, "test_SwapModels4"
    SetApplEnvir False, False, xlCalculationAutomatic
    
    'Test that populates test.valTest with True or False
    Dim tblImp As Object: Set tblImp = ExcelSteps.New_tbl
    Dim tbls As Object: Set tbls = ExcelSteps.New_tbl
    Dim mdlDest As Object: Set mdlDest = ExcelSteps.New_mdl

    With Test
    
        'Populate tblImport sheet and put SMdlType2 in mdlDest
        PopulateSMdlType2ToMdlDest Test, mdlDest, tblImp, tbls

        'SwapModels Type1 to mdlDest; Type2 to tblImport
        .Assert Test, mdlDest.SwapModels(.wkbkTest, ModelNew:=sMdl1, ModelDefnDest:=defn_dest)
    
        'Check that SMDlType2 swapped into mdlDest region and cleared from tblImport sheet
        CheckTblImpHasType2Only Test
        CheckMdlDestVarNames Test, Split(",mdl_name,,,batch_size,use_premix", ",")
        .Update Test
    End With
End Function
'-----------------------------------------------------------------------------------------------
' SwapModels Procedure where ModelNew specified and no current mdlDest
' JDL 8/23/23
'
Function test_SwapModels3(wkbk, shtTests)
    Dim Test As New Tests
    Test.Populate Test, wkbk, shtTests, "test_SwapModels3"
    SetApplEnvir False, False, xlCalculationAutomatic
    
    'Test that populates test.valTest with True or False
    Dim mdlDest As Object: Set mdlDest = ExcelSteps.New_mdl

    With Test
    
        'Populate tblImport sheet only
        PopulateDashAndMdlImportSht Test

        'SwapModels with no ModelNew so mdlDest gets transferred to tblImport sheet only
        .Assert Test, mdlDest.SwapModels(.wkbkTest, ModelNew:=sMdl2, ModelDefnDest:=defn_dest)
    
        'Check that SMDlType2 swapped into mdlDest region and cleared from tblImport sheet
        CheckTblImpHasType1Only Test
        CheckMdlDestVarNames Test, Split(",mdl_name,,,n_sections,T_start,T_start_f,", ",")

        .Update Test
    End With
End Function
'-----------------------------------------------------------------------------------------------
'JDL 8/23/23
Sub CheckTblImpHasType1Only(Test)
    Dim rng As Range, rng2 As Range, tblImp_val As Object
    
    'Instance tblImp_val for validation
    InstanceValClass Test, tblImp_val:=tblImp_val
    With tblImp_val.wksht
        
        'Set test range for SMdlType2 (should be empty)
        Set rng = Intersect(Range(.Rows(2 + nrows_type1), .Rows(20)), tblImp_val.rngTable.EntireColumn)
        Test.TestRangeIsEmpty Test, rng
   
        'Set test range and expected vals for SMdlType1 (populated)
        Set rng2 = Intersect(Range(.Rows(2), .Rows(1 + nrows_type1)), tblImp_val.colrngMdlName)
        Test.TestRngVals Test, rng2, Split(Test.CreateLstRepeatVals(sMdl1, nrows_type1, ","), ",")

    End With
End Sub
'-----------------------------------------------------------------------------------------------
'JDL 8/23/23
Sub CheckTblImpHasType2Only(Test)
    Dim rng As Range, rng2 As Range, tblImp_val As Object
    
    'Instance tblImp_val for validation
    InstanceValClass Test, tblImp_val:=tblImp_val
    With tblImp_val.wksht
        
        'Check test range for SMdlType1 (should be empty)
        Set rng = Intersect(Range(.Rows(2 + nrows_type2), .Rows(20)), tblImp_val.rngTable.EntireColumn)
        Test.TestRangeIsEmpty Test, rng
   
        'Check test range and expected vals for SMdlType1 (populated)
        Set rng2 = Intersect(Range(.Rows(2), .Rows(1 + nrows_type1)), tblImp_val.colrngMdlName)
        Test.TestRngVals Test, rng2, Split(Test.CreateLstRepeatVals(sMdl2, nrows_type2, ","), ",")

    End With
End Sub
'-----------------------------------------------------------------------------------------------
'JDL 8/23/23
Sub CheckMdlDestVarNames(Test, aryExpect)
    Dim rng As Range, mdlDest_val As Object
    
    'Instance a separate mdlDest for validation
    InstanceValClass Test, mdlDest_val:=mdlDest_val

    'Check mdlDest variable names column (Should have SMdlType2 names)
    With mdlDest_val
        Set rng = Intersect(.rngrows, .colrngVarNames)
        Test.TestRngVals Test, rng, aryExpect
    End With

End Sub
'-----------------------------------------------------------------------------------------------
'JDL 8/23/23; Modified 10/24/24 for new default IsSetColRngs:=False
Sub InstanceValClass(Test, Optional mdlDest_val, Optional tblImp_val, Optional tblS_val)
    With Test
    If Not IsMissing(mdlDest_val) Then
        Set mdlDest_val = ExcelSteps.New_mdl
        .Assert Test, mdlDest_val.Provision(mdlDest_val, .wkbkTest, defn:=defn_dest)
    End If
    
    If Not IsMissing(tblImp_val) Then
        Set tblImp_val = ExcelSteps.New_tbl
        .Assert Test, tblImp_val.Provision(tblImp_val, .wkbkTest, False, shtTblImp, nCols:=10, IsSetColRngs:=True)
    End If
    
    If Not IsMissing(tblS_val) Then
        Set tblS_val = ExcelSteps.New_tbl
        .Assert Test, tblS_val.Provision(tblS_val, .wkbkTest, False, shtSteps)
    End If
    
    End With
End Sub
'-----------------------------------------------------------------------------------------------
' SwapModels Procedure where no ModelNew specified and SMdlType2 initially in mdlDest
' JDL 8/22/23
'
Function test_SwapModels2(wkbk, shtTests)
    Dim Test As New Tests
    Test.Populate Test, wkbk, shtTests, "test_SwapModels2"
    SetApplEnvir False, False, xlCalculationAutomatic
    
    'Test that populates test.valTest with True or False
    Dim tblImp As Object: Set tblImp = ExcelSteps.New_tbl
    Dim tbls As Object: Set tbls = ExcelSteps.New_tbl
    Dim mdlDest As Object: Set mdlDest = ExcelSteps.New_mdl

    With Test
    
        'Populate SMdlType2 into mdlDest
        PopulateSMdlType2ToMdlDest Test, mdlDest, tblImp, tbls

        'SwapModels with no ModelNew so mdlDest gets transferred to tblImport sheet only
        .Assert Test, mdlDest.SwapModels(.wkbkTest, ModelDefnDest:=defn_dest)
    
        'Check that model cleared from ExcelSteps and SMdl sheet
        CheckSMdlType2Cleared Test, tbls, mdlDest

        .Update Test
    End With
End Function
'-----------------------------------------------------------------------------------------------
' SwapModels Procedure where no ModelNew specified and No current mdlDest
' JDL 8/22/23
'
Function test_SwapModels1(wkbk, shtTests)
    Dim Test As New Tests
    Test.Populate Test, wkbk, shtTests, "test_SwapModels1"
    SetApplEnvir False, False, xlCalculationAutomatic
    
    'Test that populates test.valTest with True or False
    Dim tblImp As Object: Set tblImp = ExcelSteps.New_tbl
    Dim tbls As Object: Set tbls = ExcelSteps.New_tbl
    Dim mdlDest As Object: Set mdlDest = ExcelSteps.New_mdl

    With Test
    
        'Populate tblImport sheet only
        PopulateDashAndMdlImportSht Test

        'SwapModels with no ModelNew so mdlDest gets transferred to tblImport sheet only
        .Assert Test, mdlDest.SwapModels(.wkbkTest, ModelDefnDest:=defn_dest)
    
        'Check that no transfers happen
        .Assert Test, tblImp.Provision(tblImp, .wkbkTest, False, shtTblImp, nCols:=10)
        .Assert Test, tblImp.nRows = 9
        .Assert Test, tbls.Provision(tbls, .wkbkTest, False, sht:=shtSteps)
        .Assert Test, tbls.nRows = 1

        .Update Test
    End With
End Function

'-----------------------------------------------------------------------------------------------
' Initialize mdlDest, tblImp and tblS for Swap
' JDL 7/25/23
Sub test_InitSwapModels(wkbk, shtTests)
    Dim Test As New Tests
    Test.Populate Test, wkbk, shtTests, "test_InitSwapModels"
    SetApplEnvir False, False, xlCalculationAutomatic

    'Test that populates test.valTest
    Dim tblImp As Object: Set tblImp = ExcelSteps.New_tbl
    Dim tbls As Object: Set tbls = ExcelSteps.New_tbl
    Dim mdlDest As Object: Set mdlDest = ExcelSteps.New_mdl
    
    With Test
    
        'Populate model on sMdl and tbl on tblImport sheet
        PopulateDashAndMdlImportSht Test
    
        'Initialize Classes for the swap
        .Assert Test, mdlDest.InitSwapModels(mdlDest, tblImp, .wkbkTest, defn_dest)
        
        .Assert Test, mdlDest.cellHome.Address = "$I$10"
        .Assert Test, mdlDest.nRows = 1
        
        .Assert Test, tblImp.rowCur.Address = "$11:$11"
        .Assert Test, tblImp.nRows = 9
        
        .Assert Test, tbls.rowCur.Address = "$4:$4"
        .Assert Test, tbls.nRows = 1
        
        .Update Test
    End With
End Sub
'-----------------------------------------------------------------------------------------------
'SwapModel Setup - Create and Refresh a top (non-swapped) "Dashboard model
'-----------------------------------------------------------------------------------------------
'Populate Dashboard Scenario Model (Top, non-swapped Scenario Model)
'JDL 12/15/21
'
Sub test_PopulateDashMdl(wkbk, shtTests)
    Dim Test As New Tests
    Test.Populate Test, wkbk, shtTests, "test_PopulateDashMdl"
    SetApplEnvir False, False, xlCalculationAutomatic

    'Test that populates test.valTest
    Dim wksht As Worksheet: Set wksht = wkbk.Sheets("SMdl")
    Dim mdlDash As Object
    
    'Populate model on sMdl and check
    PopulateDashMdl Test, mdlDash
    With Test
        .Assert Test, (wksht.Cells(4, 9) = "Dashboard")
        .Assert Test, (wksht.Cells(6, 14) = "xxx")
        
        .Update Test
    End With
End Sub
'-----------------------------------------------------------------------------------------------
'SwapModel Setup - Populate and set up tblImport sheet
'-----------------------------------------------------------------------------------------------
'SwapModel Setup - Populate tblImport sheet's table for testing
'JDL 7/17/23
'
Sub test_PopulateMdlImport(wkbk, shtTests)
    Dim Test As New Tests
    Test.Populate Test, wkbk, shtTests, "test_PopulateMdlImport"
    SetApplEnvir False, False, xlCalculationAutomatic

    'Test that populates test.valTest
    Dim tblImp As Object, i As Integer
        
    With Test
        PopulateMdlImportType1AndType2 Test, tblImp
        For i = 2 To nrows_type1 + 1
            .Assert Test, tblImp.wksht.Cells(i, 1) = sMdl1
        Next i
        For i = 2 + nrows_type1 To 1 + nrows_both
            .Assert Test, tblImp.wksht.Cells(i, 1) = sMdl2
        Next i
        
        'Check that Populate sub leaves blank cells blank
        .Assert Test, Len(tblImp.wksht.Cells(10, 10)) = 0
        Test.Update Test
    End With
End Sub
'-----------------------------------------------------------------------------------------------
'SwapModel Setup - Format the tblImport rows/columns sheet table
'JDL 12/15/21   Modified 7/14/23
'
Sub test_FormatMdlImport(wkbk, shtTests)
    Dim Test As New Tests
    Test.Populate Test, wkbk, shtTests, "test_FormatMdlImport"
    SetApplEnvir False, False, xlCalculationAutomatic

    'Test that populates test.valTest
    Dim tblImp As Object: Set tblImp = ExcelSteps.New_tbl
        
    'DeleteSheet wkbk, shtTblImp
    wkbk.Sheets(shtTblImp).Cells.Clear
    With tblImp
        Test.Assert Test, .Provision(tblImp, wkbk, False, shtTblImp, nCols:=10, IsSetColRngs:=True)
        .FormatMdlImport tblImp
        
        Test.Assert Test, (.colrngNumFmt.NumberFormat = "@")
        Test.Assert Test, (.colrngStrInput.NumberFormat = "@")
        Test.Assert Test, (wkbk.Sheets(shtTblImp).Cells(1, 2) = "Grp")
        Test.Update Test
    End With
End Sub

'-----------------------------------------------------------------------------------------------
'-----------------------------------------------------------------------------------------------
' TransferToTblImport Procedure - sub-procedure of SwapModels
'-----------------------------------------------------------------------------------------------
'-----------------------------------------------------------------------------------------------
' Procedure - Transfer model from mdlDest Scenario Model region to tblImport sheet rows/cols
' JDL 8/22/23
'
Function test_TransferToTblImport(wkbk, shtTests)
    Dim Test As New Tests
    Test.Populate Test, wkbk, shtTests, "test_TransferToTblImport"
    SetApplEnvir False, False, xlCalculationAutomatic
    
    'Test that populates test.valTest with True or False
    Dim tblImp As Object: Set tblImp = ExcelSteps.New_tbl
    Dim tbls As Object: Set tbls = ExcelSteps.New_tbl
    Dim mdlDest As Object: Set mdlDest = ExcelSteps.New_mdl

    With Test
        PopulateSMdlType2ToMdlDest Test, mdlDest, tblImp, tbls
        .Assert Test, mdlDest.TransferToTblImport(mdlDest, tblImp, tbls)
        
        'Check that model cleared from ExcelSteps and SMdl sheet
        CheckSMdlType2Cleared Test, tbls, mdlDest
        
        .Update Test
    End With
End Function
Sub CheckSMdlType2Cleared(Test, ByVal tbls, ByVal mdlDest)
    Dim rng As Range
    With Test
        'Check that model was cleared from ExcelSteps
        .Assert Test, tbls.rngrows.Address = "$2:$3"
        .Assert Test, IsEmpty(tbls.wksht.Cells(4, 2))
        
        'Check that model was cleared from SMdl sheet
        Set rng = Range(mdlDest.wksht.Rows(11), mdlDest.wksht.Rows(19))
        Set rng = Intersect(rng, mdlDest.colrngVarNames)
        .TestRangeIsEmpty Test, rng
    End With
End Sub
'-----------------------------------------------------------------------------------------------
' Delete trailing blank rows, if any, from tblImp
'
' JDL 8/22/23
'
Function test_DeleteTblImpTrailingBlankRows(wkbk, shtTests)
    Dim Test As New Tests
    Test.Populate Test, wkbk, shtTests, "test_DeleteTblImpTrailingBlankRows"
    SetApplEnvir False, False, xlCalculationAutomatic
    
    'Test that populates test.valTest with True or False
    Dim tblImp As Object: Set tblImp = ExcelSteps.New_tbl
    Dim tbls As Object: Set tbls = ExcelSteps.New_tbl
    Dim mdlDest As Object: Set mdlDest = ExcelSteps.New_mdl
    Dim R_MI As Object: Set R_MI = ExcelSteps.New_mdlImportRow
        Dim ModelPrev As String, aryExpect As Variant, defn_extra_rows As String, rng As Range

    With Test
        PopulateSMdlType2ToMdlDest Test, mdlDest, tblImp, tbls
        
        'Transfer mdlDest (SMdlType2) to tblImport sheet
        
        defn_extra_rows = "SMdl:10,9:10:T:T:T:F:T:SMdlDest"
        Set mdlDest = ExcelSteps.New_mdl
        .Assert Test, mdlDest.InitSwapModels(mdlDest, tblImp, tbls, .wkbkTest, defn_extra_rows)
        .Assert Test, mdlDest.rngrows.Address = "$10:$19"
        .Assert Test, mdlDest.ReadModelName(mdlDest, tbls, ModelPrev)
        .Assert Test, mdlDest.TransferMdlDestRows(mdlDest, tblImp, tbls, ModelPrev)
        
        'Check for trailing blanks
        Set rng = Range(tblImp.wksht.Rows(11), tblImp.wksht.Rows(13))
        Set rng = Intersect(rng, tblImp.colrngVarName)
        .TestRngVals Test, rng, Array("<blank>", "<blank>", "<blank>")
        .Assert Test, tblImp.rngrows.Address = "$2:$13"
        
        'Check that trailing blanks get deleted
        .Assert Test, mdlDest.DeleteTblImpTrailingBlankRows(tblImp)
        .TestRangeIsEmpty Test, rng
        .Assert Test, tblImp.rngrows.Address = "$2:$10"
        
        .Update Test
    End With
End Function
'-----------------------------------------------------------------------------------------------
' Transfer mdlDest Scenario Model rows to tblImport rows/columns table
'
' JDL 8/22/23
'
Function test_TransferMdlDestRows(wkbk, shtTests)
    Dim Test As New Tests
    Test.Populate Test, wkbk, shtTests, "test_TransferTblImportRows"
    SetApplEnvir False, False, xlCalculationAutomatic
    
    'Test that populates test.valTest
    Dim tblImp As Object: Set tblImp = ExcelSteps.New_tbl
    Dim tbls As Object: Set tbls = ExcelSteps.New_tbl
    Dim mdlDest As Object: Set mdlDest = ExcelSteps.New_mdl
    Dim ModelPrev As String, aryExpect As Variant
    
    With Test
        PopulateSMdlType2ToMdlDest Test, mdlDest, tblImp, tbls
        
        'Transfer mdlDest (SMdlType2) to tblImport sheet
        Set mdlDest = ExcelSteps.New_mdl
        .Assert Test, mdlDest.InitSwapModels(mdlDest, tblImp, tbls, .wkbkTest, defn_dest)
        .Assert Test, mdlDest.ReadModelName(mdlDest, tbls, ModelPrev)
        .Assert Test, mdlDest.TransferMdlDestRows(mdlDest, tblImp, tbls, ModelPrev)
        
        'Repeat tblImport value checks from tests_mdlImportRow (test_ToTblWriteRow)
        aryExpect = Split("SMdlType2,Setup,,Configuration Name (used by program),mdl_name,,,,,SMdlType2", ",")
        CheckWrittenRowValues Test, tblImp, 6, aryExpect
        aryExpect = Split("SMdlType2,Setup,,,<blank>,,,,,,", ",")
        CheckWrittenRowValues Test, tblImp, 7, aryExpect
        aryExpect = Split("SMdlType2,Other Plant Configuration,,No. Sections,n_sections,,,,,4", ",")
        CheckWrittenRowValues Test, tblImp, 8, aryExpect
        aryExpect = Split("SMdlType2,Other Plant Configuration,,Start Temperature (Celsius),T_start,C,,,,40", ",")
        CheckWrittenRowValues Test, tblImp, 9, aryExpect
        aryExpect = Split("SMdlType2,Other Plant Configuration,,Start Temperature (Fahrenheit),T_start_f,F,0.0,=(T_start * 9/5) + 32,,", ",")
        CheckWrittenRowValues Test, tblImp, 10, aryExpect, IsConvertNumerics:=False
        
        .Update Test
    End With
End Function
'-----------------------------------------------------------------------------------------------
' Read name of existing model (mdl_name variable value) in mdlDest region
' Modified 11/6/24 add IsSetColRngs Provision arg
'
Function test_TblImportDeleteModel(wkbk, shtTests)
    Dim Test As New Tests
    Test.Populate Test, wkbk, shtTests, "test_TblImportDeleteModel"
    SetApplEnvir False, False, xlCalculationAutomatic
    
    'Test that populates test.valTest
    Dim tblImp As Object: Set tblImp = ExcelSteps.New_tbl
    Dim tbls As Object: Set tbls = ExcelSteps.New_tbl
    Dim mdlDest As Object: Set mdlDest = ExcelSteps.New_mdl
    
    With Test
        PopulateSMdlType2ToMdlDest Test, mdlDest, tblImp, tbls
        
        'Add Type 2 model back to tblImport sheet and Provision to check number of rows
        PopulateMdlType2 tblImp
        .Assert Test, tblImp.Provision(tblImp, .wkbkTest, False, sht:=shtTblImp, nCols:=10, IsSetColRngs:=True)

        .Assert Test, tblImp.rngrows.Address = "$2:$10"

        'Delete SMdlType2 from tblImport sheet
        .Assert Test, mdlDest.TblImportDeleteModel(tblImp, sMdl2)
        .Assert Test, tblImp.rngrows.Address = "$2:$5"
        
        .Update Test
    End With
End Function
'-----------------------------------------------------------------------------------------------
' Read name of existing model (mdl_name variable value) in mdlDest region
'
Function test_ReadModelName(wkbk, shtTests)
    Dim Test As New Tests
    Test.Populate Test, wkbk, shtTests, "test_ReadModelName"
    SetApplEnvir False, False, xlCalculationAutomatic
    
    'Test that populates test.valTest
    Dim tblImp As Object: Set tblImp = ExcelSteps.New_tbl
    Dim tbls As Object: Set tbls = ExcelSteps.New_tbl
    Dim mdlDest As Object: Set mdlDest = ExcelSteps.New_mdl
    Dim ModelPrev As String
    
    With Test
        PopulateSMdlType2ToMdlDest Test, mdlDest, tblImp, tbls
        .Assert Test, mdlDest.ReadModelName(mdlDest, tbls, ModelPrev)
        .Assert Test, ModelPrev = sMdl2
        .Assert Test, mdlDest.rngStepsVars.Address = "$4:$4"

        .Update Test
    End With
End Function
'-----------------------------------------------------------------------------------------------
' Helper sub to load SMdlType2 into mdlDest as starting point for TransferMdlDestProcedure
' JDL 7/27/23
'
Sub PopulateSMdlType2ToMdlDest(Test, mdlDest, tblImp, tbls)
    With Test
        'Populate model on sMdl and tbl on tblImport sheet
        PopulateDashAndMdlImportSht Test
    
        'Transfer SMdlType2 to mdlDest region
        
        'yyy instead of defn_dest
        'Dim defn As String
        'defn = "SMdl:10,9:0:T:T:T:F:T:SMdlType2"
        
        .Assert Test, mdlDest.InitSwapModels(mdlDest, tblImp, tbls, .wkbkTest, defn_dest)
        .Assert Test, mdlDest.TransferToMdlDest(mdlDest, tblImp, tbls, sMdl2, defn_dest)
        
        'Reset model name as it would be w/o TransferToMdlDest mod
        mdlDest.MdlName = "SMdlDest"
    End With
End Sub
'-----------------------------------------------------------------------------------------------
'-----------------------------------------------------------------------------------------------
' TransferToMdlDest Procedure - sub-procedure of SwapModels
'-----------------------------------------------------------------------------------------------
'-----------------------------------------------------------------------------------------------
' Transfer Type 1 model to mdlDest (TransferToMdlDest assumes prev mdlDest pre-transferred)
' JDL 7/25/23
'
Sub test_TransferToMdlDest(wkbk, shtTests)
    Dim Test As New Tests
    Test.Populate Test, wkbk, shtTests, "test_TransferToMdlDest"
    SetApplEnvir False, False, xlCalculationAutomatic

    'Test that populates test.valTest
    Dim tblImp As Object: Set tblImp = ExcelSteps.New_tbl
    Dim tbls As Object: Set tbls = ExcelSteps.New_tbl
    Dim mdlDest As Object: Set mdlDest = ExcelSteps.New_mdl
    
    With Test
    
        'Populate model on sMdl and tbl on tblImport sheet
        PopulateDashAndMdlImportSht Test
    
        'Initialize Classes for the swap and Transfer the SMdlType2 to mdlDest region
        .Assert Test, mdlDest.InitSwapModels(mdlDest, tblImp, tbls, .wkbkTest, defn_dest)
        .Assert Test, mdlDest.TransferToMdlDest(mdlDest, tblImp, tbls, sMdl2, defn_dest)
        CheckPostTransferFormatting Test, mdlDest

        .Update Test
    End With
End Sub
'-----------------------------------------------------------------------------------------------
' Init transferring a model from tblImport sheet to Scenario Model
' JDL 7/25/23
'
Sub test_InitTransferToMdl(wkbk, shtTests)
    Dim Test As New Tests
    Test.Populate Test, wkbk, shtTests, "test_InitTransferToMdl"
    SetApplEnvir False, False, xlCalculationAutomatic

    'Test that populates test.valTest
    Dim tblImp As Object: Set tblImp = ExcelSteps.New_tbl
    Dim tbls As Object: Set tbls = ExcelSteps.New_tbl
    Dim mdlDest As Object: Set mdlDest = ExcelSteps.New_mdl
    Dim R_MI As Object: Set R_MI = ExcelSteps.New_mdlImportRow
    Dim rngMdl As Range
    
    With Test
    
        'Populate model on sMdl and tbl on tblImport sheet
        PopulateDashAndMdlImportSht Test
    
        'Initialize Classes and Init the Transfer
        .Assert Test, mdlDest.InitSwapModels(mdlDest, tblImp, tbls, .wkbkTest, defn_dest)
        .Assert Test, mdlDest.InitTransferToMdl(mdlDest, tblImp, sMdl1)
        
        'Check initialization of tblImport for Transfer
        .Assert Test, tblImp.rowCur.Row = 2
        .Assert Test, tblImp.rngRowsPopulated.Address = "$2:$5"
        
        'Check mdlDest got cleared
        .Assert Test, IsEmpty(mdlDest.cellHome)
        
        .Update Test
    End With
End Sub
'-----------------------------------------------------------------------------------------------
' Transfer a model's rows from tblImport sheet to mdlDest Scenario Model - SMdlType2
' JDL 7/25/23
'
Sub test_TransferTblImportRows1(wkbk, shtTests)
    Dim Test As New Tests
    Test.Populate Test, wkbk, shtTests, "test_TransferTblImportRows1"
    SetApplEnvir False, False, xlCalculationAutomatic

    'Test that populates test.valTest
    Dim tblImp As Object: Set tblImp = ExcelSteps.New_tbl
    Dim tbls As Object: Set tbls = ExcelSteps.New_tbl
    Dim mdlDest As Object: Set mdlDest = ExcelSteps.New_mdl
    Dim R_MI As Object: Set R_MI = ExcelSteps.New_mdlImportRow
    Dim aryExpect() As Variant, ncols_mdl As Integer, nrows_mdl As Integer
    
    'Helper sub to initialize and call Transfer method
    PopulateAndTransferTblImportRows Test, mdlDest, tblImp, tbls, sMdl1, R_MI
        
    'lbound 0 corresponding to Offset from cellHome
    ncols_mdl = 7
    nrows_mdl = 6
    ReDim aryExpect(0 To ncols_mdl - 1)
    aryExpect(0) = Split("Setup,,,Batch Plant Configuration,,", ",")
    aryExpect(1) = Split(",Configuration Name (used by program),,,Batch Size,Use Premix", ",")
    aryExpect(2) = Split(",mdl_name,,,batch_size,use_premix", ",")
    aryExpect(3) = Split(",,,,kg,kg", ",")
    aryExpect(4) = Split(",,,,,", ",")
    aryExpect(5) = Split(",SMdlType1,,,10000,True", ",")
    aryExpect(6) = Split(",,,,,", ",")
        
    'Check Lite model column ranges versus expected
    CheckMdlDestVsExpectedVals Test, mdlDest, aryExpect, ncols_mdl, nrows_mdl
    
    Test.Update Test
End Sub
'-----------------------------------------------------------------------------------------------
' Transfer a model's rows from tblImport sheet to mdlDest Scenario Model - SMdlType2
' JDL 7/27/23
'
Sub test_TransferTblImportRows2(wkbk, shtTests)
    Dim Test As New Tests
    Test.Populate Test, wkbk, shtTests, "test_TransferTblImportRows2"
    SetApplEnvir False, False, xlCalculationAutomatic

    'Test that populates test.valTest
    Dim tblImp As Object: Set tblImp = ExcelSteps.New_tbl
    Dim tbls As Object: Set tbls = ExcelSteps.New_tbl
    Dim mdlDest As Object: Set mdlDest = ExcelSteps.New_mdl
    Dim R_MI As Object: Set R_MI = ExcelSteps.New_mdlImportRow
    Dim aryExpect() As Variant, aryExpectSteps As Variant
    Dim rngMdl As Range, ncols_mdl As Integer, nrows_mdl As Integer
    
    'Helper sub to initialize and call Transfer method
    PopulateAndTransferTblImportRows Test, mdlDest, tblImp, tbls, sMdl2, R_MI
   
    'Check transfer to ExcelSteps (can't include numerical format string in TestRngVals)
    aryExpectSteps = Split("SMdlType2,T_start_f,Col_Insert,=(T_start * 9/5) + 32,,,,", ",")
    With Test.wkbkTest.Sheets(shtSteps)
        Test.TestRngVals Test, Range(.Cells(4, 1), .Cells(4, 7)), aryExpectSteps
        Test.Assert Test, .Cells(4, 8) = "0.0"
    End With

    'lbound 0 corresponding to Offset from cellHome
    ncols_mdl = 7
    nrows_mdl = 7
    ReDim aryExpect(0 To ncols_mdl - 1)
    aryExpect(0) = Split("Setup,,,Other Plant Configuration,,,", ",")
    aryExpect(1) = Split(",Configuration Name (used by program),,,No. Sections," & _
                        "Start Temperature (Celsius),Start Temperature (Fahrenheit)", ",")
    aryExpect(2) = Split(",mdl_name,,,n_sections,T_start,T_start_f", ",")
    aryExpect(3) = Split(",,,,,C,F", ",")
    aryExpect(4) = Split(",,,,,,", ",")
    aryExpect(5) = Split(",SMdlType2,,,4,40,", ",")
    aryExpect(6) = Split(",,,,,,", ",")

    'Check Lite model column ranges versus expected
    CheckMdlDestVsExpectedVals Test, mdlDest, aryExpect, ncols_mdl, nrows_mdl
    
    Test.Update Test
End Sub
'-----------------------------------------------------------------------------------------------
' Helper sub to initialize and call Transfer method
' JDL 7/25/23
'
Sub PopulateAndTransferTblImportRows(Test, mdlDest, tblImp, tbls, ModelNew, R_MI)
    With Test
    
        'Populate model on sMdl and tbl on tblImport sheet
        PopulateDashAndMdlImportSht Test
    
        'Initialize Classes and Init the Transfer
        .Assert Test, mdlDest.InitSwapModels(mdlDest, tblImp, tbls, .wkbkTest, defn_dest)
        .Assert Test, mdlDest.InitTransferToMdl(mdlDest, tblImp, ModelNew)
        .Assert Test, mdlDest.TransferTblImportRows(R_MI, mdlDest, tblImp, tbls)
        
        'Check cellHome
        .Assert Test, mdlDest.cellHome.Address = "$I$10"
    End With
End Sub
'-----------------------------------------------------------------------------------------------
' Check Lite model column ranges versus expected
' JDL 7/27/23
'
Sub CheckMdlDestVsExpectedVals(Test, mdlDest, aryExpect, ncols_mdl, nrows_mdl)
    Dim i As Integer
    With mdlDest.cellHome
        For i = 0 To ncols_mdl - 1
            Test.TestRngVals Test, Range(.Offset(0, i), .Offset(nrows_mdl - 1, i)), aryExpect(i)
        Next i
    End With
End Sub
'-----------------------------------------------------------------------------------------------
' Clear transferred model from tblImport sheet and Refresh mdlDest post-transfer
Sub test_ResetPostTransfer(wkbk, shtTests)
    Dim Test As New Tests
    Test.Populate Test, wkbk, shtTests, "test_ResetPostTransfer"
    SetApplEnvir False, False, xlCalculationAutomatic

    'Test that populates test.valTest
    Dim tblImp As Object: Set tblImp = ExcelSteps.New_tbl
    Dim tbls As Object: Set tbls = ExcelSteps.New_tbl
    Dim mdlDest As Object: Set mdlDest = ExcelSteps.New_mdl
    Dim R_MI As Object: Set R_MI = ExcelSteps.New_mdlImportRow
    Dim aryExpect() As Variant
    
    With Test

        'Helper sub to initialize and call Transfer method
        PopulateAndTransferTblImportRows Test, mdlDest, tblImp, tbls, sMdl2, R_MI
        
        'Re-initialize/refresh mdlDest; delete rngModel rows from tblImport table
        .Assert Test, mdlDest.ResetPostTransfer(mdlDest, tblImp.rngRowsPopulated, sMdl2, defn_dest)
        
        'Check Refresh results
        CheckPostTransferFormatting Test, mdlDest
        
        .Update Test
    End With
End Sub
'-----------------------------------------------------------------------------------------------
' Check SMdlType2 formula-calculated value and model formatting
' JDL 7/27/23
'
Sub CheckPostTransferFormatting(Test, mdlDest)
    With mdlDest
        'Check calculated cell value and formatting
        Test.Assert Test, .wksht.Cells(16, 14) = 104#
        Test.CkStyleMatch Test, .wksht.Cells(16, 14), "Calculation"

        'Check that formatting happened
        Test.CkStyleMatch Test, Intersect(.rngPopRows.EntireRow, .colrngVarNames), "Note"
        Test.CkStyleMatch Test, Intersect(.rngPopRows.EntireRow, .colrngUnits), "Note"
    End With
End Sub
