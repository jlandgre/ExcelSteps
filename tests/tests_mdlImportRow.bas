Attribute VB_Name = "tests_mdlImportRow"
' Version 1/28/26 refactor to use procs framework
' 12/13/24 tblRowsCols mods
Option Explicit
'-----------------------------------------------------------------------------------------------
' Test suite for mdlImportRow Class
' JDL 8/23/23; Refactored 1/28/26
'
Sub TestDriver_mdlImportRow()
    Dim procs As New Procedures, AllEnabled As Boolean
    
    With procs
        .Init procs, ThisWorkbook, "tests_mdlImportRow", "tests_mdlImportRow"
        SetApplEnvir False, False, xlCalculationAutomatic
        
        'Enable testing of all or individual procedures
        AllEnabled = False
        .mdlImportRow.Enabled = True
    End With
    
    'Setup procedure group
    With procs.mdlImportRow
        If .Enabled Or AllEnabled Then
            procs.curProcedure = .name
            test_Init procs
            test_ReadMdlDestRow procs
            test_ReadStepsRow procs
            test_SetBooleanFlags procs
            test_ToTblWriteRow procs
        End If
    End With
    
    procs.EvalOverall procs
    SetApplEnvir True, True, xlCalculationAutomatic
End Sub

'--------------------------------------------------------------------------------------
'--------------------------------------------------------------------------------------
' procs.mdlImportRow
'--------------------------------------------------------------------------------------
'--------------------------------------------------------------------------------------
' Set mdlImportRow attributes for previous Group and Subgroup; Re-initialize Class
' JDL 8/12/23; refactored 1/28/26
'
Sub test_Init(procs)
    Dim tst As New Test: tst.Init tst, "test_Init"
    Dim R_MI As Object: Set R_MI = ExcelSteps.New_mdlImportRow

    With tst
        R_MI.Grp = "xxx"
        R_MI.SubGrp = "yyy"
        R_MI.Model = "dummy"
        
        'Check that class was re-initialized and Prev attributes got set
        .Assert tst, R_MI.Init(R_MI)
        .Assert tst, R_MI.GrpPrev = "xxx"
        .Assert tst, R_MI.SubgrpPrev = "yyy"
        .Assert tst, R_MI.Model = "dummy"
        .Update tst, procs
    End With
End Sub

'-----------------------------------------------------------------------------------------------
' Read a mdlDest row into tblImportRow attributes
' JDL 8/12/23; refactored 1/28/26
'
Sub test_ReadMdlDestRow(procs)
    Dim tst As New Test: tst.Init tst, "test_ReadMdlDestRow"
    Dim tblImp As Object: Set tblImp = ExcelSteps.New_tbl
    Dim tbls As Object: Set tbls = ExcelSteps.New_tbl
    Dim mdlDest As Object: Set mdlDest = ExcelSteps.New_mdl
    Dim R_MI As Object: Set R_MI = ExcelSteps.New_mdlImportRow
    Dim aryExpect As Variant

    'Populate the Scenario model to mdlDest
    PopulateForMdlImportRow tst, mdlDest, tblImp

    With tst
    
        'First row in model (Group name only)
        Set mdlDest.rowCur = mdlDest.cellHome.EntireRow
        .Assert tst, R_MI.Init(R_MI)
        .Assert tst, R_MI.ReadMdlDestRow(R_MI, mdlDest)
        .Assert tst, R_MI.Grp = "Setup"
        .Assert tst, R_MI.VarName = ""
        .Assert tst, mdlDest.rowCur.Address = mdlDest.cellHome.Offset(1, 0).EntireRow.Address
        
        'A variable's row (Setup Group)
        .Assert tst, R_MI.Init(R_MI)
        .Assert tst, R_MI.ReadMdlDestRow(R_MI, mdlDest)
        aryExpect = Split("Setup,mdl_name,Configuration Name (used by program),SMdlType2", ",")
        .TestAryVals tst, Array(R_MI.Grp, R_MI.VarName, R_MI.Desc, R_MI.Value), aryExpect
        
        'A blank row (Still Setup Group
        .Assert tst, R_MI.Init(R_MI)
        .Assert tst, R_MI.ReadMdlDestRow(R_MI, mdlDest)
        .Assert tst, R_MI.IsBlankRow = True
        .Assert tst, R_MI.Grp = "Setup"

        'A new group
        .Assert tst, R_MI.Init(R_MI)
        .Assert tst, R_MI.ReadMdlDestRow(R_MI, mdlDest)
        .Assert tst, R_MI.Grp = "Other Plant Configuration"
        .Assert tst, R_MI.VarName = ""
        
        'A formula-containing row (Len > 0 triggers reading from Steps table)
        Set mdlDest.rowCur = mdlDest.rowCur.Offset(2, 0)
        .Assert tst, R_MI.Init(R_MI)
        .Assert tst, R_MI.ReadMdlDestRow(R_MI, mdlDest)
        .Assert tst, R_MI.VarName = "T_start_f"
        .Assert tst, Len(R_MI.Value) > 0
    
        .Update tst, procs
    End With
End Sub

'-----------------------------------------------------------------------------------------------
' Read a mdlDest row's attributes from tblS ExcelSteps table
' JDL 8/12/23; refactored 1/28/26
'
Sub test_ReadStepsRow(procs)
    Dim tst As New Test: tst.Init tst, "test_ReadStepsRow"
    Dim tblImp As Object: Set tblImp = ExcelSteps.New_tbl
    Dim tbls As Object: Set tbls = ExcelSteps.New_tbl
    Dim mdlDest As Object: Set mdlDest = ExcelSteps.New_mdl
    Dim R_MI As Object: Set R_MI = ExcelSteps.New_mdlImportRow

    'Populate the Scenario model to mdlDest
    PopulateForMdlImportRow tst, mdlDest, tblImp

    With tst
        Set mdlDest.rowCur = mdlDest.cellHome.Offset(6, 0).EntireRow
        .Assert tst, R_MI.Init(R_MI)
        .Assert tst, R_MI.ReadMdlDestRow(R_MI, mdlDest)
        .Assert tst, R_MI.VarName = "T_start_f"
        .Assert tst, Len(R_MI.Value) > 0
        
        'Read the variable's params from ExcelSteps sheet
        .Assert tst, R_MI.ReadStepsRow(R_MI, mdlDest.tblSteps, mdlDest.rngStepsVars)
        .Assert tst, R_MI.StrInput = "=(T_start * 9/5) + 32"
        .Assert tst, R_MI.NumFmt = "0.0"
        .Update tst, procs
    End With
End Sub
'-----------------------------------------------------------------------------------------------
' Set Boolean flags describing a newly-read mdlDest row
'
Sub test_SetBooleanFlags(procs)
    Dim tst As New Test: tst.Init tst, "test_SetBooleanFlags"
    Dim tblImp As Object: Set tblImp = ExcelSteps.New_tbl
    Dim tbls As Object: Set tbls = ExcelSteps.New_tbl
    Dim mdlDest As Object: Set mdlDest = ExcelSteps.New_mdl
    Dim R_MI As Object: Set R_MI = ExcelSteps.New_mdlImportRow

    'Populate the Scenario model to mdlDest
    PopulateForMdlImportRow tst, mdlDest, tblImp

    With tst
            
        'First row in model (Group name only)
        Set mdlDest.rowCur = mdlDest.cellHome.EntireRow
        SetRowBooleanFlags tst, R_MI, mdlDest
        .TestAryVals tst, Array(R_MI.IsNewGrp, R_MI.HasStepsRow, R_MI.IsBlankRow), _
            Array(True, False, False)
        
        'A variable's row (Setup Group)
        SetRowBooleanFlags tst, R_MI, mdlDest
        .TestAryVals tst, Array(R_MI.IsNewGrp, R_MI.HasStepsRow, R_MI.IsBlankRow), _
            Array(False, False, False)
        
        'A blank row (Still Setup Group)
        SetRowBooleanFlags tst, R_MI, mdlDest
        .TestAryVals tst, Array(R_MI.IsNewGrp, R_MI.HasStepsRow, R_MI.IsBlankRow), _
            Array(False, False, True)
            
        'A new group
        SetRowBooleanFlags tst, R_MI, mdlDest
        .TestAryVals tst, Array(R_MI.IsNewGrp, R_MI.HasStepsRow, R_MI.IsBlankRow), _
            Array(True, False, False)
        
        'A formula-containing row
        Set mdlDest.rowCur = mdlDest.rowCur.Offset(2, 0)
        SetRowBooleanFlags tst, R_MI, mdlDest
        .TestAryVals tst, Array(R_MI.IsNewGrp, R_MI.HasStepsRow, R_MI.IsBlankRow, _
                R_MI.HasFormula, R_MI.HasNumFmt), Array(False, True, False, True, True)
            
        .Update tst, procs
    End With
End Sub
Sub SetRowBooleanFlags(tst, R_MI, mdlDest)
    With tst
        .Assert tst, R_MI.Init(R_MI)
        .Assert tst, R_MI.ReadMdlDestRow(R_MI, mdlDest)
        .Assert tst, R_MI.ReadStepsRow(R_MI, mdlDest.tblSteps, mdlDest.rngStepsVars)
        .Assert tst, R_MI.SetBooleanFlags(R_MI)
    End With
End Sub
'-----------------------------------------------------------------------------------------------
' Write a mdlDest row to the tblImp rows/cols table on mdlImport sheet
' JDL 8/12/23; refactored 1/28/26
'
Sub test_ToTblWriteRow(procs)
    Dim tst As New Test: tst.Init tst, "test_ToTblWriteRow"
    Dim tblImp As Object: Set tblImp = ExcelSteps.New_tbl
    Dim tbls As Object: Set tbls = ExcelSteps.New_tbl
    Dim mdlDest As Object: Set mdlDest = ExcelSteps.New_mdl
    Dim R_MI As Object: Set R_MI = ExcelSteps.New_mdlImportRow
    Dim aryExpect As Variant, rng As Range

    'Populate the Scenario model to mdlDest
    PopulateForMdlImportRow tst, mdlDest, tblImp
    'R_MI.Model = "SMdlType2"
    mdlDest.MdlName = "SMdlType2"

    With tst
            
        'First row in model (Group name only - no writing)
        Set mdlDest.rowCur = mdlDest.cellHome.EntireRow
        SetRowBooleanFlags tst, R_MI, mdlDest
        .Assert tst, R_MI.ToTblWriteRow(R_MI, mdlDest, tblImp)
        .Assert tst, tblImp.rowCur.Row = 6
        
        'A variable's row (Group name "Setup")
        SetRowBooleanFlags tst, R_MI, mdlDest
        .Assert tst, R_MI.ToTblWriteRow(R_MI, mdlDest, tblImp)
        aryExpect = Split("SMdlType2,Setup,,Configuration Name (used by program),mdl_name,,,,,SMdlType2", ",")
        CheckWrittenRowValues tst, tblImp, 6, aryExpect
        
        'A blank row (Still Setup Group)
        SetRowBooleanFlags tst, R_MI, mdlDest
        .Assert tst, R_MI.ToTblWriteRow(R_MI, mdlDest, tblImp)
        aryExpect = Split("SMdlType2,Setup,,,<blank>,,,,,,", ",")
        CheckWrittenRowValues tst, tblImp, 7, aryExpect
        
        'New Group (Group name only - no writing)
        SetRowBooleanFlags tst, R_MI, mdlDest
        .Assert tst, R_MI.ToTblWriteRow(R_MI, mdlDest, tblImp)
        .Assert tst, tblImp.rowCur.Row = 8
        
        'Numeric variable - tblImp Row 8
        SetRowBooleanFlags tst, R_MI, mdlDest
        .Assert tst, R_MI.ToTblWriteRow(R_MI, mdlDest, tblImp)
        aryExpect = Split("SMdlType2,Other Plant Configuration,,No. Sections,n_sections,,,,,4", ",")
        CheckWrittenRowValues tst, tblImp, 8, aryExpect, True
        
        'Numeric variable - tblImp Row 9
        SetRowBooleanFlags tst, R_MI, mdlDest
        .Assert tst, R_MI.ToTblWriteRow(R_MI, mdlDest, tblImp)
        aryExpect = Split("SMdlType2,Other Plant Configuration,,Start Temperature (Celsius),T_start,C,,,,40", ",")
        CheckWrittenRowValues tst, tblImp, 9, aryExpect, True
        
        'Variable with formula and number format - tblImp Row 10
        SetRowBooleanFlags tst, R_MI, mdlDest
        .Assert tst, R_MI.ToTblWriteRow(R_MI, mdlDest, tblImp)
        
        'Need to individually check 0.0 number format string - entered as 0 in aryExpect
        aryExpect = Split("SMdlType2,Other Plant Configuration,,Start Temperature (Fahrenheit),T_start_f,F,0.0,=(T_start * 9/5) + 32,,", ",")
        CheckWrittenRowValues tst, tblImp, 10, aryExpect, IsConvertNumerics:=False
        
        .Update tst, procs
    End With
End Sub
Sub CheckWrittenRowValues(tst, tblImp, iRow, aryExpect, Optional IsConvertNumerics = True)
    Dim rng As Range, IsConvertNumericsLocal As Boolean

    'Overtly set ConvertNumerics to avoid VBA bug that ignores optional args in called subs
    IsConvertNumericsLocal = IsConvertNumerics


    
    With tblImp.wkbk.Sheets(tblImp.sht)
        Set rng = Range(.Cells(iRow, 1), .Cells(iRow, 10))
    End With
    tst.TestRngVals tst, rng, aryExpect, ConvertNumerics:=IsConvertNumericsLocal

End Sub




