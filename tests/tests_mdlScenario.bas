Attribute VB_Name = "tests_mdlScenario"
Option Explicit
'Version 10/29/25 Refactor for .Provision + .Refresh performance
'-----------------------------------------------------------------------------------------------
'Definition for testing parsing - non-default model with $I$4 cellHome
Public Const defn_val As String = "SMdl:4,9:0:T:F:T:T:T"
Public Const defn_val_mdlName As String = "SMdl:4,9:0:T:T:T:F:T:SMdlType2"
'Booleans: IsCalc, IsSuppHeader, IsRngNames, IsMdlNmPrefix, IsLiteModel
Public Const shtMdl As String = "SMdl"
'-----------------------------------------------------------------------------------------------
' New Test suite for mdlScenario and Refresh Classes
' JDL 12/15/21  Refactored 10/21/24 to use Procedures class
Sub TestDriver_mdlScenario()
    Dim procs As New Procedures, AllEnabled As Boolean
    
    With procs
        
        'Initialize Procedure objects; Set up tst_Results sheet; Set Procedures attributes
        .Init procs, ThisWorkbook, "tests_mdlScenario", "tests_mdlScenario"
        
        'Turn off events and Screenupdataing; calculation Automatic
        SetApplEnvir False, False, xlCalculationAutomatic
        
        'Enable/disable all or groups of tests by procedure
        AllEnabled = True
        .mdlInit.Enabled = False
        .mdlRow.Enabled = False
        .mdlVariations.Enabled = False
        .mdlDropdowns.Enabled = True
        .mdlRefreshSpeed.Enabled = False
    End With
        
    '*** mdl initialization and utilities***
    If procs.mdlInit.Enabled Or AllEnabled Then
        procs.curProcedure = procs.mdlInit.name
        
        'ExcelSteps initialization unpopulated and populated
        test_PrepExcelSteps1 procs
        test_PrepExcelSteps2 procs
    
        'mdlScenario Init and helper functions
        test_ParseMdlScenDefn1 procs
        test_ParseMdlScenDefn2 procs
        
        'Init -- see permutations of arg options in ParseMdlScenDefn docstring
        test_mdl_Init1 procs
        test_mdl_Init2 procs
        test_mdl_Init3 procs
        test_mdl_Init4 procs
        test_mdl_Init5 procs
        test_mdl_Init6 procs

        'ClearModel
        test_ClearModel procs
    End If
    
    '***mdlRow Class - Refactor Refresh Class 1/7/22***
    If procs.mdlRow.Enabled Or AllEnabled Then
        procs.curProcedure = procs.mdlRow.name
        test_mdlRowDefaultModel procs
        test_mdlRowDefaultModel2 procs
        test_mdlRowLiteModel procs
    End If
    
    '***mdlScenario variations***
    If procs.mdlVariations.Enabled Or AllEnabled Then
        procs.curProcedure = procs.mdlVariations.name
    
        'Default model - Calculator/Single column
        test_PopulateSMdl1 procs
        test_ProvisionSMdl1 procs
        test_RefreshSMdl1 procs
                
        'Multicolumn default model
        test_ProvisionSMdl2 procs
        test_RefreshSMdl2 procs
        test_RefreshSMdl2a procs 'w/o range names
        
        'Default, Multicolumn non-contiguous columns
        test_ProvisionSMdl3 procs
        test_RefreshSMdl3 procs
        
        'Calculator model  - w/o and with suppressed header
        test_RefreshSMdl4 procs    'No name prefix
        test_RefreshSMdl4a procs    'With name prefix
        
        'Lite Model
        test_RefreshSMdl5 procs 'Lite Model
        test_RefreshSMdl6 procs 'Lite Model, Non-homed
        test_RefreshSMdl6_HiddenVars procs 'Lite, Non-homed; hide variable names column
        
        'Lite Model, Non-homed; Defn read from Setting; nrows specified
        test_WriteMdlSetting procs
        test_ProvisionSMdl7 procs
        test_RefreshSMdl7 procs
    End If
    
    '***Dropdown list capability***
    If procs.mdlDropdowns.Enabled Or AllEnabled Then
        procs.curProcedure = procs.mdlDropdowns.name
        test_PopulateNamedList procs 'create a named list
        test_AddDropdownSMdl4a procs 'Calculator Model
        test_AddDropdownSMdl2 procs 'Default Model - Multi-column
        test_AddDropdownSMdl5 procs  'Lite Model
    End If
    
    '*** mdl Refresh speedup (7/15/25)***
    If procs.mdlRefreshSpeed.Enabled Or AllEnabled Then
        procs.curProcedure = procs.mdlInit.name
        'test_WriteRngFormulas procs
        'test_UpdateWriteFormulaRng procs
        'test_LegacyRefresh procs
        'test_RefreshSpeedup procs
                
        '10/29/25
        test_RefreshSMdl5 procs 'Lite Model
    End If

    procs.EvalOverall procs
End Sub
'-----------------------------------------------------------------------------------------------
'-----------------------------------------------------------------------------------------------
' procs.mdlRefreshSpeed functions
'-----------------------------------------------------------------------------------------------
'-----------------------------------------------------------------------------------------
' Write array of formulas for current contiguous row range
' JDL 7/17/25
Sub test_UpdateWriteFormulaRng(procs)
    Dim tst As New Test: tst.Init tst, "test_UpdateWriteFormulaRng"
    Dim mdl As Object: Set mdl = ExcelSteps.New_mdl
    Dim r As Object, rngRow As Range
    Dim aryFormulas As Variant, rngFormulas As Range, f1 As String, f2 As String, f3 As String
    Dim tStart As Double, tEnd As Double, i As Integer
    With tst
    
        'Populate and Provision model with non-contiguous row and column blocks
        PopulateMdlSpeedup2 .wkbkTest.Sheets(shtMdl)
        .Assert tst, mdl.Provision(mdl, .wkbkTest, shtMdl)
        
        'Mockup Model's rows have 6 Areas; last two are formula-containing
        .Assert tst, mdl.rngPopRows.Areas.Count = 6
        
        f1 = "=(@side_a_1^2 + @side_b_1^2)^0.5"
        f2 = "=(@side_a_2^2 + @side_b_2^2)^0.5"
        f3 = "=(@side_a_3^2 + @side_b_3^2)^0.5"
    
        'Instance mdlRow for first formula-containing row
        Set r = ExcelSteps.New_mdlRow()
        Set rngRow = .wkbkTest.Sheets(shtMdl).Cells(15, 4)
        r.Init r, mdl, rngRow
        .Assert tst, r.rngVarRow.Address = "$15:$15"
        .Assert tst, r.sformula = f1
        
        'Update Formula Range (first row initializes aryFormulas and rngFormulas)
        .Assert tst, mdl.UpdateFormulaRng(r, mdl, aryFormulas, rngFormulas)
        .Assert tst, UBound(aryFormulas) = 1
        .Assert tst, aryFormulas(1) = f1
        .Assert tst, rngFormulas.Address = "$15:$15"
        
        'Second row - contiguous with first
        Set r = ExcelSteps.New_mdlRow()
        Set rngRow = .wkbkTest.Sheets(shtMdl).Cells(16, 4)
        r.Init r, mdl, rngRow
        .Assert tst, r.rngVarRow.Address = "$16:$16"
        .Assert tst, r.sformula = f2
        
        'Update Formula Range with contiguous second row
        .Assert tst, mdl.UpdateFormulaRng(r, mdl, aryFormulas, rngFormulas)
        .Assert tst, UBound(aryFormulas) = 2
        .Assert tst, aryFormulas(1) = f1
        .Assert tst, aryFormulas(2) = f2
        .Assert tst, rngFormulas.Address = "$15:$16"
        
        'Third row - not contiguous with first two
        Set r = ExcelSteps.New_mdlRow()
        Set rngRow = .wkbkTest.Sheets(shtMdl).Cells(18, 4)
        r.Init r, mdl, rngRow
        .Assert tst, r.rngVarRow.Address = "$18:$18"
        .Assert tst, r.sformula = f3
        
        'Write first two rows' FormulaRng and reset to non-contiguous 3rd row
        .Assert tst, mdl.UpdateFormulaRng(r, mdl, aryFormulas, rngFormulas)
        .Assert tst, UBound(aryFormulas) = 1
        .Assert tst, aryFormulas(1) = f3
        .Assert tst, rngFormulas.Address = "$18:$18"
        
        'Check that first two rows' formulas were written
        .Assert tst, .wkbkTest.Sheets(shtMdl).Cells(15, 9).Formula = f1
        .Assert tst, .wkbkTest.Sheets(shtMdl).Cells(15, 11).Formula = f1
        .Assert tst, .wkbkTest.Sheets(shtMdl).Cells(15, 12).Formula = f1
        
        .Assert tst, .wkbkTest.Sheets(shtMdl).Cells(16, 9).Formula = f2
        .Assert tst, .wkbkTest.Sheets(shtMdl).Cells(16, 11).Formula = f2
        .Assert tst, .wkbkTest.Sheets(shtMdl).Cells(16, 12).Formula = f2

 
        .Update tst, procs
    End With
End Sub
'-----------------------------------------------------------------------------------------
' Write array of formulas for current contiguous row range
' JDL 7/17/25
Sub test_WriteRngFormulas(procs)
    Dim tst As New Test: tst.Init tst, "test_WriteRngFormulas"
    Dim mdl As Object: Set mdl = ExcelSteps.New_mdl
    Dim aryFormulas As Variant, rngFormulas As Range, f1 As String, f2 As String
    Dim tStart As Double, tEnd As Double, i As Integer
    With tst
    
        'Populate and Provision model with non-contiguous row and column blocks
        PopulateMdlSpeedup2 .wkbkTest.Sheets(shtMdl)
        .Assert tst, mdl.Provision(mdl, .wkbkTest, shtMdl)
        
        'Mockup Model's rows have 6 Areas; last two are formula-containing
        .Assert tst, mdl.rngPopRows.Areas.Count = 6
            
        'Point aryFormulas to range of formula strings for .rngPopRows.Areas(5)
        With .wkbkTest.Sheets(shtMdl)
            aryFormulas = Application.Transpose(Range(.Cells(15, 7), .Cells(16, 7)).Value)
        End With
        
        'Set area_formulas to entire rows for .Areas(5) and call Write method
        Set rngFormulas = mdl.rngPopRows.Areas(5).EntireRow
        .Assert tst, mdl.WriteRngFormulas(mdl, aryFormulas, rngFormulas)
    
        'Check formulas were applied
        f1 = "=(@side_a_1^2 + @side_b_1^2)^0.5"
        f2 = "=(@side_a_2^2 + @side_b_2^2)^0.5"

        .Assert tst, .wkbkTest.Sheets(shtMdl).Cells(15, 9).Formula = f1
        .Assert tst, .wkbkTest.Sheets(shtMdl).Cells(15, 11).Formula = f1
        .Assert tst, .wkbkTest.Sheets(shtMdl).Cells(15, 12).Formula = f1
        
        .Assert tst, .wkbkTest.Sheets(shtMdl).Cells(16, 9).Formula = f2
        .Assert tst, .wkbkTest.Sheets(shtMdl).Cells(16, 11).Formula = f2
        .Assert tst, .wkbkTest.Sheets(shtMdl).Cells(16, 12).Formula = f2
            
        'Benchmark time for 100 writes (0.2 s 7/17/25)
        If True Then
            tStart = Timer
            For i = 1 To 100
                .Assert tst, mdl.WriteRngFormulas(mdl, aryFormulas, rngFormulas)
            Next i
            tEnd = Timer
            Debug.Print "WriteRngFormulas Elapsed time (seconds): " & Round((tEnd - tStart), 1)
        End If
        .Update tst, procs
    End With
End Sub
'----------------------------------------------------------------------------
' Populate/Refresh a model with multiple variable blocks (Areas)
' JDL 7/18/25
'
' Speedups:
' 1. Write formulas as arrays instead of line by line (.Refresh)
' 2. Don't delete previous version of name before creating (wasn't needed)
' 3. Hard code row naming during Refresh by adding new name prefix string attribute
' 4.
'
Sub test_RefreshSpeedup(procs)
    Dim tst As New Test: tst.Init tst, "test_RefreshSpeedup"
    Dim mdl As Object: Set mdl = ExcelSteps.New_mdl
    Dim aryVals As Variant, aryExpected As Variant, rng As Range, s As String
    Dim f1 As String, f2 As String, f3 As String, f4 As String
    Dim tStart As Double, tEnd As Double, i As Integer
    
    With tst
        
        'Populate, Provision and Refresh
        PopulateMdlSpeedup2 .wkbkTest.Sheets(shtMdl)
        .Assert tst, mdl.Provision(mdl, .wkbkTest, shtMdl)
        .Assert tst, mdl.Refresh(mdl)
    
        'Populated rows multirange has 4 Areas
        .Assert tst, mdl.rngPopRows.Areas.Count = 6
        
        'Check formulas were applied
        f1 = "=(@side_a_1^2 + @side_b_1^2)^0.5"
        f2 = "=(@side_a_2^2 + @side_b_2^2)^0.5"
        f3 = "=(@side_a_3^2 + @side_b_3^2)^0.5"
        f4 = "=(@side_a_4^2 + @side_b_4^2)^0.5"
        
        .Assert tst, .wkbkTest.Sheets(shtMdl).Cells(15, 9).Formula = f1
        .Assert tst, .wkbkTest.Sheets(shtMdl).Cells(15, 11).Formula = f1
        .Assert tst, .wkbkTest.Sheets(shtMdl).Cells(15, 12).Formula = f1
        
        .Assert tst, .wkbkTest.Sheets(shtMdl).Cells(16, 9).Formula = f2
        .Assert tst, .wkbkTest.Sheets(shtMdl).Cells(16, 11).Formula = f2
        .Assert tst, .wkbkTest.Sheets(shtMdl).Cells(16, 12).Formula = f2
        
        .Assert tst, .wkbkTest.Sheets(shtMdl).Cells(18, 9).Formula = f3
        .Assert tst, .wkbkTest.Sheets(shtMdl).Cells(18, 11).Formula = f3
        .Assert tst, .wkbkTest.Sheets(shtMdl).Cells(18, 12).Formula = f3
        
        .Assert tst, .wkbkTest.Sheets(shtMdl).Cells(19, 9).Formula = f4
        .Assert tst, .wkbkTest.Sheets(shtMdl).Cells(19, 11).Formula = f4
        .Assert tst, .wkbkTest.Sheets(shtMdl).Cells(19, 12).Formula = f4
        
        '1 Benchmark time for 100 refreshes (7.6 s with legacy refresh RefreshPrev)
        '2 Benchmark with speedup (5.7 s with Refresh)
        If True Then
            tStart = Timer
            For i = 1 To 100
                .Assert tst, mdl.Refresh(mdl)
            Next i
            tEnd = Timer
            Debug.Print "Speedup Elapsed time (seconds): " & Round((tEnd - tStart), 1)
        End If
        .Update tst, procs
    End With
End Sub
'----------------------------------------------------------------------------
' Populate/Refresh a model with multiple variable blocks (Areas)
' JDL 7/15/25
'
Sub test_LegacyRefresh(procs)
    Dim tst As New Test: tst.Init tst, "test_LegacyRefresh"
    Dim mdl As Object: Set mdl = ExcelSteps.New_mdl
    Dim aryVals As Variant, aryExpected As Variant, rng As Range, s As String
    Dim tStart As Double, tEnd As Double, i As Integer
    
    With tst
        
        'Populate, Provision and RefreshPrev
        PopulateMdlSpeedup2 .wkbkTest.Sheets(shtMdl)
        .Assert tst, mdl.Provision(mdl, .wkbkTest, shtMdl)
        .Assert tst, mdl.Refresh(mdl)
    
        'Populated rows multirange has 6 Areas
        .Assert tst, mdl.rngPopRows.Areas.Count = 6
        
        'Benchmark time for 100 RefreshPreves
        If True Then
            tStart = Timer
            For i = 1 To 100
                .Assert tst, mdl.Refresh(mdl)
            Next i
            tEnd = Timer
            Debug.Print "Elapsed time (seconds): " & Round((tEnd - tStart), 1)
        End If
        .Update tst, procs
    End With
End Sub
'-----------------------------------------------------------------------------------------------
'-----------------------------------------------------------------------------------------------
' procs.mdlInit functions
'-----------------------------------------------------------------------------------------------
'-----------------------------------------------------------------------------------------------
' PrepExcelStepsSht with Add new blank ExcelSteps
' Updated 7/21/23 JDL; Refactored 10/18/24; updated 11/17/25 for
' elimination of setting tblSteps col ranges
'
Sub test_PrepExcelSteps1(procs)
    Dim tst As New Test: tst.Init tst, "test_PrepExcelSteps1"
    Dim tblSteps As Object: Set tblSteps = ExcelSteps.New_tbl
    Dim refr As Object: Set refr = ExcelSteps.New_Refresh
    
    'Arrays for checking results
    Dim aryVals As Variant, aryExpected As Variant
    
    With tst
    
        'Clear prior ExcelSteps sheet if any
        If SheetExists(.wkbkTest, shtSteps) Then .wkbkTest.Sheets(shtSteps).Cells.Clear
        
        'Initialize Refresh class
        .Assert tst, refr.InitMdl(refr, .wkbkTest)
        
        'Prep recreates/reformats ExcelSteps and sets tblSteps.rowCur
        .Assert tst, refr.PrepExcelStepsSht(refr, tblSteps)
        .Assert tst, SheetExists(.wkbkTest, shtSteps)
        
        aryVals = Array(ListFromArray(tblSteps.rngHeader), _
                        tblSteps.colrngStrInput.NumberFormat, _
                        tblSteps.colrngNumFmt.NumberFormat, _
                        tblSteps.cellHome.CurrentRegion.Address, _
                        tblSteps.rngrows.Address)
                        
        aryExpected = Array(sHeaderSteps, _
                            "@", _
                            "@", _
                            "$A$1:$I$2", _
                            "$2:$21")
        .TestAryVals tst, aryVals, aryExpected
            
        tst.Update tst, procs
    End With
End Sub
'-----------------------------------------------------------------------------------------------
' PrepExcelStepsSht with existing, populated ExcelSteps
' Updated 7/21/23 JDL; Refactored 10/18/24
'
Sub test_PrepExcelSteps2(procs)
    Dim tst As New Test: tst.Init tst, "test_PrepExcelSteps2"
    Dim tblSteps As Object: Set tblSteps = ExcelSteps.New_tbl
    Dim refr As Object: Set refr = ExcelSteps.New_Refresh
    Dim aryVals As Variant, aryExpected As Variant
    
    With tst
    
        'Clear prior ExcelSteps sheet if any
        If SheetExists(.wkbkTest, shtSteps) Then .wkbkTest.Sheets(shtSteps).Cells.Clear
        
        'Initialize Refresh class
        .Assert tst, refr.InitMdl(refr, .wkbkTest)
        
        'Prep blank sheet and populate with instructions
        .Assert tst, refr.PrepExcelStepsSht(refr, tblSteps)
        PopulateStepsSMdl .wkbkTest, "SMdlDash"
        
        'Prep recreates/reformats ExcelSteps and sets tblSteps.rowCur
        .Assert tst, refr.PrepExcelStepsSht(refr, tblSteps)
        
        aryVals = Array(ListFromArray(tblSteps.rngHeader), _
                        tblSteps.colrngStrInput.NumberFormat, _
                        tblSteps.colrngNumFmt.NumberFormat, _
                        tblSteps.cellHome.CurrentRegion.Address, _
                        tblSteps.rngrows.Address)
                        
        aryExpected = Array(sHeaderSteps, _
                            "@", _
                            "@", _
                            "$A$1:$I$3", _
                            "$2:$3")
        .TestAryVals tst, aryVals, aryExpected
            
        tst.Update tst, procs
    End With
End Sub
'-----------------------------------------------------------------------------
' mdlScenario Init and helper functions
'-----------------------------------------------------------------------------
' Parse Scenario model defn string from argument-specified Defn
' JDL 7/17/23; Refactored 10/18/24
'
Sub test_ParseMdlScenDefn1(procs)
    Dim tst As New Test: tst.Init tst, "test_ParseMdlScenDefn1"
    Dim mdl As Object: Set mdl = ExcelSteps.New_mdl
    
    'Default - sht as mdlName; Defn not Missing
    With tst
        Set mdl.wkbk = .wkbkTest
        .Assert tst, mdl.ParseMdlScenDefn(mdl, defn_val) = True
        .Assert tst, mdl.sht = shtMdl
        .Assert tst, mdl.cellHome.Address = "$I$4"
        .Assert tst, mdl.nRows = 0
        .Assert tst, mdl.IsCalc = True
        .Assert tst, mdl.IsSuppHeader = False
        .Assert tst, mdl.IsRngNames = True
        .Assert tst, mdl.IsMdlNmPrefix = True
        .Assert tst, mdl.IsLiteModel = True
        .Assert tst, mdl.MdlName = shtMdl
    
    'Need to add test of Read Setting if Defn missing
    
        .Update tst, procs
    End With
End Sub

'-----------------------------------------------------------------------------
' Parse Scenario model defn string from argument - Defn includes mdlName
' JDL 2/15/25
'
Sub test_ParseMdlScenDefn2(procs)
    Dim tst As New Test: tst.Init tst, "test_ParseMdlScenDefn2"
    Dim mdl As Object: Set mdl = ExcelSteps.New_mdl
    
    'MdlName specified in Defn overrides using .sht
    With tst
        Set mdl.wkbk = .wkbkTest
        .Assert tst, mdl.ParseMdlScenDefn(mdl, defn_val_mdlName) = True
        .Assert tst, mdl.sht = "SMdl"
        .Assert tst, mdl.MdlName = "SMdlType2"
    
        .Update tst, procs
    End With
End Sub
'-----------------------------------------------------------------------------
' Init Argument-Specified Scenario Model (locations and parameters)
' JDL 7/17/23; Refactored in separate tests 2/15/25
'
Sub test_mdl_Init1(procs)
    Dim tst As New Test: tst.Init tst, "test_mdl_Init1"
    Dim mdl As Object: Set mdl = ExcelSteps.New_mdl

    With tst
    
        'Default model (only sht arg specified)
        Set mdl = ExcelSteps.New_mdl
        .Assert tst, mdl.Init(mdl, .wkbkTest, sht:=shtMdl)
        .Assert tst, mdl.wkbk.name = ThisWorkbook.name
        .Assert tst, mdl.sht = shtMdl
        .Assert tst, mdl.wksht.name = shtMdl
        .Assert tst, mdl.cellHome.Address = "$A$2"
        .Assert tst, mdl.IsCalc = False
        .Assert tst, mdl.IsSuppHeader = False
        .Assert tst, mdl.IsRngNames = True
        .Assert tst, mdl.IsMdlNmPrefix = False
        .Assert tst, mdl.IsLiteModel = False
        .Update tst, procs
    End With
End Sub
'-----------------------------------------------------------------------------
' Init Argument-Specified Scenario Model (locations and parameters)
' JDL 7/17/23; Refactored in separate tests 2/15/25
'
Sub test_mdl_Init2(procs)
    Dim tst As New Test: tst.Init tst, "test_mdl_Init2"
    Dim mdl As Object: Set mdl = ExcelSteps.New_mdl

    With tst
    
        'Default Model with Suppress header
        Set mdl = ExcelSteps.New_mdl
        .Assert tst, mdl.Init(mdl, .wkbkTest, sht:=shtMdl, IsSuppHeader:=True)
        .Assert tst, mdl.IsSuppHeader = True
        .Update tst, procs
    End With
End Sub
'-----------------------------------------------------------------------------
' Init Argument-Specified Scenario Model (locations and parameters)
' JDL 7/17/23; Refactored in separate tests 2/15/25
'
Sub test_mdl_Init3(procs)
    Dim tst As New Test: tst.Init tst, "test_mdl_Init3"
    Dim mdl As Object: Set mdl = ExcelSteps.New_mdl

    With tst
    
        'Boolean args all True
        .Assert tst, mdl.Init(mdl, .wkbkTest, sht:=shtMdl, IsCalc:=True, _
                IsSuppHeader:=True, IsLiteModel:=True, IsRngNames:=True, _
                IsMdlNmPrefix:=True)
        .Assert tst, mdl.IsCalc = True
        .Assert tst, mdl.IsSuppHeader = True
        .Assert tst, mdl.IsLiteModel = True
        .Assert tst, mdl.IsRngNames = True
        .Assert tst, mdl.IsMdlNmPrefix = True
        .Update tst, procs
    End With
End Sub
'-----------------------------------------------------------------------------
' Init Argument-Specified Scenario Model (locations and parameters)
' JDL 7/17/23; Refactored in separate tests 2/15/25
'
Sub test_mdl_Init4(procs)
    Dim tst As New Test: tst.Init tst, "test_mdl_Init4"
    Dim mdl As Object: Set mdl = ExcelSteps.New_mdl

    With tst
    
        'Non-default mdl; specified Defn (no sht or mdlName args to override)
        .Assert tst, mdl.Init(mdl, .wkbkTest, defn:=defn_val)
        .Assert tst, mdl.sht = shtMdl
        .Assert tst, mdl.MdlName = shtMdl
        .Assert tst, mdl.cellHome.Address = "$I$4"
        .Assert tst, mdl.nRows = 0
        .Assert tst, mdl.IsCalc = True
        .Assert tst, mdl.IsSuppHeader = False
        .Assert tst, mdl.IsRngNames = True
        .Assert tst, mdl.IsMdlNmPrefix = True
        .Assert tst, mdl.IsLiteModel = True
        
        .Update tst, procs
    End With
End Sub
'-----------------------------------------------------------------------------
' Init Argument-Specified Scenario Model (locations and parameters)
' JDL 7/17/23; Refactored in separate tests 2/15/25
'
Sub test_mdl_Init5(procs)
    Dim tst As New Test: tst.Init tst, "test_mdl_Init5"
    Dim mdl As Object:
    With tst
    
        'sht arg overrides sht name in parsed Defn
        Set mdl = ExcelSteps.New_mdl
        .Assert tst, mdl.Init(mdl, .wkbkTest, sht:="SMdl2", defn:=defn_val)
        .Assert tst, mdl.sht = "SMdl2"
        .Assert tst, mdl.MdlName = "SMdl2"
        
        'mdlName arg overrides sht name as mdlName in parsed Defn (no mdlName in defn)
        Set mdl = ExcelSteps.New_mdl
        .Assert tst, mdl.Init(mdl, .wkbkTest, defn:=defn_val, MdlName:="SMdl2")
        .Assert tst, mdl.sht = shtMdl
        .Assert tst, mdl.MdlName = "SMdl2"

        'mdlName arg overrides mdlName name in parsed Defn (no mdlName in defn)
        Set mdl = ExcelSteps.New_mdl
        .Assert tst, mdl.Init(mdl, .wkbkTest, defn:=defn_val_mdlName, MdlName:="SMdl2")
        .Assert tst, mdl.sht = shtMdl
        .Assert tst, mdl.MdlName = "SMdl2"

        'both sht and mdlName args specified; override defn values
        Set mdl = ExcelSteps.New_mdl
        .Assert tst, mdl.Init(mdl, .wkbkTest, sht:="SMdl3", defn:=defn_val_mdlName, MdlName:="SMdl2")
        .Assert tst, mdl.sht = "SMdl3"
        .Assert tst, mdl.MdlName = "SMdl2"
        
        'Check for non-default attribute value from defn (e.g defn got parsed/used)
        .Assert tst, mdl.IsCalc = True
        .Update tst, procs
    End With
End Sub
'-----------------------------------------------------------------------------
' Init Argument-Specified Scenario Model (locations and parameters)
' JDL 7/17/23; Refactored in separate tests 2/15/25
' Tests of sht names that are not valid Excel rng names
'
Sub test_mdl_Init6(procs)
    Dim tst As New Test: tst.Init tst, "test_mdl_Init6"
    Dim mdl As Object:
    
    With tst
    
        'default model; sht arg is not valid rng name
        Set mdl = ExcelSteps.New_mdl
        .Assert tst, mdl.Init(mdl, .wkbkTest, sht:="SMdl 2")
        .Assert tst, mdl.sht = "SMdl 2"
        .Assert tst, mdl.MdlName = "SMdl2"
        
        'sht arg override is not valid rng name; overrides name in parsed Defn
        'mdlName based on xlName(.sht)
        Set mdl = ExcelSteps.New_mdl
        .Assert tst, mdl.Init(mdl, .wkbkTest, sht:="SMdl 2", defn:=defn_val)
        .Assert tst, mdl.sht = "SMdl 2"
        .Assert tst, mdl.MdlName = "SMdl2"
        .Update tst, procs
    End With
End Sub
'-----------------------------------------------------------------------------------------------
'ClearModel (and ApplyBorderAroundModel)
'JDL 1/5/21; Refactored 10/18/24
'
Sub test_ClearModel(procs)
    Dim tst As New Test: tst.Init tst, "test_ClearModel"
    Dim mdl As Object: Set mdl = ExcelSteps.New_mdl
    Dim IsBlank As Boolean, w As Variant, xlEdge As Variant
    
    With tst
        PopulateSMdl1 .wkbkTest.Sheets(shtMdl)
        .Assert tst, mdl.Provision(mdl, .wkbkTest, shtMdl)
        .Assert tst, mdl.Refresh(mdl)
        mdl.ApplyBorderAroundModel mdl, True
        
        .Assert tst, (Len(mdl.wksht.Cells(2, 3)) > 0)
    
        mdl.ClearModel mdl, True
        
        'Test that the model cell range is blank after Clear
        IsBlank = True
        For Each w In mdl.rngMdl
            If Not IsEmpty(w) Then IsBlank = False
        Next w
        .Assert tst, IsBlank
        
        IsBlank = True
        For Each w In mdl.rngMdl
            For Each xlEdge In Array(xlEdgeLeft, xlEdgeTop, xlEdgeBottom, xlEdgeRight)
                If Not w.Borders(xlEdge).LineStyle = xlNone Then IsBlank = False
            Next xlEdge
        Next w
        .Assert tst, IsBlank

        .Update tst, procs
    End With
End Sub
'-----------------------------------------------------------------------------------------------
'-----------------------------------------------------------------------------------------------
' procs.mdlRow procedure tests
'-----------------------------------------------------------------------------------------------
'-----------------------------------------------------------------------------------------------
'Use mdlRow to set row properties - Default multi-column model (SMdl2)
'JDL 1/7/22; Updated 7/21/23; Refactored 10/18/24
'
Sub test_mdlRowDefaultModel(procs)
    Dim tst As New Test: tst.Init tst, "test_mdlRowDefaultModel"
    Dim mdl As Object: Set mdl = ExcelSteps.New_mdl
    Dim r As Object, tblSteps As Object
    
    'Arrays for checking results
    Dim aryVals As Variant, aryExpected As Variant
    
    With tst
        PopulateSMdl2 .wkbkTest.Sheets(shtMdl)
        .Assert tst, mdl.Provision(mdl, .wkbkTest, shtMdl)
        .Assert tst, mdl.rngFormulaRows.Address = "$D$6"
    
        'Instance mdlRow for "side_a" row
        Set r = ExcelSteps.New_mdlRow
        r.Init r, mdl, mdl.wksht.Cells(3, 4)
    
        aryVals = Array(r.rngVar.Address, _
                        r.rngVarRow.Address, _
                        r.sVar, _
                        "x_" & r.NumFmt, _
                        r.HasFormula, _
                        r.sformula)
                        
        aryExpected = Array("$D$3", _
                            "$3:$3", _
                            "side_a", _
                            "x_0", _
                            False, _
                            "Input")

        .TestAryVals tst, aryVals, aryExpected
        
        'Instance mdlRow for "side_c" (e.g. formula-containing) row
        Set r = ExcelSteps.New_mdlRow
        r.Init r, mdl, mdl.wksht.Cells(6, 4)
                                    
        aryVals = Array(r.rngVar.Address, _
                        r.rngVarRow.Address, _
                        r.sVar, _
                        "x_" & r.NumFmt, _
                        r.HasFormula, _
                        r.sformula)
                                    
        aryExpected = Array("$D$6", _
                            "$6:$6", _
                            "side_c", _
                            "x_0.00", _
                            True, _
                            "=(@side_a^2 + @side_b^2)^0.5")
        .TestAryVals tst, aryVals, aryExpected
        .Update tst, procs
    End With
End Sub
'-----------------------------------------------------------------------------------------------
'Use mdlRow to set row properties - Default multi-column model (SMdl2)
'JDL 12/5/24 Troubleshoot provision - Model with empty $D$2 variable name
'
Sub test_mdlRowDefaultModel2(procs)
    Dim tst As New Test: tst.Init tst, "test_mdlRowDefaultModel"
    Dim mdl As Object: Set mdl = ExcelSteps.New_mdl
    Dim r As Object
    
    'Arrays for checking results
    Dim aryVals As Variant, aryExpected As Variant
    
    With tst
        PopulateSMdl2a .wkbkTest.Sheets(shtMdl)
        .Assert tst, mdl.Provision(mdl, .wkbkTest, shtMdl)
        .Assert tst, mdl.rngPopRows.Address = "$D$2,$D$6"
        .Assert tst, mdl.rngrows.Address = "$2:$6"
        .Update tst, procs
    End With
End Sub
'-----------------------------------------------------------------------------------------------
'Use mdlRow to set row properties - Lite multi-column model (SMdl6)
'JDL 1/7/22; Refactored 10/18/24
'
Sub test_mdlRowLiteModel(procs)
    Dim tst As New Test: tst.Init tst, "test_mdlRowLiteModel"
    Dim mdl As Object: Set mdl = ExcelSteps.New_mdl
    Dim refr As Object: Set refr = ExcelSteps.New_Refresh
    Dim r As Object, tblSteps As Object
    
    'Array for checking results and temp variables
    Dim aryVals As Variant, aryExpected As Variant
    Dim cellHome As Range
    
    With tst
        
        'Recreate and populate ExcelSteps
        PrepBlankStepsForTesting .wkbkTest, refr, tblSteps
        PopulateStepsSMdl .wkbkTest, shtMdl

        'Populate and Provision Scenario Model
        PopulateSMdl6 tst, .wkbkTest.Sheets(shtMdl)
        Set cellHome = .wkbkTest.Sheets(shtMdl).Cells(10, 6)
        .Assert tst, mdl.Provision(mdl, .wkbkTest, shtMdl, IsLiteModel:=True, _
                                IsSuppHeader:=True, cellHome:=cellHome)
        .Assert tst, mdl.rngFormulaRows.Address = "$H$15"
    
        'Instance mdlRow for "side_a" row
        Set r = ExcelSteps.New_mdlRow
        r.Init r, mdl, mdl.wksht.Cells(12, 8)

        .TestAryVals tst, Array("x_" & r.NumFmt, r.HasFormula), Array("x_0.000", False)
    
        'Instance mdlRow for "side_c" row
        Set r = ExcelSteps.New_mdlRow
        r.Init r, mdl, mdl.wksht.Cells(15, 8)
    
        aryExpected = Array("x_0.00", "=(side_a^2 + side_b^2)^0.5")
        .TestAryVals tst, Array("x_" & r.NumFmt, r.sformula), aryExpected
        .Update tst, procs
    End With
End Sub
'-----------------------------------------------------------------------------------------------
'-----------------------------------------------------------------------------------------------
' procs.mdlVariations
'-----------------------------------------------------------------------------------------------
'----------------------------------------------------------------------------
' Populate a default model (single column but not IsLite)
' Updated JDL 7/17/23; Refactored 10/18/24
'
Sub test_PopulateSMdl1(procs)
    Dim tst As New Test: tst.Init tst, "test_PopulateSMdl1"
    
    With tst
        PopulateSMdl1 .wkbkTest.Sheets(shtMdl)  'Sub in modPopulateMdl
        .Assert tst, .wkbkTest.Sheets(shtMdl).Cells(1, 1) = "Grp"
        .Assert tst, .wkbkTest.Sheets(shtMdl).Cells(4, 9) = 4
        .Assert tst, .wkbkTest.Sheets(shtMdl).Cells(4, 6).NumberFormat = "@"
        .Update tst, procs
    End With
End Sub
'----------------------------------------------------------------------------
' Provision a default model (single column but not IsLite)
' Updated JDL 7/17/23; Refactored 10/18/24
'
Sub test_ProvisionSMdl1(procs)
    Dim tst As New Test: tst.Init tst, "test_ProvisionSMdl1"
    Dim mdl As Object: Set mdl = ExcelSteps.New_mdl
    Dim aryVals As Variant, aryExpected As Variant, rng As Range
    
    With tst
        
        'Provision a Scenario Model on sheet SM1
        PopulateSMdl1 .wkbkTest.Sheets(shtMdl)
        .Assert tst, mdl.Provision(mdl, .wkbkTest, shtMdl)
        
        aryVals = Array(mdl.sht, _
                        mdl.cellHome.Address, _
                        mdl.IsCalc, _
                        mdl.IsSuppHeader, _
                        mdl.IsRngNames, _
                        mdl.IsMdlNmPrefix, _
                        mdl.IsLiteModel, _
                        mdl.rngrows.Address)
                        
        'Booleans: IsCalc, IsSuppHeader, IsRngNames, IsMdlNmPrefix, IsLiteModel
        aryExpected = Array(shtMdl, _
                            "$A$2", _
                            False, _
                            False, _
                            True, _
                            False, _
                            False, _
                            "$2:$6")
        .TestAryVals tst, aryVals, aryExpected
        
        'Check header address, populated cols, populated rows and row range
        aryVals = Array(mdl.rngHeader.Address, _
                        mdl.rngPopCols.Address, _
                        mdl.rngPopRows.Address, _
                        mdl.rngrows.Address)
                        
        aryExpected = Array("$A$1:$G$1", _
                            "$I$2", _
                            "$D$2:$D$4,$D$6", _
                            "$2:$6")
        .TestAryVals tst, aryVals, aryExpected
        
        .Update tst, procs
    End With
End Sub
'----------------------------------------------------------------------------
' Refresh a default model (single column but not IsCalc)
' Updated JDL 7/17/23; Refactored 10/18/24
'
Sub test_RefreshSMdl1(procs)
    Dim tst As New Test: tst.Init tst, "test_RefreshSMdl1"
    Dim mdl As Object: Set mdl = ExcelSteps.New_mdl
    Dim aryVals As Variant, aryExpected As Variant, rng As Range, s As String
    
    With tst
        
        'Populate, Provision and Refresh
        PopulateSMdl1 .wkbkTest.Sheets(shtMdl)
        .Assert tst, mdl.Provision(mdl, .wkbkTest, shtMdl)
        .Assert tst, mdl.Refresh(mdl)
    
        'Check value, formula and cell style
        aryVals = Array(mdl.wksht.Cells(6, 9), _
                        mdl.wksht.Cells(6, 9).Formula)
                                
        .TestAryVals tst, aryVals, Array(5, "=(@side_a^2 + @side_b^2)^0.5")
        '.CkStyleMatch tst, Range(mdl.wksht.Cells(2, 4), mdl.wksht.Cells(2, 7)), "Note"
        
        'Check range name creation
        .CkRngNameAddresses tst, Array("side_a", "side_b", "side_c"), _
                                  Array("$3:$3", "$4:$4", "$6:$6")
        .CkRngNameAddresses tst, Array("Triangle1"), Array("$I:$I")
        
        .Update tst, procs
    End With
End Sub
'-----------------------------------------------------------------------------------------------
' Multi-column default model (Contiguous Columns)
' Updated JDL 7/21/23; refactored 10/18/24
'
Sub test_ProvisionSMdl2(procs)
    Dim tst As New Test: tst.Init tst, "test_ProvisionSMdl2"
    Dim mdl As Object: Set mdl = ExcelSteps.New_mdl
    Dim aryVals As Variant, aryExpected As Variant
    
    With tst
        PopulateSMdl2 .wkbkTest.Sheets(shtMdl)
        .Assert tst, mdl.Provision(mdl, .wkbkTest, shtMdl)
        
        aryVals = Array(mdl.rngPopRows.Count, _
                        mdl.IsCalc, _
                        mdl.rngPopCols.Address, _
                        mdl.rngrows.Address, _
                        mdl.IsRngNames)
        aryExpected = Array(4, _
                            False, _
                            "$I$2:$J$2", _
                            "$2:$6", _
                            True)
        .TestAryVals tst, aryVals, aryExpected
        
        .Update tst, procs
    End With
End Sub
'-----------------------------------------------------------------------------------------------
'Multi-column default model
' Refactored 10/18/24
'
Sub test_RefreshSMdl2(procs)
    Dim tst As New Test: tst.Init tst, "test_RefreshSMdl2"
    Dim mdl As Object: Set mdl = ExcelSteps.New_mdl
    Dim aryVals As Variant, aryExpected As Variant
    
    With tst
        PopulateSMdl2 .wkbkTest.Sheets(shtMdl)
        .Assert tst, mdl.Provision(mdl, .wkbkTest, shtMdl)
        .Assert tst, mdl.Refresh(mdl)
            
        'Test Header column and calculated cell styles; check ExcelSteps specified format
        CheckHeaderColFormat tst, mdl
        .CkStyleMatch tst, Range(mdl.wksht.Cells(6, 9), mdl.wksht.Cells(6, 10)), "Calculation"
        '.CkStyleMatch tst, Range(mdl.wksht.Cells(2, 9), mdl.wksht.Cells(2, 10)), "Note"
        
        'Test Formula values correct
        aryVals = Array(mdl.wksht.Cells(6, 9), _
                        mdl.wksht.Cells(6, 10))
        aryExpected = Array(5, 10)
        .TestAryVals tst, aryVals, aryExpected

        'Test variable name creation
        .CkRngNameAddresses tst, Array("side_a", "side_b", "side_c"), _
                                  Array("$3:$3", "$4:$4", "$6:$6")
                                  
        .CkRngNameAddresses tst, Array("Triangle1", "Triangle2"), _
                                  Array("$I:$I", "$J:$J")
        .Update tst, procs
    End With
End Sub
'-----------------------------------------------------------------------------------------------
'Multi-column default model - Same as RefreshSMdl2 but no range names
'Refactored 10/18/24
'
Sub test_RefreshSMdl2a(procs)
    Dim tst As New Test: tst.Init tst, "test_RefreshSMdl2a"
    Dim mdl As Object: Set mdl = ExcelSteps.New_mdl
    Dim aryVals As Variant, aryExpected As Variant
    
    With tst
        PopulateSMdl2 .wkbkTest.Sheets(shtMdl)
        .Assert tst, mdl.Provision(mdl, .wkbkTest, shtMdl, IsRngNames:=False)
        .Assert tst, mdl.Refresh(mdl)
        
        .Assert tst, mdl.IsRngNames = False
                            
        'Test no row or column names
        .Assert tst, Not NameExists(.wkbkTest, "side_a")
        .Assert tst, Not NameExists(.wkbkTest, "side_b")
        .Assert tst, Not NameExists(.wkbkTest, "side_c")
        .Assert tst, Not NameExists(.wkbkTest, "Triangle1")
        .Assert tst, Not NameExists(.wkbkTest, "Triangle2")
                    
        .Update tst, procs
    End With
End Sub
'-----------------------------------------------------------------------------------------------
' Multi-column default model (Non-Contiguous Columns)
' Updated JDL 7/18/23; Refactored 10/18/24
'
Sub test_ProvisionSMdl3(procs)
    Dim tst As New Test: tst.Init tst, "test_ProvisionSMdl3"
    Dim mdl As Object: Set mdl = ExcelSteps.New_mdl
    Dim aryVals As Variant, aryExpected As Variant
    
    With tst

        'Populate and Add a column outline to test hidden model column
        PopulateAndProvisionSMdl3 tst, mdl
    
        'Check mdl attributes match expected
        With mdl
            aryVals = Array(.sht, _
                            .wksht.name, _
                            .rngPopRows.Count, _
                            .cellHome.Address, _
                            .rngPopRows.Address, _
                            .rngPopCols.Address, _
                            .rngrows.Address, _
                            .nRows)
            aryExpected = Array(shtMdl, _
                                shtMdl, _
                                4, _
                                "$A$2", _
                                "$D$2:$D$4,$D$6", _
                                Union(.wksht.Cells(2, 9), .wksht.Cells(2, 11)).Address, _
                                "$2:$6", _
                                5)
        End With
        .TestAryVals tst, aryVals, aryExpected
        
        .Update tst, procs
    End With
End Sub
'-----------------------------------------------------------------------------------------------
' Populate and Add a column outline to test hidden model column
' JDL 7/18/23; Refactored 10/18/24
'
Sub PopulateAndProvisionSMdl3(tst, mdl)
    PopulateSMdl3 tst.wkbkTest.Sheets(shtMdl)
    With tst.wkbkTest.Sheets(shtMdl)
        .Columns(11).Columns.Group
        .Outline.ShowLevels ColumnLevels:=1
    End With
    
    tst.Assert tst, mdl.Provision(mdl, tst.wkbkTest, shtMdl)
End Sub
'-----------------------------------------------------------------------------------------------
' Multi-column default model (Non-Contiguous columns with column outline)
' Updated 7/18/23; Refactored 10/18/24
'
Sub test_RefreshSMdl3(procs)
    Dim tst As New Test: tst.Init tst, "test_RefreshSMdl3"
    Dim mdl As Object: Set mdl = ExcelSteps.New_mdl
    Dim aryVals As Variant, aryExpected As Variant
    
    With tst
        'Populate, Provision and Refresh the model
        PopulateAndProvisionSMdl3 tst, mdl
        .Assert tst, mdl.Refresh(mdl)
    
        'Check header column formatting Scenario Row (2) formatting
        CheckHeaderColFormat tst, mdl
        '.CkStyleMatch tst, Union(mdl.wksht.Cells(2, 9), mdl.wksht.Cells(2, 11)), "Note"

        'Test Formula values correct
        aryVals = Array(mdl.wksht.Cells(6, 9), _
                        mdl.wksht.Cells(6, 11))
        .TestAryVals tst, aryVals, Array(5, 10)

        'Check variable name creation
        aryExpected = Array("$3:$3", "$4:$4", "$6:$6")
        .CkRngNameAddresses tst, Array("side_a", "side_b", "side_c"), aryExpected
        
        .CkRngNameAddresses tst, Array("Triangle1", "Triangle2"), _
                                  Array("$I:$I", "$K:$K")

        tst.Update tst, procs
    End With
End Sub
'-----------------------------------------------------------------------------------------------
' Calculator, Header suppressed
' Updated JDL 7/17/23; Refactored 10/18/24
'
Sub test_RefreshSMdl4(procs)
    Dim tst As New Test: tst.Init tst, "test_RefreshSMdl4"
    Dim mdl As Object: Set mdl = ExcelSteps.New_mdl
    
    With tst
        
        'Populate, Provision and Refresh model
        PopulateSMdl4 .wkbkTest.Sheets(shtMdl)
        .Assert tst, mdl.Provision(mdl, .wkbkTest, shtMdl, IsCalc:=True, IsSuppHeader:=True, _
                IsMdlNmPrefix:=False)
        .Assert tst, mdl.Refresh(mdl)
        
        'Check style/formatting
        CheckHeaderColFormat tst, mdl
        
        'Check Formula calculation value and cell style
        .CkStyleMatch tst, mdl.wksht.Cells(5, 9), "Calculation"
        .Assert tst, mdl.wksht.Cells(5, 9).Value = 5
        
        'Check variable and Scenario column name creation
        .CkRngNameAddresses tst, Array("side_a", "side_b", "side_c"), _
                                  Array("$I$2", "$I$3", "$I$5")
                                  
        .CkRngNameAddresses tst, Array(shtMdl), Array("$I:$I")
        .Update tst, procs
    End With
End Sub
'-----------------------------------------------------------------------------------------------
' Default except Calculator (IsCalc/single-column) and w/ range name prefixes (IsMdlNmPrefix)
' Updated JDL 7/18/23; Refactored 10/18/24
'
Sub test_RefreshSMdl4a(procs)
    Dim tst As New Test: tst.Init tst, "test_RefreshSMdl4a"
    Dim mdl As Object: Set mdl = ExcelSteps.New_mdl
    Dim aryVals As Variant, aryExpected As Variant
    
    With tst
    
        'Populate, Provision and Refresh
        PopulateSMdl4a .wkbkTest.Sheets(shtMdl)
        .Assert tst, mdl.Provision(mdl, .wkbkTest, shtMdl, IsCalc:=True, IsMdlNmPrefix:=True)
        .TestAryVals tst, Array(mdl.IsCalc, mdl.IsMdlNmPrefix), Array(True, True)
        tst.Assert tst, mdl.Refresh(mdl)
                        
        'Check formula transfer
        aryVals = Array(mdl.rngFormulaRows.Address, _
                        LCase(mdl.wksht.Cells(6, 9).Formula), _
                        mdl.wksht.Cells(6, 9))
                        
        aryExpected = Array("$D$6", _
                            "=(smdl_side_a^2 + smdl_side_b^2)^0.5", _
                            5)
                            
        .TestAryVals tst, aryVals, aryExpected
        
        'Check variable name creation
        aryVals = Array("SMdl_side_a", "SMdl_side_b", "SMdl_side_c")
        .CkRngNameAddresses tst, aryVals, Array("$I$3", "$I$4", "$I$6")
        
        .CkRngNameAddresses tst, Array(shtMdl), Array("$I:$I")
    
        'Test Scenario name creation
        .Assert tst, (.wkbkTest.Names(shtMdl).RefersToRange.Address = "$I:$I")

        .Update tst, procs
    End With
End Sub
'-----------------------------------------------------------------------------------------------
' Calculator, Header-suppressed, Lite model
' Updated 7/19/23; Refactored 10/21/24; Add iteration and timing 10/29/25
'
Sub test_RefreshSMdl5(procs)
    Dim tst As New Test: tst.Init tst, "test_RefreshSMdl5"
    Dim mdl As Object
    Dim aryVals As Variant, aryExpected As Variant
    Dim timeStart As Double, timeEnd As Double, n As Long, j As Long
    
    With tst
    
        'Populate model and ExcelSteps
        PopulateSMdl5 tst, .wkbkTest.Sheets(shtMdl)
    
        'Iterate for timing
        n = 50
        timeStart = Timer
        For j = 1 To n
            Set mdl = ExcelSteps.New_mdl
        
            
            'Provision the model
            .Assert tst, mdl.Provision(mdl, .wkbkTest, shtMdl, IsCalc:=True, _
                IsLiteModel:=True, IsSuppHeader:=True)
            
            'Refresh the model
            .Assert tst, mdl.Refresh(mdl)
        Next j
        timeEnd = Timer

        .Assert tst, mdl.rngFormulaRows.Address = "$C$5"
         
        'Check header column formatting
        CheckHeaderColFormat tst, mdl
        .CkStyleMatch tst, mdl.wksht.Cells(5, 6), "Calculation"
         
        'Test Formula value and format correct
        aryVals = Array(mdl.wksht.Cells(5, 6), _
                        "x_" & mdl.wksht.Cells(5, 6).NumberFormat)
        .TestAryVals tst, aryVals, Array(5, "x_0.00")
        
        'Check variable name creation
        aryExpected = Array("$F$2", "$F$3", "$F$5")
        .CkRngNameAddresses tst, Array("side_a", "side_b", "side_c"), aryExpected
        .CkRngNameAddresses tst, Array(shtMdl), Array("$F:$F")
        
        'Cosmetics
        mdl.wksht.Activate
        mdl.wksht.Cells(1, 1).Select
        
        .Update tst, procs
        
        Debug.Print "test_RefreshSMdl5 Refresh: " & Format((timeEnd - timeStart), "0.000") & " seconds for " & n & " iterations"

    End With
End Sub
'-----------------------------------------------------------------------------------------------
' Multi-column, Non-Homed, Header suppressed, Lite model
'
Sub test_RefreshSMdl6(procs)
    Dim tst As New Test: tst.Init tst, "test_RefreshSMdl6"
    Dim mdl As Object: Set mdl = ExcelSteps.New_mdl
    Dim aryVals As Variant, rng As Range
    
    With tst
        PopulateSMdl6 tst, .wkbkTest.Sheets(shtMdl)
        .Assert tst, mdl.Provision(mdl, .wkbkTest, shtMdl, IsLiteModel:=True, IsSuppHeader:=True, _
                                    cellHome:=.wkbkTest.Sheets(shtMdl).Cells(10, 6))
        .Assert tst, mdl.Refresh(mdl)
        
        'Check ranges of populated rows and columns
        aryVals = Array(mdl.rngPopRows.Address, mdl.rngPopCols.Address)
        .TestAryVals tst, aryVals, Array("$H$10,$H$12:$H$13,$H$15", "$K$10:$L$10")
    
        'Check number format specified in ExcelSteps for side_a
        Set rng = Range(mdl.wksht.Cells(12, 11), mdl.wksht.Cells(12, 12))
        .Assert tst, rng.NumberFormat = "0.000"

        'Check header column format and calculated cell format
        CheckHeaderColFormat tst, mdl
        .CkStyleMatch tst, Range(mdl.wksht.Cells(15, 11), mdl.wksht.Cells(15, 12)), "Calculation"
        
        'Check Formula values and number format
        .TestAryVals tst, Array(mdl.wksht.Cells(15, 11), mdl.wksht.Cells(15, 12)), Array(5, 10)
        .Assert tst, (mdl.wksht.Cells(15, 11).NumberFormat = "0.00")
        .Assert tst, (mdl.wksht.Cells(15, 12).NumberFormat = "0.00")

        'Check variable name creation
        .CkRngNameAddresses tst, Array("side_a", "side_b", "side_c"), _
                                  Array("$12:$12", "$13:$13", "$15:$15")
                                  
        .CkRngNameAddresses tst, Array("Triangle1", "Triangle2"), _
                                  Array("$K:$K", "$L:$L")
        
        'Cosmetics
        mdl.wksht.Activate
        mdl.wksht.Cells(1, 1).Select
        .Update tst, procs
    End With
End Sub
'-----------------------------------------------------------------------------------------------
' Multi-column, Non-Homed, Header suppressed, Lite model - Hide Variable names for aesthetics
' Refactored 10/21/24
'
Sub test_RefreshSMdl6_HiddenVars(procs)
    Dim tst As New Test: tst.Init tst, "test_RefreshSMdl6_HiddenVars"
    Dim mdl As Object: Set mdl = ExcelSteps.New_mdl
    Dim aryVals As Variant, rng As Range
    
    With tst
        'Populate model and Hide the variable names column
        PopulateSMdl6 tst, .wkbkTest.Sheets(shtMdl)
        .wkbkTest.Sheets(shtMdl).Columns(8).EntireColumn.Hidden = True

        'Provision and Refresh
        .Assert tst, mdl.Provision(mdl, .wkbkTest, shtMdl, IsLiteModel:=True, IsSuppHeader:=True, _
                                    cellHome:=.wkbkTest.Sheets(shtMdl).Cells(10, 6))
        .Assert tst, mdl.Refresh(mdl)
        
        'Check ranges of populated rows and columns
        aryVals = Array(mdl.rngPopRows.Address, mdl.rngPopCols.Address)
        .TestAryVals tst, aryVals, Array("$H$10,$H$12:$H$13,$H$15", "$K$10:$L$10")
    
        'Check number format specified in ExcelSteps for side_a
        Set rng = Range(mdl.wksht.Cells(12, 11), mdl.wksht.Cells(12, 12))
        .Assert tst, rng.NumberFormat = "0.000"

        'Check header column format and calculated cell format
        CheckHeaderColFormat tst, mdl
        .CkStyleMatch tst, Range(mdl.wksht.Cells(15, 11), mdl.wksht.Cells(15, 12)), "Calculation"
        
        'Check Formula values and number format
        .TestAryVals tst, Array(mdl.wksht.Cells(15, 11), mdl.wksht.Cells(15, 12)), Array(5, 10)
        .Assert tst, (mdl.wksht.Cells(15, 11).NumberFormat = "0.00")
        .Assert tst, (mdl.wksht.Cells(15, 12).NumberFormat = "0.00")

        'Check variable name creation
        .CkRngNameAddresses tst, Array("side_a", "side_b", "side_c"), _
                                  Array("$12:$12", "$13:$13", "$15:$15")
                                  
        .CkRngNameAddresses tst, Array("Triangle1", "Triangle2"), _
                                  Array("$K:$K", "$L:$L")
        
        'Cosmetics
        mdl.wksht.Activate
        mdl.wksht.Cells(1, 1).Select
        .Update tst, procs
    End With
End Sub
'-----------------------------------------------------------------------------------------------
'Write a model specification setting to Settings sheet
'JDL 12/14/21; Updated 7/21/23; Refactored 10/21/24
'
Sub test_WriteMdlSetting(procs)
    Dim tst As New Test: tst.Init tst, "test_WriteMdlSetting"
    Dim aryVals As Variant, aryExpected As Variant, sVal As String, sName As String
    
    With tst
        
        'Clear/initialize Settings sheet
        With .wkbkTest.Sheets(shtSettings)
            .Cells.Clear
            Range(.Cells(1, 1), .Cells(1, 2)) = Split("setting_name|value", "|")
        End With
        
        'Write a setting
        sVal = "SMdlDash:4,9,0:T:T:T:F:T"
        sName = "SMdlDash"
        PopulateSMdlSetting .wkbkTest, sName, sVal
        
        'Check setting written correctly
        aryVals = Array(.wkbkTest.Sheets(shtSettings).Cells(2, 1), _
                        .wkbkTest.Sheets(shtSettings).Cells(2, 2))
        .TestAryVals tst, aryVals, Array(sName, sVal)
        .Update tst, procs
    End With
End Sub
'-----------------------------------------------------------------------------------------------
' Provision Calculator, Config Stored in Settings; specified nrows
' Updated JDL 7/21/23; Refactored 10/21/24
'
Sub test_ProvisionSMdl7(procs)
    Dim tst As New Test: tst.Init tst, "test_ProvisionSMdl7"
    Dim mdl As Object: Set mdl = ExcelSteps.New_mdl
    Dim aryVals As Variant, aryExpected As Variant
    
    With tst
        PopulateSMdl7 tst, .wkbkTest.Sheets(shtMdl)
        .Assert tst, mdl.Provision(mdl, .wkbkTest, MdlName:="SMdlDash")
        
        aryVals = Array(mdl.rngPopRows.Count, _
                        mdl.IsCalc, _
                        mdl.IsSuppHeader, _
                        mdl.IsRngNames, _
                        mdl.rngPopCols.Address, _
                        mdl.rngrows.Address, _
                        mdl.cellHome.Address)
        aryExpected = Array(3, _
                            True, _
                            True, _
                            True, _
                            "$N:$N", _
                            "$4:$8", _
                            "$I$4")
        .TestAryVals tst, aryVals, aryExpected
        
        tst.Update tst, procs
    End With
End Sub
'-----------------------------------------------------------------------------------------------
' Calculator, Config Stored in Settings; specified nrows
' Updated 7/19/23 JDL; Refactored 10/21/24
'
Sub test_RefreshSMdl7(procs)
    Dim tst As New Test: tst.Init tst, "test_RefreshSMdl7"
    Dim mdl As Object: Set mdl = ExcelSteps.New_mdl
    Dim aryVals As Variant
    
    With tst
        PopulateSMdl7 tst, .wkbkTest.Sheets(shtMdl)
        .Assert tst, mdl.Provision(mdl, .wkbkTest, MdlName:="SMdlDash")
        .Assert tst, mdl.Refresh(mdl)
        mdl.ApplyBorderAroundModel mdl, IsBufferRow:=False, IsBufferCol:=True
        
        'Check that nrows=5 Defn param set rngRows and nrows properly
        .TestAryVals tst, Array(mdl.rngrows.Address, mdl.nRows), Array("$4:$8", 5)

            
        'Test Header column and calculated cell styles; check ExcelSteps specified format
        CheckHeaderColFormat tst, mdl
        .Assert tst, (mdl.wksht.Cells(7, 14).Style = "Calculation")
        
        aryVals = Array(mdl.wksht.Cells(7, 14).Style, _
                        "x_" & mdl.wksht.Cells(4, 14).NumberFormat, _
                        mdl.wksht.Cells(7, 14).Value, _
                        "x_" & mdl.wksht.Cells(7, 14).NumberFormat)
                        
        .TestAryVals tst, aryVals, Array("Calculation", "x_0.000", 5, "x_0.00")
                        
        'Test variable name creation
        .CkRngNameAddresses tst, Array("side_a", "side_b", "side_c"), _
                                  Array("$N$4", "$N$5", "$N$7")
                                  
        .CkRngNameAddresses tst, Array("SMdlDash"), Array("$N:$N")
        .Update tst, procs
    End With
    
    'Cosmetics
    mdl.wksht.Activate
    mdl.wksht.Cells(1, 1).Select
End Sub
'-----------------------------------------------------------------------------------------------
'-----------------------------------------------------------------------------------------------
'-----------------------------------------------------------------------------------------------
'procs.mdlDropdowns procedure
'-----------------------------------------------------------------------------------------------
'-----------------------------------------------------------------------------------------------
'-----------------------------------------------------------------------------------------------
' Populate a named list to use for dropdown testing with Scenario Models
' JDL updated 7/21/23; Refactored 10/21/24
'
Sub test_PopulateNamedList(procs)
    Dim tst As New Test: tst.Init tst, "test_PopulateNamedList"
    Dim aryVals As Variant
    
    With tst
    
        PopulateNamedList .wkbkTest, IsClear:=True
        aryVals = Array(.wkbkTest.Sheets(shtMdl).Cells(1, 20), _
                        .wkbkTest.Names("list_test").RefersToRange.Address)
                        
        .TestAryVals tst, aryVals, Array("list_test", "$T$2:$T$5")
        .Update tst, procs
    End With
End Sub
'-----------------------------------------------------------------------------
'SMdl4a Default "calculator"/single column; add list validation to side_a
' Refactored 10/21/24
'
Sub test_AddDropdownSMdl4a(procs)
    Dim tst As New Test: tst.Init tst, "test_AddDropdownSMdl4a"
    Dim mdl As Object: Set mdl = ExcelSteps.New_mdl
    
    With tst
    
        'Add a single-column Scenario model and Provision
        PopulateSMdl4a .wkbkTest.Sheets(shtMdl)
        .wkbkTest.Sheets(shtMdl).Cells(3, 7) = "Input Dropdown:list_test"
        mdl.Provision mdl, .wkbkTest, shtMdl, IsCalc:=True, IsMdlNmPrefix:=True
        
        'Add a named list "list_test" to use as dropdown
        PopulateNamedList .wkbkTest, IsClear:=False
        
        'Refresh the model and check list validation
        .Assert tst, mdl.Refresh(mdl)
        .Assert tst, mdl.wksht.Cells(3, 9).Validation.Type = xlValidateList
        .Update tst, procs
    End With
End Sub
'-----------------------------------------------------------------------------
' SMdl2 Default multi-column; add list validation to side_a
' Updated 7/21/23 JDL; Refactored 10/21/24
'
Sub test_AddDropdownSMdl2(procs)
    
    Dim tst As New Test: tst.Init tst, "test_AddDropdownSMdl2"
    Dim mdl As Object: Set mdl = ExcelSteps.New_mdl
    Dim rng As Range
    
    With tst
            
        'Populate values and provision the model
        PopulateSMdl2 .wkbkTest.Sheets(shtMdl)
        .wkbkTest.Sheets(shtMdl).Cells(3, 7) = "Input Dropdown:list_test"
        mdl.Provision mdl, .wkbkTest, shtMdl
        
        'Add a named list "list_test" to use as dropdown
        PopulateNamedList .wkbkTest, IsClear:=False
        
        'Refresh the model and check list validation
        .Assert tst, mdl.Refresh(mdl)
        Set rng = Range(mdl.wksht.Cells(3, 9), mdl.wksht.Cells(3, 10))
        .Assert tst, rng.Validation.Type = xlValidateList
        .Update tst, procs
    End With
End Sub
'-----------------------------------------------------------------------------------------------
' SMdl5 Lite Model; add list validation to side_a
' Refactored 10/21/24
'
Sub test_AddDropdownSMdl5(procs)
    Dim tst As New Test: tst.Init tst, "test_AddDropdownSMdl5"
    Dim mdl As Object: Set mdl = ExcelSteps.New_mdl
    Dim rng As Range, rngValCells As Range
    
    With tst
    
        'Populate Lite model; customize ExcelSteps for side_a dropdown
        PopulateSMdl5_Dropdown tst, .wkbkTest.Sheets(shtMdl)
        
        'Add a named list "list_test" to use as dropdown
        PopulateNamedList .wkbkTest, IsClear:=False
    
        'Provision and Refresh the model
        .Assert tst, mdl.Provision(mdl, .wkbkTest, shtMdl, IsCalc:=True, IsLiteModel:=True, _
                                    IsSuppHeader:=True)
        .Assert tst, mdl.Refresh(mdl)
        
        'Check list validation on side_a
        Set rngValCells = mdl.wksht.Cells.SpecialCells(xlCellTypeAllValidation)
        Set rng = .wkbkTest.Sheets(shtMdl).Cells(2, 6)

        'Check that side_a has list validation
        If Intersect(rng, rngValCells) Is Nothing Then
            .Assert tst, False
            
        'Check validation type and that side_a is only one with validation
        Else
            .Assert tst, rng.Validation.Type = xlValidateList
            .Assert tst, rngValCells.Count = 1
        End If
        .Update tst, procs
    End With
End Sub
'-----------------------------------------------------------------------------------------------
' Check header column formatting
' JDL 7/19/23
'
Sub CheckHeaderColFormat(Test, mdl)
    Dim i As Integer, r As Variant, icol1 As Integer, icol2 As Integer
    
    'Set header col locations relative to cellHome and IsCalc
    With mdl
        icol1 = .cellHome.Column + 3
        icol2 = icol1 + 3
        If .IsLiteModel Then
            icol1 = .cellHome.Column + 2
            icol2 = icol1 + 1
        End If
    End With
    
    With mdl.wksht
        For Each r In mdl.rngPopRows
            i = r.Row
            Test.CkStyleMatch Test, Range(.Cells(i, icol1), .Cells(i, icol2)), "Note"
        Next r
    End With
End Sub






