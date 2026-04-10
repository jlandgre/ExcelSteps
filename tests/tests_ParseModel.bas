Attribute VB_Name = "tests_ParseModel"
Option Explicit
'Version 4/10/26; all pass
Const Row1WithDesc As String = "Scenario Description,Scenario,side_a,side_b,side_c"
Const Row2WithDesc As String = "T1,Triangle1,3,4,5"
Const Row3WithDesc As String = "T2,Triangle2,6,8,10"

Const Row1NoDesc As String = "Scenario,side_a,side_b,side_c"
Const Row2NoDesc As String = "Triangle1,3,4,5"
Const Row3NoDesc As String = "Triangle2,6,8,10"
'---------------------------------------------------------------------------------------
' Tests of ParseModel (ExcelSteps.modParseSM module)
' 11/19/25; Updated 1/28/26
'
Sub TestDriver_ParseModel()
    Dim procs As New Procedures, AllEnabled As Boolean
    
    With procs
        .Init procs, ThisWorkbook, "tests_ParseModel", "tests_ParseModel"
        SetApplEnvir False, False, xlCalculationManual
        
        'Enable testing of all or individual procedures
        AllEnabled = True
        .ParseModel.Enabled = True
    End With
    
    'Setup procedure group
    With procs.ParseModel
        If .Enabled Or AllEnabled Then
            procs.curProcedure = .Name
            test_ParseDefaultModel procs
            test_ParseDefaultCalcModel procs
            test_ParseSuppHeaderCalcModel procs
            test_ParseSuppHeaderModel procs
            test_ParseWithDeleteFlag procs
            test_ParseNonHomed procs
        End If
    End With
    
    procs.EvalOverall procs
End Sub
'-----------------------------------------------------------------------------------------------
'Multi-column default model
' JDL 11/17/25; updated 4/10/26
'
Sub test_ParseDefaultModel(procs)
    Dim tst As New Test: tst.Init tst, "test_ParseDefaultModel"
    Dim mdl As Object: Set mdl = ExcelSteps.New_mdl
    Dim aryVals As Variant, wkshtrc As Worksheet
    
    With tst
    
        'Create default, multicolumn model (blank out one Description)
        PopulateSMdl2 .wkbkTest.Sheets(shtMdl)
        
        'Blank out one Description and one units cell to confirm use of colrngVarNames
        .wkbkTest.Sheets(shtMdl).Cells(4, 3).ClearContents
        .wkbkTest.Sheets(shtMdl).Cells(4, 5).ClearContents
        
        .Assert tst, mdl.Provision(mdl, .wkbkTest, shtMdl)
        
        .Assert tst, mdl.Refresh(mdl)
        Application.Calculate
        
        'Call Parse function and check values
        ExcelSteps.ParseMdl mdl
        CheckValsWithDescColumn tst
        
        If Not mdl.wkbkParsed Is Nothing Then mdl.wkbkParsed.Close False
        .Update tst, procs
    End With
End Sub
'-----------------------------------------------------------------------------------------------
'Calculator default model
' JDL 11/17/25
'
Sub test_ParseDefaultCalcModel(procs)
    Dim tst As New Test: tst.Init tst, "test_ParseDefaultCalcModel"
    Dim mdl As Object: Set mdl = ExcelSteps.New_mdl
    
    With tst
    
        'Create default, calculator model
        PopulateSMdl4b .wkbkTest.Sheets(shtMdl)
        .Assert tst, mdl.Provision(mdl, .wkbkTest, shtMdl, IsCalc:=True)
        .Assert tst, mdl.Refresh(mdl)
        Application.Calculate
        
        'Call Parse function and check values
        ExcelSteps.ParseMdl mdl
        CheckValsCalcWithDesc tst
                
        If Not mdl.wkbkParsed Is Nothing Then mdl.wkbkParsed.Close False
        .Update tst, procs
    End With
End Sub
'-----------------------------------------------------------------------------------------------
'Calculator model; header suppressed
' Blank Row 1 (no Scenario variable / "Calculator" values
' JDL 11/17/25
'
Sub test_ParseSuppHeaderCalcModel(procs)
    Dim tst As New Test: tst.Init tst, "test_ParseSuppHeaderCalcModel"
    Dim mdl As Object: Set mdl = ExcelSteps.New_mdl
    
    With tst
    
        'Create header-suppressed calculator model
        PopulateSMdl4c .wkbkTest.Sheets(shtMdl)
        .Assert tst, mdl.Provision(mdl, .wkbkTest, shtMdl, IsCalc:=True, IsSuppHeader:=True)
        .Assert tst, mdl.Refresh(mdl)
        Application.Calculate
        
        'Call Parse function and check values
        ExcelSteps.ParseMdl mdl
        CheckValsCalc tst
        
        If Not mdl.wkbkParsed Is Nothing Then mdl.wkbkParsed.Close False
        .Update tst, procs
    End With
End Sub
'-----------------------------------------------------------------------------------------------
'Multicolumn model; header suppressed
' JDL 11/17/25
'
Sub test_ParseSuppHeaderModel(procs)
    Dim tst As New Test: tst.Init tst, "test_ParseSuppHeaderModel"
    Dim mdl As Object: Set mdl = ExcelSteps.New_mdl
    
    With tst
    
        'Create header-suppressed, multicolumn model
        PopulateSMdl2b .wkbkTest.Sheets(shtMdl)
        .Assert tst, mdl.Provision(mdl, .wkbkTest, shtMdl, IsSuppHeader:=True)
        .Assert tst, mdl.Refresh(mdl)
        Application.Calculate
        
        'Call Parse function and check values
        ExcelSteps.ParseMdl mdl
        CheckValsNoDescColumn tst
        
        If Not mdl.wkbkParsed Is Nothing Then mdl.wkbkParsed.Close False
        .Update tst, procs
    End With
End Sub
'-----------------------------------------------------------------------------------------------
'Multi-column default model with delete flag for a variable
' JDL 11/17/25
'
Sub test_ParseWithDeleteFlag(procs)
    Dim tst As New Test: tst.Init tst, "test_ParseWithDeleteFlag"
    Dim mdl As Object: Set mdl = ExcelSteps.New_mdl
    
    With tst
    
        'Create default, multicolumn model
        PopulateSMdl2c .wkbkTest.Sheets(shtMdl)
        .Assert tst, mdl.Provision(mdl, .wkbkTest, shtMdl)
        .Assert tst, mdl.Refresh(mdl)
        Application.Calculate
        
        'Call Parse function and check values
        ExcelSteps.ParseMdl mdl
        CheckValsDeleteFlag tst
        
        If Not mdl.wkbkParsed Is Nothing Then mdl.wkbkParsed.Close False
        .Update tst, procs
    End With
End Sub
'-----------------------------------------------------------------------------------------------
'Lite, Non-homed model; Header suppressed
' JDL 11/17/25
'
Sub test_ParseNonHomed(procs)
    Dim tst As New Test: tst.Init tst, "test_ParseWithDeleteFlag"
    Dim mdl As Object: Set mdl = ExcelSteps.New_mdl
    
    With tst
    
        'Create non-homed, header-suppressed, multicolumn model
        PopulateSMdl6 tst, .wkbkTest.Sheets(shtMdl)
        .Assert tst, mdl.Provision(mdl, .wkbkTest, shtMdl, IsLiteModel:=True, _
            IsSuppHeader:=True, cellHome:=.wkbkTest.Sheets(shtMdl).Cells(10, 6))
        .Assert tst, mdl.Refresh(mdl)
        Application.Calculate
        
        'Call Parse function and check values
        ExcelSteps.ParseMdl mdl
        CheckValsNoDescColumn tst
        
        If Not mdl.wkbkParsed Is Nothing Then mdl.wkbkParsed.Close False
        .Update tst, procs
    End With
End Sub
Sub CheckValsWithDescColumn(tst)
    Dim wkshtrc As Worksheet, aryVals As Variant

    With tst
        Set wkshtrc = ActiveSheet
        aryVals = Range(wkshtrc.Cells(1, 1), wkshtrc.Cells(1, 5)).Value2
        .CompareRangeToExpected tst, aryVals, Split(Row1WithDesc, ",")
        
        aryVals = Range(wkshtrc.Cells(2, 1), wkshtrc.Cells(2, 5)).Value2
        .CompareRangeToExpected tst, aryVals, Split(Row2WithDesc, ",")
        
        aryVals = Range(wkshtrc.Cells(3, 1), wkshtrc.Cells(3, 5)).Value2
        .CompareRangeToExpected tst, aryVals, Split(Row3WithDesc, ",")
    End With
End Sub
Sub CheckValsNoDescColumn(tst)
    Dim wkshtrc As Worksheet, aryVals As Variant

    With tst
        Set wkshtrc = ActiveSheet
        aryVals = Range(wkshtrc.Cells(1, 1), wkshtrc.Cells(1, 4)).Value2
        .CompareRangeToExpected tst, aryVals, Split(Row1NoDesc, ",")
        
        aryVals = Range(wkshtrc.Cells(2, 1), wkshtrc.Cells(2, 4)).Value2
        .CompareRangeToExpected tst, aryVals, Split(Row2NoDesc, ",")
          
        aryVals = Range(wkshtrc.Cells(3, 1), wkshtrc.Cells(3, 4)).Value2
        .CompareRangeToExpected tst, aryVals, Split(Row3NoDesc, ",")
    End With
End Sub
Sub CheckValsCalc(tst)
    Dim wkshtrc As Worksheet, aryVals As Variant

    With tst
        Set wkshtrc = ActiveSheet
        aryVals = Range(wkshtrc.Cells(1, 1), wkshtrc.Cells(1, 4)).Value2
        .CompareRangeToExpected tst, aryVals, Split(Row1NoDesc, ",")
        
        aryVals = Range(wkshtrc.Cells(2, 1), wkshtrc.Cells(2, 4)).Value2
        .CompareRangeToExpected tst, aryVals, Split("Calculator,3,4,5", ",")
    End With
End Sub
Sub CheckValsCalcWithDesc(tst)
    Dim wkshtrc As Worksheet, aryVals As Variant

    With tst
        Set wkshtrc = ActiveSheet
        aryVals = Range(wkshtrc.Cells(1, 1), wkshtrc.Cells(1, 5)).Value2
        .CompareRangeToExpected tst, aryVals, Split(Row1WithDesc, ",")
        
        aryVals = Range(wkshtrc.Cells(2, 1), wkshtrc.Cells(2, 5)).Value2
        .CompareRangeToExpected tst, aryVals, Split("T1,Calculator,3,4,5", ",")
    End With
End Sub
Sub CheckValsDeleteFlag(tst)
    Dim wkshtrc As Worksheet, aryVals As Variant

    With tst
        Set wkshtrc = ActiveSheet
        aryVals = Range(wkshtrc.Cells(1, 1), wkshtrc.Cells(1, 4)).Value2
        .CompareRangeToExpected tst, aryVals, Split("Scenario Description,Scenario,side_a,side_b", ",")
        
        aryVals = Range(wkshtrc.Cells(2, 1), wkshtrc.Cells(2, 4)).Value2
        .CompareRangeToExpected tst, aryVals, Split("T1,Triangle1,3,4", ",")
        
        aryVals = Range(wkshtrc.Cells(3, 1), wkshtrc.Cells(3, 4)).Value2
        .CompareRangeToExpected tst, aryVals, Split("T2,Triangle2,6,8", ",")
    End With
End Sub


