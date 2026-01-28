Attribute VB_Name = "modPopulateMdl"
Option Explicit
'-----------------------------------------------------------------------------------------------
'This section Creates named list for validating dropdown feature
'-----------------------------------------------------------------------------------------------
'Purpose:   Create a named list to right of Scenario Model area
'
'Created:   6/29/22 JDL
'
Sub PopulateNamedList(wkbk, Optional IsClear = True)
    Dim ary As Variant, sNameString As String

    If Not SheetExists(wkbk, "SMdl") Then AddSheet wkbk, "SMdl", "tst_results"
    With wkbk.Sheets("SMdl")
    
        'Clear cells and names if instructed
        If IsClear Then
            .Cells.Clear
            DeleteAllWorkbookNames wkbk
        End If
             
        'Write list name and values to SMdl
        ary = Split("list_test,No Selection,3,6,8", ",")
        Range(.Cells(1, 20), .Cells(5, 20)) = Application.Transpose(ary)
        
        'Name the list
        sNameString = MakeRefNameString("SMdl", 2, 5, 20, 20)
        MakeXLName wkbk, xlName(.Cells(1, 20).Value), sNameString
        
        ShadeYellow Range(.Cells(2, 20), .Cells(5, 20))

    End With
End Sub
'-----------------------------------------------------------------------------------------------
'This section populates various Scenario Models
'-----------------------------------------------------------------------------------------------
'Purpose:   Populate Settings with a Scenario Model definition
'
'Created:   12/14/21 JDL      Modified: 1/3/22 Add IsClear arg
'                                       7/19/23 remove mdl_ Setting name prefix
'
Sub PopulateSMdlSetting(wkbk, MdlName, sSetting, Optional IsClear = True)
    Dim i As Integer

    If Not SheetExists(wkbk, shtSettings) Then AddSheet wkbk, "Settings_", "Tests_Util"
    With wkbk.Sheets(shtSettings)
        If IsClear Then .Cells.Clear
        Range(.Cells(1, 1), .Cells(1, 2)) = Split("setting_name|value", "|")
        
        i = rngLastPopCell(.Cells(1, 1), xlDown).Offset(1, 0).Row
        Range(.Cells(i, 1), .Cells(i, 2)) = Array(MdlName, sSetting)
    End With
End Sub
'-----------------------------------------------------------------------------------------------
'Clear test sheet and populate with sample data
' Default multicolumn model architecture
'
'Created:   11/22/21 JDL      Modified:
'
Sub PopulateSMdl1(wksht)
    Dim ary As Variant, s As String, LstVals() As Variant, i As Integer
    
    With wksht
        ClearTestSheetAndNames wksht
        
        'Write Header
        s = "Grp,Subgrp,Description,Variable Names,Units,Number Fmt,Formula/Row Type,,T1"
        Range(.Cells(1, 1), .Cells(1, 9)) = Split(s, ",")
        
        'Populate lists of column values
        ReDim LstVals(1 To 9)
        LstVals(1) = "Triangles,,,,"
        LstVals(2) = ",,,,"
        LstVals(3) = "Scenario Name,Side A,Side B,,Hypotenuse"
        LstVals(4) = "Scenario,side_a,side_b,,side_c"
        LstVals(5) = ",mm,mm,,mm"
        LstVals(6) = ",0,0,,0.00"
        LstVals(7) = "Input,Input,Input,,=(@side_a^2 + @side_b^2)^0.5"
        LstVals(8) = ",,,,"
        LstVals(9) = "Triangle1,3,4,,"
        
        'Set Number format of text columns
        Range(.Cells(2, 6), .Cells(8, 6)).NumberFormat = "@"
        Range(.Cells(2, 7), .Cells(8, 7)).NumberFormat = "@"
        
        'Populate by columns
        For i = 1 To 9
            ary = Split(LstVals(i), ",")
            Range(.Cells(2, i), .Cells(6, i)) = WorksheetFunction.Transpose(ary)
        Next i
    End With
End Sub
'-----------------------------------------------------------------------------------------------
'Multi-column default model
Sub PopulateSMdl2(wksht)
    Dim ary As Variant
    PopulateSMdl1 wksht
    ary = Split("T2,Triangle2,6,8", ",")
    Range(wksht.Cells(1, 10), wksht.Cells(4, 10)) = WorksheetFunction.Transpose(ary)
End Sub
'-----------------------------------------------------------------------------------------------
'Multicolumn, Header suppressed model
Sub PopulateSMdl2b(wksht)
    PopulateSMdl2 wksht
    wksht.Rows(1).Delete
End Sub
'-----------------------------------------------------------------------------------------------
'Multi-column default model with delete flag for side_c variable
Sub PopulateSMdl2c(wksht)
    Dim ary As Variant
    PopulateSMdl1 wksht
    ary = Split("T2,Triangle2,6,8", ",")
    Range(wksht.Cells(1, 10), wksht.Cells(4, 10)) = WorksheetFunction.Transpose(ary)
    wksht.Cells(6, 8).Value2 = "d"
End Sub

'-----------------------------------------------------------------------------------------------
'Purpose:   Model with variation in which rows are populated
'
'Created:   12/5/24 JDL      Modified:
'
Sub PopulateSMdl2a(wksht)
    Dim ary As Variant, s As String, LstVals() As Variant, i As Integer
    
    With wksht
        ClearTestSheetAndNames wksht
        
        'Write Header
        s = "Grp,Subgrp,Description,Variable Names,Units,Number Fmt,Formula/Row Type,,T1"
        Range(.Cells(1, 1), .Cells(1, 9)) = Split(s, ",")
        
        'Populate lists of column values
        ReDim LstVals(1 To 9)
        LstVals(1) = "Triangles,,,,"
        LstVals(2) = ",,,,"
        LstVals(3) = ",,,,Side A"
        LstVals(4) = ",,,,side_a"
        LstVals(5) = ",,,,,mm"
        LstVals(6) = ",,,,0.00"
        LstVals(7) = ",,,,Input"
        LstVals(8) = ",,,,"
        LstVals(9) = "Triangle1,,,,3"
        
        'Set Number format of text columns
        Range(.Cells(2, 6), .Cells(8, 6)).NumberFormat = "@"
        Range(.Cells(2, 7), .Cells(8, 7)).NumberFormat = "@"
        
        'Populate by columns
        For i = 1 To 9
            ary = Split(LstVals(i), ",")
            Range(.Cells(2, i), .Cells(6, i)) = WorksheetFunction.Transpose(ary)
        Next i
    End With
End Sub
'-----------------------------------------------------------------------------------------------
'Multi-column default model (Non-Contiguous columns)
Sub PopulateSMdl3(wksht)
    PopulateSMdl2 wksht
    wksht.Columns(10).Insert
    wksht.Cells(1, 11) = "T2"
End Sub
'-----------------------------------------------------------------------------------------------
'Calculator, Header suppressed model
Sub PopulateSMdl4(wksht)
    PopulateSMdl1 wksht
    wksht.Rows(1).Delete
    wksht.Rows(1).Clear
End Sub
'-----------------------------------------------------------------------------------------------
'Calculator model; sheet name prefix (e.g. non-default)
Sub PopulateSMdl4a(wksht)
    PopulateSMdl1 wksht
    wksht.Cells(6, 7) = "=(smdl_side_a^2 + smdl_side_b^2)^0.5"
    Range(wksht.Cells(2, 3), wksht.Cells(2, 9)).Clear
    wksht.Cells(2, 9).Value2 = "Calculator"
End Sub
'-----------------------------------------------------------------------------------------------
'Default Calculator model
Sub PopulateSMdl4b(wksht)
    PopulateSMdl1 wksht
    Range(wksht.Cells(2, 3), wksht.Cells(2, 9)).Clear
    wksht.Cells(2, 9).Value2 = "Calculator"
End Sub
'-----------------------------------------------------------------------------------------------
'Calculator, Header suppressed model
Sub PopulateSMdl4c(wksht)
    PopulateSMdl1 wksht
    wksht.Rows(1).Delete
    wksht.Rows(1).Clear
End Sub
'-----------------------------------------------------------------------------------------------
'Lite, calculator model with header suppressed
Sub PopulateSMdl5(tst, wksht)
    Dim tbls As Object, refr As Object
    
    PopulateSMdl1 wksht
    With wksht
        .Rows(1).Delete
        .Rows(1).Clear
    
        'Delete number format and Formula columns to make Lite model
        .Columns(6).Delete
        .Columns(6).Delete
        .Columns(1).Delete
    End With
    
    'Recreate and populate ExcelSteps
    PrepBlankStepsForTesting tst.wkbkTest, refr, tbls
    PopulateStepsSMdl tst.wkbkTest, "SMdl"
End Sub
'-----------------------------------------------------------------------------------------------
'Add a dropdown list instruction to ExcelSteps recipe
Sub PopulateSMdl5_Dropdown(Test, wksht)
    
    PopulateSMdl5 Test, wksht
    With Test.wkbkTest.Sheets(shtSteps)
        .Cells(2, 3) = "Col_Dropdown"
        .Cells(2, 4) = "list_test"
    End With
End Sub

'-----------------------------------------------------------------------------------------------
'Non-homed Lite, multi-column model with header suppressed
Sub PopulateSMdl6(Test, wksht)
    Dim tbls As Object, refr As Object
    
    PopulateSMdl1 wksht
    With wksht
        .Rows(1).Delete
        Range(.Cells(1, 1), .Cells(1, 3)).Clear
        .Cells(1, 9) = "Triangle1"
        .Cells(1, 10) = "Triangle2"
        .Rows(2).Insert
        .Cells(3, 10) = 6
        .Cells(4, 10) = 8
        
        'Delete number format and Formula columns; Delete Sub-group column
        .Columns(6).Delete
        .Columns(6).Delete
        .Columns(1).Delete
        
        'Insert white space rows/columns to create non-homed table
        Range(.Rows(1), .Rows(9)).Insert
        Range(.Columns(1), .Columns(5)).Insert
    End With
        
    'Recreate and populate ExcelSteps
    PrepBlankStepsForTesting Test.wkbkTest, refr, tbls
    PopulateStepsSMdl Test.wkbkTest, "SMdl"
End Sub
'-----------------------------------------------------------------------------------------------
'Lite, calculator model with header suppressed - Non-homed
Sub PopulateSMdl7(Test, wksht)
    Dim tbls As Object, refr As Object
    
    PopulateSMdl1 wksht
    With wksht
        .Rows(1).Delete
        .Rows(1).Clear
    
        'Delete number format and Formula columns
        .Columns(6).Delete
        .Columns(6).Delete
    
        'Move the table for a (4,9) cellHome in Lite model
        Range(.Rows(1), .Rows(2)).Insert
        Range(.Columns(1), .Columns(7)).Insert
    End With
    
    'Recreate and populate ExcelSteps
    PrepBlankStepsForTesting Test.wkbkTest, refr, tbls
    PopulateStepsSMdl Test.wkbkTest, "SMdlDash"
    
    'Populate a Setting with model's definition
    PopulateSMdlSetting Test.wkbkTest, "SMdlDash", "SMdl:4,9:5:T:T:T:F:T"
End Sub
'-----------------------------------------------------------------------------------------------
Sub PopulateStepsSMdl(wkbk, MdlName)
    With wkbk.Sheets(shtSteps)
        .Cells(2, 1) = MdlName
        .Cells(2, 2) = "side_a"
        .Cells(2, 8) = "0.000"
        
        .Cells(3, 1) = MdlName
        .Cells(3, 2) = "side_c"
        .Cells(3, 4) = "=(side_a^2 + side_b^2)^0.5"
        .Cells(3, 8) = "0.00"
    End With
End Sub
'-----------------------------------------------------------------------------------------------
Sub ClearTestSheetAndNames(wksht)
    With wksht
    
        'Turn off sheet's filter if on and clear cells
        .AutoFilterMode = False
        RevealWkshtCells wksht
        .Cells.Clear
        Range(.Columns(1), .Columns(20)).Delete
        
        'Clear workbook names
        DeleteAllWorkbookNames wksht.Parent
        
        wksht.Activate
        Cells(1, 1).Select
    End With
End Sub
'-----------------------------------------------------------------------------------------------
'Model with four, non-contiguous variable blocks for .mdlRefreshSpeed testing
'JDL 7/15/25
'
Sub PopulateMdlSpeedup(wksht)
    Dim ary As Variant, s As String, LstVals() As Variant, i As Integer, j As Integer
    Dim rowBlockHome As Integer
    
    With wksht
        ClearTestSheetAndNames wksht
        
        'Write Header and Scenario Names values
        s = "Grp,Subgrp,Description,Variable Names,Units,Number Fmt,Formula/Row Type,,T1"
        Range(.Cells(1, 1), .Cells(1, 9)) = Split(s, ",")
        s = "Triangles,,Scenario_Names,Scenario,,,Input,,Triangle1"
        Range(.Cells(2, 1), .Cells(2, 9)) = Split(s, ",")
        
        'Set Number format of text columns
        .Columns(6).NumberFormat = "@"
        .Columns(7).NumberFormat = "@"
        
        For j = 1 To 4
            rowBlockHome = 3 + 4 * (j - 1)
            
            'Populate lists of column values
            ReDim LstVals(1 To 7)
            LstVals(1) = "Side A,Side B,Hypotenuse"
            LstVals(2) = "side_a_" & j & ",side_b_" & j & ",side_c_" & j
            LstVals(3) = "mm,mm,mm"
            LstVals(4) = "0,0,0.00"
            LstVals(5) = "Input,Input,=(@side_a_" & j & "^2 + @side_b_" & j & "^2)^0.5"
            LstVals(6) = ",,"
            LstVals(7) = "3,4,"
        
            'Populate by columns
            For i = 1 To 7
                ary = Split(LstVals(i), ",")
                Range(.Cells(rowBlockHome, i + 2), .Cells(rowBlockHome + 2, i + 2)) = _
                    WorksheetFunction.Transpose(ary)
            Next i
        Next j
    End With
End Sub

'-----------------------------------------------------------------------------------------------
'Model with contiguous formula rows; xxx stop 15:20 modify to have non-contiguous multiple columns
'with two or more triangles
'JDL 7/15/25
'
Sub PopulateMdlSpeedup2(wksht)
    Dim ary As Variant, s As String, LstVals() As Variant, i As Integer, j As Integer
    Dim rowBlockHome As Integer
    
    With wksht
        ClearTestSheetAndNames wksht
        
        'Write Header and Scenario Names values
        s = "Grp,Subgrp,Description,Variable Names,Units,Number Fmt,Formula/Row Type,,T1,,T2,T3"
        Range(.Cells(1, 1), .Cells(1, 12)) = Split(s, ",")
        s = "Triangles,,Scenario_Names,Scenario,,,Input,,Triangle1,,Triangle2,Triangle3"
        Range(.Cells(2, 1), .Cells(2, 12)) = Split(s, ",")
        
        'Set Number format of text columns
        .Columns(6).NumberFormat = "@"
        .Columns(7).NumberFormat = "@"
        
        For j = 1 To 4
            rowBlockHome = 3 + 3 * (j - 1)
            
            'Populate lists of column values
            ReDim LstVals(1 To 10)
            LstVals(1) = "Side A,Side B"
            LstVals(2) = "side_a_" & j & ",side_b_" & j
            LstVals(3) = "mm,mm"
            LstVals(4) = "0,0"
            LstVals(5) = "Input,Input"
            LstVals(6) = ","
            LstVals(7) = "3,4"
            LstVals(8) = ","
            LstVals(9) = "6,8"
            LstVals(10) = "12,16"
        
            'Populate by columns
            For i = 1 To 10
                ary = Split(LstVals(i), ",")
                Range(.Cells(rowBlockHome, i + 2), .Cells(rowBlockHome + 1, i + 2)) = _
                    WorksheetFunction.Transpose(ary)
            Next i
        Next j
    
        'Put all formula-containing rows in a contiguous block (same .Areas for iteration)
        rowBlockHome = 15
        
        'Populate lists of column values
        ReDim LstVals(1 To 5)
        LstVals(1) = "Hypotenuse,Hypotenuse,,Hypotenuse,Hypotenuse"
        LstVals(2) = "side_c_1,side_c_2,,side_c_3,side_c_4"
        
        LstVals(3) = "mm,mm,,mm,mm"
        LstVals(4) = "0.00,0.00,,0.00,0.00"
        LstVals(5) = "=(@side_a_1^2 + @side_b_1^2)^0.5,=(@side_a_2^2 + @side_b_2^2)^0.5,," _
            & "=(@side_a_3^2 + @side_b_3^2)^0.5,=(@side_a_4^2 + @side_b_4^2)^0.5"
    
        'Populate by columns
        For i = 1 To 5
            ary = Split(LstVals(i), ",")
            Range(.Cells(rowBlockHome, i + 2), .Cells(rowBlockHome + 4, i + 2)) = _
                WorksheetFunction.Transpose(ary)
        Next i
    End With
End Sub
'-----------------------------------------------------------------------------------------------
' Helper sub to load SMdlType2 into mdlDest for mdlImportRow tests
' JDL 1/28/26 (refactored from tests_SwapModels.PopulateSMdlType2ToMdlDest)
'
Sub PopulateForMdlImportRow(tst, mdlDest, tblImp)
    With tst
        'Populate model on sMdl and tbl on tblImport sheet
        PopulateDashAndMdlImportSht tst
    
        'Transfer SMdlType2 to mdlDest region
        .Assert tst, mdlDest.InitSwapModels(mdlDest, tblImp, .wkbkTest, defn_dest)
        .Assert tst, mdlDest.TransferToMdlDest(mdlDest, tblImp, sMdl2, defn_dest)
        
        'Reset model name as it would be w/o TransferToMdlDest mod
        mdlDest.MdlName = "SMdlDest"
    End With
End Sub


