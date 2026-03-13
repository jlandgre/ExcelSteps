Attribute VB_Name = "tests_tblRowsCols"
'Version 10/29/25
Option Explicit
Public Const shtTbl As String = "SMdl"
Public Const defn_test As String = "SMdl:6,2:T:F:T:F:T:F:1:-2:12:5"
Public errs As Object
Public Const shtT_temp As String = "tst_Results"
'-----------------------------------------------------------------------------------------------
'tst driver for SalesSubsetter VBA class methods
'   * To initialize, set a custom name (in VBA Properties) for the code project's VBAProject
'   * Add a reference to the custom project name in tests.xlsm's VBAProject
'   * In this driver sub, procs attributes include named Procedure instances for each Procedure
'   * procs instance also houses .wkbk_testing, .wksht_results and .test_suite_name attributes
'   * shtT is module-level constant for sheet name displaying test results. It is used to
'     initialize procs only
'
'Code style
'   * To create correspondence with project code, test subroutine names should match
'     project function names exactly (e.g. Function ProjFunction() --> sub test_ProjFunction()
'   * tst docstrings should be copy/paste from project function docstrings (with additional
'     lines added to explain test details and scenarios as appropriate
'
' JDL 5/28/24 based on 2021 previous version
'   Updated 10/16/24 to add tblRefreshAPI Procedure
'
Sub TestDriver_TblRowsCols()
    Dim procs As New Procedures, AllEnabled As Boolean
    With procs
        
        'Initialize Procedure objects; Set up tst_Results sheet; Set Procedures attributes
        .Init procs, ThisWorkbook, "tblRowsCols", "tblRowsCols"
        
        'Turn off events and Screenupdataing; calculation Automatic
        SetApplEnvir False, False, xlCalculationAutomatic
        
        'Enable/disable all or groups of tests by procedure
        AllEnabled = True
        .tblInit.Enabled = False
        .tblSetDimensions.Enabled = False
        .tblSetArysNamesRngs.Enabled = False
        .tblFormat.Enabled = False
        .tblProvision.Enabled = False
        .tblRefresh.Enabled = False
        .tblRefreshAPI.Enabled = False
    End With
    
    With procs.tblInit
        If .Enabled Or AllEnabled Then
            procs.curProcedure = .Name
            test_SetIsCustomTbl procs
            test_ReadDefnSetting procs
            test_PopulateCustomTblParams1 procs
            test_PopulateCustomTblParams2 procs
            test_PopulateCustomTblParams3 procs
            test_SetHomedTblParams procs
            test_OverrideWithArgs procs
            test_tblInitProcedure1 procs
            test_tblInitProcedure_sht_CaseError procs
        End If
    End With
    
    With procs.tblSetDimensions
        If .Enabled Or AllEnabled Then
            procs.curProcedure = .Name
            test_InitializeWkshtAndRanges1 procs
            test_InitializeWkshtAndRanges2 procs
            test_SetIsBlankSheet1 procs
            test_SetIsBlankSheet2 procs
            test_SetIsNoData1 procs
            test_SetIsNoData2 procs
            test_SetIsNoData3 procs
            test_SetIsNoData4 procs
            test_SetNRows1 procs
            test_SetNRows2 procs
            test_SetNRows3 procs
            test_SetNCols1 procs
            test_SetNCols2 procs
            test_SetNCols3 procs
            test_SetNCols4 procs
            test_SetRngTable1 procs
            test_SetRngTable2 procs
            test_SetRngTable3 procs
            test_SetDimensions1 procs
            test_SetDimensions2 procs
            test_SetDimensions3 procs
        End If
    End With
    
    With procs.tblSetArysNamesRngs
        If .Enabled Or AllEnabled Then
            procs.curProcedure = .Name
            test_SetAryColRngs procs
            test_SetTblNames procs
            test_NameColumn procs
            test_SetAllColNames procs
        End If
    End With
    
    With procs.tblFormat
        If .Enabled Or AllEnabled Then
            procs.curProcedure = .Name
            '<<tests>>
        End If
    End With
    
    With procs.tblProvision
        If .Enabled Or AllEnabled Then
            procs.curProcedure = .Name
            test_PopulateTbl procs
            test_ProvisionTbl procs
            test_PopulateTbl2 procs
            test_ProvisionTbl2 procs
            test_ProvisionTbl2HeaderOnly procs
            test_ProvisionTbl2EmptyTbl procs
            test_ProvisionTbl2EmptySpec procs
        End If
    End With
    
    'xxx stop 10/17/24 14:30 test_PrepExcelStepsSht has error in refr.PrepExcelStepsSht from tbl.Provision step in procedure
    With procs.tblRefresh
        If .Enabled Or AllEnabled Then
            procs.curProcedure = .Name
            test_PrepExcelStepsSht procs
            test_RefreshTbl2 procs
            test_RefreshTbl3 procs
        End If
    End With
    
    With procs.tblRefreshAPI
        If .Enabled Or AllEnabled Then
            procs.curProcedure = .Name
            test_RefreshTblAPI1 procs
            test_RefreshTblAPI2 procs
            test_RefreshTblAPI3 procs
            test_RefreshTblAPI4 procs
        End If
    End With
    procs.EvalOverall procs

End Sub
'-----------------------------------------------------------------------------------------
'-----------------------------------------------------------------------------------------
' tblRefreshAPI Procedure - modInterface.RefreshAPI validation for tblRowsCols
'-----------------------------------------------------------------------------
'tst Refresh API all-in-one sub in modInterface of ExcelSteps
'Non-default tables:
' Example1: Defn specified; with IsSetTblNames IsSetColNames = True)
' Example2: TblName and Defn specified; with IsSetTblNames IsSetColNames = True)
'
'JDL 10/17/24
'
Sub test_RefreshTblAPI4(procs)
    Dim tst As New Test: tst.Init tst, "test_RefreshTblAPI4"
    Dim defn_test2 As String, rng As Range
    Dim refr As Object, tblSteps As Object
    PopulateTbl2 tst.wkbkTest, shtTbl, rowHome:=5, colHome:=3
    
    'Prep Excel Steps (clear previous and replace)
    PrepBlankStepsForTesting tst.wkbkTest, refr, tblSteps
    PopulateStepsTblRefresh tst.wkbkTest, shtTbl
    
    With tst
    
        'Just Defn specified (tblName is same as sht)
        defn_test2 = "SMdl:5,3:T:T:F:F:F:F:0:-1:0:0"
        .Assert tst, ExcelSteps.RefreshTblAPI(.wkbkTest, IsReplace:=True, IsTblFormat:=True, _
            defn:=defn_test2)
        
        .Assert tst, (tst.wkbkTest.Names("Desc").RefersToRange.Parent.Name = shtTbl)
        .Assert tst, (tst.wkbkTest.Names("Desc").RefersToRange.Address = "$C:$C")
        Set rng = Range(tst.wkbkTest.Sheets(shtTbl).Cells(4, 3), tst.wkbkTest.Sheets(shtTbl).Cells(4, 8))
        .Assert tst, rng.Style = "Accent1"

        'Check for header range name and table name
        .Assert tst, (.wkbkTest.Names(shtTbl & "_Header").RefersToRange.Address = "$4:$4")
        .Assert tst, (.wkbkTest.Names(shtTbl).RefersToRange.Address = "$C:$H")
    
        CheckHomedTableRefreshed tst, rowHome:=5, colHome:=3

        'TblName and Defn specified; Add TblName prefix to column names
        defn_test2 = "SMdl:5,3:T:T:T:F:F:F:0:-1:0:0"
        
        'Add table name prefix to ExcelSteps Insert instruction cell
        With tst.wkbkTest.Sheets(shtSteps)
            .Cells(3, 4).Value = "=@Custom_Data_2 + @Custom_Data_3"
        End With
        
        'Refresh the table
        .Assert tst, ExcelSteps.RefreshTblAPI(.wkbkTest, IsReplace:=True, IsTblFormat:=True, _
            TblName:="Custom", defn:=defn_test2)

        .Assert tst, (tst.wkbkTest.Names("Custom_Desc").RefersToRange.Parent.Name = shtTbl)
        .Assert tst, (.wkbkTest.Names("Custom_Header").RefersToRange.Address = "$4:$4")

        With tst.wkbkTest.Sheets(shtTbl)
            tst.Assert tst, .Cells(5, 3 + 6).Formula = "=@Custom_Data_2 + @Custom_Data_3"
            tst.Assert tst, .Cells(5, 3 + 6).Value = 39.6
        End With

        tst.Update tst, procs
    End With
End Sub
'-----------------------------------------------------------------------------
'tst Refresh API all-in-one sub in modInterface of ExcelSteps
'Non-default table (IsSetTblNames = True) but mis-specified because sht not specified
'
'JDL 10/17/24
'
Sub test_RefreshTblAPI3(procs)
    Dim tst As New Test: tst.Init tst, "test_RefreshTblAPI3"
    Dim msg As String
    PopulateTbl2 tst.wkbkTest, shtTbl
    
    'Initialize error handling to allow checking warning and error message(s)
    Set ExcelSteps.errs = ExcelSteps.New_ErrorHandling
    ExcelSteps.errs.Init wkbkE:=ExcelSteps.ThisWorkbook
    ExcelSteps.errs.IsShowMsgs = False

    With tst
        .Assert tst, ExcelSteps.RefreshTblAPI(.wkbkTest, IsReplace:=True, IsTblFormat:=True, _
            IsSetTblNames:=True)
        
        msg = ExcelSteps.errs.Msgs_accum
        tst.Assert tst, Left(msg, 51) = "The following tblRowsCols object is underspecified."
    
        tst.Update tst, procs
    End With
End Sub '-----------------------------------------------------------------------------
'tst Refresh API all-in-one sub in modInterface of ExcelSteps
'Non-default table (.IsSetTblNames = True)
'
'JDL 10/17/24
'
Sub test_RefreshTblAPI2(procs)
    Dim tst As New Test: tst.Init tst, "test_RefreshTblAPI2"
    Dim refr As Object, tblSteps As Object
    PopulateTbl2 tst.wkbkTest, shtTbl
    
    'Prep Excel Steps (clear previous and replace)
    PrepBlankStepsForTesting tst.wkbkTest, refr, tblSteps
    PopulateStepsTblRefresh tst.wkbkTest, shtTbl

    With tst
        .Assert tst, ExcelSteps.RefreshTblAPI(.wkbkTest, IsReplace:=True, IsTblFormat:=True, _
            sht:=shtTbl, IsSetTblNames:=True)
        CheckRefreshedTable tst
        CheckHomedTableRefreshed tst, rowHome:=2, colHome:=1
        
        'Check for header range name and table name
        .Assert tst, (.wkbkTest.Names(shtTbl & "_Header").RefersToRange.Address = "$1:$1")
        .Assert tst, (.wkbkTest.Names(shtTbl).RefersToRange.Address = "$A:$F")
    
        tst.Update tst, procs
    End With
End Sub
'-----------------------------------------------------------------------------
'tst Refresh API all-in-one sub in modInterface of ExcelSteps
'
'JDL 3/6/23; Moved to procs.tblRefreshAPI section 10/17/24
'
Sub test_RefreshTblAPI1(procs)
    Dim tst As New Test: tst.Init tst, "test_RefreshTblAPI1"
    Dim refr As Object, tblSteps As Object
    PopulateTbl2 tst.wkbkTest, shtTbl
    
    'Prep Excel Steps (clear previous and replace)
    PrepBlankStepsForTesting tst.wkbkTest, refr, tblSteps
    PopulateStepsTblRefresh tst.wkbkTest, shtTbl
    
    'Default table
    tst.Assert tst, ExcelSteps.RefreshTblAPI(tst.wkbkTest, IsReplace:=True, IsTblFormat:=True, sht:=shtTbl)
    CheckRefreshedTable tst
    CheckHomedTableRefreshed tst

    tst.Update tst, procs
End Sub
Sub CheckHomedTableRefreshed(tst, Optional rowHome = 2, Optional colHome = 1)
    With tst.wkbkTest.Sheets(shtTbl)
        tst.Assert tst, .Cells(rowHome, colHome + 4).NumberFormat = "0.000"
        tst.Assert tst, .Cells(rowHome, colHome + 6).NumberFormat = "0.00"
        tst.Assert tst, .Cells(rowHome, colHome + 6).Formula = "=@Data_2 + @Data_3"
        tst.Assert tst, .Cells(rowHome, colHome + 6).Value = 39.6
    End With
End Sub
'-----------------------------------------------------------------------------
' Refresh rows/columns table - Two-row ExcelSteps sheet
' JDL 5/28/24; Modified 10/21/24
'
Sub test_tblRefreshAPI1(procs)
    Dim tst As New Test: tst.Init tst, "test_tblRefreshAPI1"
    Dim refr As Object, tblSteps As Object
    
    With tst
        PopulateTbl2 .wkbkTest, shtTbl
        
        'Prep Excel Steps (clear previous and replace)
        PrepBlankStepsForTesting .wkbkTest, refr, tblSteps
        PopulateStepsTblRefresh .wkbkTest, shtTbl
    End With
    
    With tst.wkbkTest.Sheets(shtTbl)
        tst.Assert tst, .Cells(2, 5).NumberFormat = "0.000"
        tst.Assert tst, .Cells(2, 7).NumberFormat = "0.00"
        tst.Assert tst, .Cells(2, 7).Formula = "=@Data_2 + @Data_3"
        tst.Assert tst, .Cells(2, 7).Value = 39.6
    End With
    
    tst.Update tst, procs
End Sub
'-----------------------------------------------------------------------------------------
' SetArysNamesRngs Procedure
'-----------------------------------------------------------------------------------------
' This section tests SetTblNames method
' JDL 5/23/24; Modified 10/21/24
'
Sub test_SetAllColNames(procs)
    Dim tst As New Test: tst.Init tst, "test_SetAllColNames"
    Dim tbl As Object: Set tbl = ExcelSteps.New_tbl
    
    With tst
        'Run the method being tested and precursors
        PopulateTbl .wkbkTest, shtTbl
        .Assert tst, tbl.Init(tbl, wkbk:=tst.wkbkTest, sht:=shtTbl)
        .Assert tst, tbl.SetDimensions(tbl)
        .Assert tst, tbl.SetAllColNames(tbl)
        
        'Check results of method
        .Assert tst, .wkbkTest.Names("Col_A").RefersTo = "=SMdl!$A:$A"
        .Assert tst, .wkbkTest.Names("Col_B").RefersTo = "=SMdl!$B:$B"
        .Assert tst, .wkbkTest.Names("Col_C").RefersTo = "=SMdl!$C:$C"
    End With

    tst.Update tst, procs
End Sub
'-----------------------------------------------------------------------------------------
' This section tests SetTblNames method
' JDL 5/23/24; Modified 10/21/24
'
Sub test_NameColumn(procs)
    Dim tst As New Test: tst.Init tst, "test_NameColumn"
    Dim tbl As Object: Set tbl = ExcelSteps.New_tbl
    
    With tst
        'Run the method being tested and precursors
        PopulateTbl .wkbkTest, shtTbl
        .Assert tst, tbl.Init(tbl, wkbk:=.wkbkTest, sht:=shtTbl)
        .Assert tst, tbl.NameColumn(tbl, tbl.wkbk.Sheets(tbl.sht).Cells(1, 1))
        
        'Check results of method
        .Assert tst, .wkbkTest.Names("Col_A").RefersTo = "=SMdl!$A:$A"
        
        .Update tst, procs
    End With
End Sub
'-----------------------------------------------------------------------------------------
' This section tests SetTblNames method
' JDL 5/24/24; Modified 10/21/24
'
Sub test_SetTblNames(procs)
    Dim tst As New Test: tst.Init tst, "test_SetTblNames"
    Dim tbl As Object: Set tbl = ExcelSteps.New_tbl
    
    With tst
        'Run the method being tested and precursor methods
        PopulateTbl .wkbkTest, shtTbl
        .Assert tst, tbl.Init(tbl, wkbk:=.wkbkTest, sht:=shtTbl)
        .Assert tst, tbl.SetDimensions(tbl)
        .Assert tst, tbl.SetTblNames(tbl)

        'Check results of method
        .Assert tst, .wkbkTest.Names("SMdl").RefersTo = "=SMdl!$A:$C"
        .Assert tst, .wkbkTest.Names("SMdl_Header").RefersTo = "=SMdl!$1:$1"

        .Update tst, procs
    End With
End Sub
'-----------------------------------------------------------------------------------------
' Set an array of column ranges relative to cellHome
' Added Let/Get Properties to tblRowsCols to make this attribute more robust
' JDL 5/24/24; Modified 10/21/24
'
Sub test_SetAryColRngs(procs)
    Dim tst As New Test: tst.Init tst, "test_SetAryColRngs"
    Dim i As Integer, myArray As Variant, aryExpected As Variant
    Dim tbl As Object: Set tbl = ExcelSteps.New_tbl
    
    With tst

        'Run the methods being tested (use .Assert to check T/F result)
        PopulateTbl .wkbkTest, shtTbl
        .Assert tst, tbl.Init(tbl, wkbk:=.wkbkTest, sht:=shtTbl)
        .Assert tst, tbl.SetDimensions(tbl)
        .Assert tst, tbl.SetAryColRngs(tbl)
                
        'This code works (See extensive JDL Evernote note on this topic)
        If True Then
            myArray = tbl.aryColRngs
            aryExpected = Array("$A:$A", "$B:$B", "$C:$C")

            For i = LBound(myArray) To UBound(myArray)
                'Debug.Print myArray(i).Address
                .Assert tst, myArray(i).Address = aryExpected(i)
            Next i
            
        'This Gives error "Property Let procedure not defined and property
        'Get procedure did not return an object"
        Else
            For i = LBound(tbl.aryColRngs) To UBound(tbl.aryColRngs)
                Debug.Print tbl.aryColRngs(i).Address
            Next i
        End If
        .Update tst, procs
    End With

End Sub
'-----------------------------------------------------------------------------
'-----------------------------------------------------------------------------
'This section tests SetDimensions procedure
'-----------------------------------------------------------------------------
'-----------------------------------------------------------------------------
' Full SetDimensions Procedure
' Blank sheet and header-only sheet
' JDL 5/24/24; Modified 10/21/24
'
Sub test_SetDimensions1(procs)
    Dim tst As New Test: tst.Init tst, "test_SetDimensions1"
    Dim tbl As Object: Set tbl = ExcelSteps.New_tbl
    
    With tst
        InitializeHomedTbl tbl, .wkbkTest

        'Clear sheet contents
        .wkbkTest.Sheets(tbl.sht).Cells.ClearContents
        .Assert tst, tbl.SetDimensions(tbl)

        'Check results of method
        .Assert tst, tbl.rngTable Is Nothing
        .Assert tst, tbl.rngHeader Is Nothing
        .Assert tst, tbl.rngrows Is Nothing
    End With
    
    'tst with header-only sheet
    Set tbl = ExcelSteps.New_tbl
    With tst
        InitializeHomedTbl tbl, .wkbkTest
        PopulateTbl .wkbkTest, shtTbl

        'Clear data rows
        Range(.wkbkTest.Sheets(shtTbl).Rows(2), .wkbkTest.Sheets(shtTbl).Rows(4)).ClearContents
        .Assert tst, tbl.SetDimensions(tbl)

        'Check results of method
        .Assert tst, tbl.rngTable.Address = "$A$1:$C$1"
        .Assert tst, tbl.rngHeader.Address = "$A$1:$C$1"
        .Assert tst, tbl.rngrows Is Nothing

        .Update tst, procs
    End With

End Sub
'-----------------------------------------------------------------------------
' Full SetDimensions Procedure
' Populated table
' JDL 5/24/24; Modified 10/21/24
'
Sub test_SetDimensions2(procs)
    Dim tst As New Test: tst.Init tst, "test_SetDimensions2"
    Dim tbl As Object: Set tbl = ExcelSteps.New_tbl
    
    With tst
        InitializeHomedTbl tbl, .wkbkTest
        PopulateTbl .wkbkTest, shtTbl

        .Assert tst, tbl.SetDimensions(tbl)

        'Check results of method
        .Assert tst, tbl.rngTable.Address = "$A$1:$C$4"
        .Assert tst, tbl.rngHeader.Address = "$A$1:$C$1"
        .Assert tst, tbl.rngrows.Address = "$2:$4"

        .Update tst, procs
    End With
End Sub
'-----------------------------------------------------------------------------
' Set final .rngTable, .rngHeader and .rngRows (useful for clearing or formatting entire table)
' Populated table with header offset -2 from data
' JDL 5/24/24; Modified 10/21/24
'
Sub test_SetDimensions3(procs)
    Dim tst As New Test: tst.Init tst, "test_SetDimensions3"
    Dim tbl As Object: Set tbl = ExcelSteps.New_tbl
    
    'Populate a 3-row table - offset the data an extra row from header
    With tst
        InitializeHomedTbl tbl, .wkbkTest
        
        PopulateTbl .wkbkTest, shtTbl
        .wkbkTest.Sheets(shtTbl).Rows(2).Insert
        tbl.rowHome = 3
        tbl.iOffsetHeader = -2

        .Assert tst, tbl.SetDimensions(tbl)

        'Check results of method
        .Assert tst, tbl.rngTable.Address = "$A$1:$C$5"
        .Assert tst, tbl.rngHeader.Address = "$A$1:$C$1"
        .Assert tst, tbl.rngrows.Address = "$3:$5"

        'Populate a 3-row table - offset the data one column (Blank col A)
        Set tbl = ExcelSteps.New_tbl
        
        InitializeHomedTbl tbl, .wkbkTest
        
        PopulateTbl .wkbkTest, shtTbl
        .wkbkTest.Sheets(shtTbl).Rows(2).Insert
        .wkbkTest.Sheets(shtTbl).Columns(1).Insert
        tbl.rowHome = 3
        tbl.colHome = 2
        tbl.iOffsetHeader = -2

        .Assert tst, tbl.SetDimensions(tbl)

        'Check results of method
        .Assert tst, tbl.rngTable.Address = "$B$1:$D$5"
        .Assert tst, tbl.rngHeader.Address = "$B$1:$D$1"
        .Assert tst, tbl.rngrows.Address = "$3:$5"
    
        .Update tst, procs
    End With
End Sub
'-----------------------------------------------------------------------------
' Set .wksht, .cellHome and .rngTable
' Example with blank sheet
' JDL 5/23/24; Modified 10/21/24
'
Sub test_InitializeWkshtAndRanges1(procs)
    Dim tst As New Test: tst.Init tst, "test_InitializeWkshtAndRanges1"
    Dim tbl As Object: Set tbl = ExcelSteps.New_tbl
    
    With tst
        InitializeHomedTbl tbl, tst.wkbkTest
        .wkbkTest.Sheets(shtTbl).Cells.ClearContents

        'Run method to initialize .wksht, .cellHome and .rngTbl
        .Assert tst, tbl.SetWkshtAndRanges(tbl)
        .Assert tst, tbl.wksht.Name = tbl.sht
        .Assert tst, tbl.cellHome.Address = "$A$2"
        .Assert tst, tbl.cellHome.Parent.Name = tbl.sht
        .Assert tst, tbl.rngTable.Address = "$A$2"
    
        .Update tst, procs
    End With
End Sub
'-----------------------------------------------------------------------------
' Set .wksht, .cellHome and .rngTable
' Example with populated table
' JDL 5/23/24; Modified 10/21/24
'
Sub test_InitializeWkshtAndRanges2(procs)
    Dim tst As New Test: tst.Init tst, "test_InitializeWkshtAndRanges2"
    Dim tbl As Object: Set tbl = ExcelSteps.New_tbl
    
    With tst
        InitializeHomedTbl tbl, .wkbkTest
        PopulateTbl .wkbkTest, shtTbl

        'Run method to initialize .wksht, .cellHome and .rngTbl
        .Assert tst, tbl.SetWkshtAndRanges(tbl)
        .Assert tst, tbl.wksht.Name = tbl.sht
        .Assert tst, tbl.cellHome.Address = "$A$2"
        .Assert tst, tbl.cellHome.Parent.Name = tbl.sht
        .Assert tst, tbl.rngTable.Address = "$A$1:$C$4"
    
        .Update tst, procs
    End With
End Sub
'-----------------------------------------------------------------------------
' Set the .IsBlankSht attribute (True if .sht is blank)
' with blank sheet
' JDL 5/23/24; Modified 10/21/24
'
Sub test_SetIsBlankSheet1(procs)
    Dim tst As New Test: tst.Init tst, "test_SetIsBlankSheet1"
    Dim tbl As Object: Set tbl = ExcelSteps.New_tbl
    
    With tst
        InitializeHomedTbl tbl, .wkbkTest
        PopulateTbl .wkbkTest, shtTbl
        
        'Clear sheet contents and initialize .wksht, .cellHome and .rngTbl
        .wkbkTest.Sheets(shtTbl).Cells.ClearContents
        .Assert tst, tbl.SetWkshtAndRanges(tbl)
        .Assert tst, tbl.SetIsBlankSheet(tbl)
        .Assert tst, tbl.IsBlankSht = True

        .Update tst, procs
    End With

End Sub
'-----------------------------------------------------------------------------
' Set the .IsBlankSht attribute (True if .sht is blank)
' with populated table
' JDL 5/23/24; Modified 10/21/24
'
Sub test_SetIsBlankSheet2(procs)
    Dim tst As New Test: tst.Init tst, "test_SetIsBlankSheet2"
    Dim tbl As Object: Set tbl = ExcelSteps.New_tbl
    
    With tst
        InitializeHomedTbl tbl, .wkbkTest

        'Populate table onto sht and initialize .wksht, .cellHome and .rngTbl
        PopulateTbl .wkbkTest, shtTbl
        .Assert tst, tbl.SetWkshtAndRanges(tbl)
        .Assert tst, tbl.SetIsBlankSheet(tbl)
        .Assert tst, tbl.IsBlankSht = False

        .Update tst, procs
    End With
End Sub
'-----------------------------------------------------------------------------
' Set the .IsNoData attribute to flag case where only header is populated
' with blank sheet
' JDL 5/23/24; Modified 10/21/24
'
Sub test_SetIsNoData1(procs)
    Dim tst As New Test: tst.Init tst, "test_SetIsNoData1"
    Dim tbl As Object: Set tbl = ExcelSteps.New_tbl
    
    With tst
        InitializeHomedTbl tbl, .wkbkTest

        'Clear sheet contents and call methods
        .wkbkTest.Sheets(shtTbl).Cells.ClearContents
        InitializeForSetDimensions tst, tbl
        .Assert tst, tbl.IsNoData = True

        .Update tst, procs
    End With
End Sub
'-----------------------------------------------------------------------------
' Set the .IsNoData attribute to flag case where only header is populated
' with populated table
' JDL 5/23/24; Modified 10/21/24
'
Sub test_SetIsNoData2(procs)
    Dim tst As New Test: tst.Init tst, "test_SetIsNoData2"
    Dim tbl As Object: Set tbl = ExcelSteps.New_tbl
    
    With tst
        InitializeHomedTbl tbl, .wkbkTest
        
        'Populate a 3-row table and clear the data rows
        PopulateTbl .wkbkTest, shtTbl
        Range(.wkbkTest.Sheets(shtTbl).Rows(2), .wkbkTest.Sheets(shtTbl).Rows(4)).ClearContents

        'call methods
        InitializeForSetDimensions tst, tbl
        .Assert tst, tbl.IsNoData = True

        .Update tst, procs
    End With
End Sub
'-----------------------------------------------------------------------------
' Set the .IsNoData attribute to flag case where only header is populated
' with populated table
' JDL 5/23/24; Modified 10/21/24
'
Sub test_SetIsNoData3(procs)
    Dim tst As New Test: tst.Init tst, "test_SetIsNoData3"
    Dim tbl As Object: Set tbl = ExcelSteps.New_tbl
    
    With tst
        InitializeHomedTbl tbl, .wkbkTest
        
        'Populate a 3-row table
        PopulateTbl .wkbkTest, shtTbl

        'call methods
        InitializeForSetDimensions tst, tbl
        .Assert tst, tbl.IsNoData = False
        .Assert tst, tbl.rngTable.Address = "$A$1:$C$4"

        .Update tst, procs
    End With
End Sub
'-----------------------------------------------------------------------------
' Set the .IsNoData attribute to flag case where only header is populated
' with populated table - data in two cells of first data row
' JDL 5/23/24; Modified 10/21/24
'
Sub test_SetIsNoData4(procs)
    Dim tst As New Test: tst.Init tst, "test_SetIsNoData4"
    Dim tbl As Object: Set tbl = ExcelSteps.New_tbl
    
    With tst
        InitializeHomedTbl tbl, .wkbkTest
        
        'Populate a 3-row table and clear all but two cells (B2 and C2)
        PopulateTbl .wkbkTest, shtTbl
        Range(.wkbkTest.Sheets(shtTbl).Rows(3), .wkbkTest.Sheets(shtTbl).Rows(4)).ClearContents
        .wkbkTest.Sheets(shtTbl).Cells(2, 1).ClearContents

        InitializeForSetDimensions tst, tbl
        .Assert tst, tbl.IsNoData = False
        .Assert tst, tbl.rngTable.Address = "$A$1:$C$2"

        .Update tst, procs
    End With
End Sub
'-----------------------------------------------------------------------------
' If not already initialized, set .nrows; set .lastrow and .rngRows
' Blank sheet
' JDL 5/23/24; Modified 10/21/24
'
Sub test_SetNRows1(procs)
    Dim tst As New Test: tst.Init tst, "test_SetNRows1"
    Dim tbl As Object: Set tbl = ExcelSteps.New_tbl
    
    With tst
        InitializeHomedTbl tbl, .wkbkTest

        'Clear sheet contents
        .wkbkTest.Sheets(shtTbl).Cells.ClearContents
        InitializeForSetDimensions tst, tbl
        .Assert tst, tbl.SetNRows(tbl)

        'Check results of method
        .Assert tst, tbl.nRows = 0
        .Assert tst, tbl.rngrows Is Nothing
        .Assert tst, tbl.lastrow = 0
    
        'tst with header-only sheet
        Set tbl = ExcelSteps.New_tbl
        
        InitializeHomedTbl tbl, tst.wkbkTest

        'Clear data rows
        Range(.wkbkTest.Sheets(shtTbl).Rows(2), .wkbkTest.Sheets(shtTbl).Rows(4)).ClearContents
        InitializeForSetDimensions tst, tbl
        .Assert tst, tbl.SetNRows(tbl)

        .Assert tst, tbl.nRows = 0
        .Assert tst, tbl.rngrows Is Nothing
        .Assert tst, tbl.lastrow = 0

        .Update tst, procs
    End With
End Sub
'-----------------------------------------------------------------------------
' If not already initialized, set .nrows; set .lastrow and .rngRows
' with header-only sheet
' JDL 5/23/24; Modified 10/21/24
'
Sub test_SetNRows2(procs)
    Dim tst As New Test: tst.Init tst, "test_SetNRows2"
    Dim tbl As Object: Set tbl = ExcelSteps.New_tbl
    
    With tst
        InitializeHomedTbl tbl, .wkbkTest
        
        PopulateTbl .wkbkTest, shtTbl

        'Clear sheet contents and initialize .wksht, .cellHome and .rngTbl
        InitializeForSetDimensions tst, tbl
        .Assert tst, tbl.SetNRows(tbl)

        .Assert tst, tbl.nRows = 3
        .Assert tst, tbl.rngrows.Address = "$2:$4"
        .Assert tst, tbl.lastrow = 4

        .Update tst, procs
    End With
End Sub
'-----------------------------------------------------------------------------
' If not already initialized, set .nrows; set .lastrow and .rngRows
' with populated data and header offset -2 from cellHome
' JDL 5/23/24; Modified 10/21/24
'
Sub test_SetNRows3(procs)
    Dim tst As New Test: tst.Init tst, "test_SetNRows3"
    Dim tbl As Object: Set tbl = ExcelSteps.New_tbl
    
    With tst
        InitializeHomedTbl tbl, tst.wkbkTest
        
        'Populate a 3-row table - offset the data an extra row from header
        PopulateTbl .wkbkTest, shtTbl
        .wkbkTest.Sheets(shtTbl).Rows(2).Insert
        tbl.rowHome = 3
        tbl.iOffsetHeader = -2

        'Clear sheet contents and initialize .wksht, .cellHome and .rngTbl
        InitializeForSetDimensions tst, tbl
        .Assert tst, tbl.SetNRows(tbl)

        'Check results of method
        .Assert tst, tbl.nRows = 3
        .Assert tst, tbl.rngrows.Address = "$3:$5"
        .Assert tst, tbl.lastrow = 5

        .Update tst, procs
    End With
End Sub
'-----------------------------------------------------------------------------
' If not already initialized, set .ncols; set .lastcol, and .rngheader
' Blank sheet and header-only sheet
' JDL 5/23/24; Modified 10/21/24
'
Sub test_SetNCols1(procs)
    Dim tst As New Test: tst.Init tst, "test_SetNCols1"
    Dim tbl As Object: Set tbl = ExcelSteps.New_tbl
    
    With tst
        InitializeHomedTbl tbl, .wkbkTest

        'Clear sheet contents
        .wkbkTest.Sheets(shtTbl).Cells.ClearContents
        InitializeForSetDimensions tst, tbl
        .Assert tst, tbl.SetNCols(tbl)

        'Check results of method
        .Assert tst, tbl.nCols = 0
        .Assert tst, tbl.rngHeader Is Nothing
        .Assert tst, tbl.lastcol = 0
    
        'tst with header-only sheet
        Set tbl = ExcelSteps.New_tbl
        
        InitializeHomedTbl tbl, tst.wkbkTest
        PopulateTbl .wkbkTest, shtTbl

        'Clear data rows
        Range(.wkbkTest.Sheets(shtTbl).Rows(2), .wkbkTest.Sheets(shtTbl).Rows(4)).ClearContents
        InitializeForSetDimensions tst, tbl
        .Assert tst, tbl.SetNCols(tbl)

        'Check results of method
        .Assert tst, tbl.nCols = 3
        .Assert tst, tbl.rngHeader.Address = "$A$1:$C$1"
        .Assert tst, tbl.lastcol = 3

        .Update tst, procs
    End With
End Sub
'-----------------------------------------------------------------------------
' If not already initialized, set .ncols; set .lastcol, and .rngheader
' Populated table
' JDL 5/23/24; Modified 10/21/24
'
Sub test_SetNCols2(procs)
    Dim tst As New Test: tst.Init tst, "test_SetNCols2"
    Dim tbl As Object: Set tbl = ExcelSteps.New_tbl
    
    With tst
        InitializeHomedTbl tbl, tst.wkbkTest
        PopulateTbl .wkbkTest, shtTbl

        InitializeForSetDimensions tst, tbl
        .Assert tst, tbl.SetNCols(tbl)

        'Check results of method
        .Assert tst, tbl.nCols = 3
        .Assert tst, tbl.rngHeader.Address = "$A$1:$C$1"
        .Assert tst, tbl.lastcol = 3

        .Update tst, procs
    End With
End Sub
'-----------------------------------------------------------------------------
' If not already initialized, set .ncols; set .lastcol, and .rngheader
' Populated table with header offset -2 from data
' JDL 5/23/24; Modified 10/21/24
'
Sub test_SetNCols3(procs)
    Dim tst As New Test: tst.Init tst, "test_SetNCols3"
    Dim tbl As Object: Set tbl = ExcelSteps.New_tbl
    
    With tst
        InitializeHomedTbl tbl, .wkbkTest
        
        PopulateTbl .wkbkTest, shtTbl
        .wkbkTest.Sheets(shtTbl).Rows(2).Insert
        tbl.rowHome = 3
        tbl.iOffsetHeader = -2

        InitializeForSetDimensions tst, tbl
        .Assert tst, tbl.SetNCols(tbl)

        'Check results of method
        .Assert tst, tbl.nCols = 3
        .Assert tst, tbl.rngHeader.Address = "$A$1:$C$1"
        .Assert tst, tbl.lastcol = 3

        'Populate a 3-row table - offset the data one column (Blank col A)
        Set tbl = ExcelSteps.New_tbl
        
        InitializeHomedTbl tbl, .wkbkTest
        
        PopulateTbl .wkbkTest, shtTbl
        .wkbkTest.Sheets(shtTbl).Rows(2).Insert
        .wkbkTest.Sheets(shtTbl).Columns(1).Insert
        tbl.rowHome = 3
        tbl.colHome = 2
        tbl.iOffsetHeader = -2

        InitializeForSetDimensions tst, tbl
        .Assert tst, tbl.SetNRows(tbl)
        .Assert tst, tbl.SetNCols(tbl)

        'Check results of method
        .Assert tst, tbl.nCols = 3
        .Assert tst, tbl.rngHeader.Address = "$B$1:$D$1"
        .Assert tst, tbl.lastcol = 4
        .Assert tst, tbl.lastrow = 5
        .Assert tst, tbl.rngrows.Address = "$3:$5"

        .Update tst, procs
    End With
End Sub
'-----------------------------------------------------------------------------
' If not already initialized, set .ncols; set .lastcol, and .rngheader
' preset nrows and ncols different than data extent
' JDL 5/23/24
'
Sub test_SetNCols4(procs)
    Dim tst As New Test
    tst.Init tst, "test_SetNCols4"
    
    'Test-specific code
    '------------------
    Dim tbl As Object
    Set tbl = ExcelSteps.New_tbl
    With tbl
        InitializeHomedTbl tbl, tst.wkbkTest
        PopulateTbl tst.wkbkTest, shtTbl
        .nRows = 10
        .nCols = 4

        InitializeForSetDimensions tst, tbl
        tst.Assert tst, .SetNRows(tbl)
        tst.Assert tst, .SetNCols(tbl)

        'Check results of method
        tst.Assert tst, .nCols = 4
        tst.Assert tst, .nRows = 10
        tst.Assert tst, .rngHeader.Address = "$A$1:$D$1"
        tst.Assert tst, .lastcol = 4
        tst.Assert tst, .rngrows.Address = "$2:$11"
    End With

    tst.Update tst, procs
End Sub
'-----------------------------------------------------------------------------
' Set final .rngTable (useful for clearing or formatting entire table)
' Blank sheet and header-only sheet
' JDL 5/24/24
'
Sub test_SetRngTable1(procs)
    Dim tst As New Test
    tst.Init tst, "test_SetRngTable1"
    
    'Test-specific code
    '------------------
    Dim tbl As Object
    
    'tst with blank sheet
    Set tbl = ExcelSteps.New_tbl
    With tbl
        InitializeHomedTbl tbl, tst.wkbkTest

        'Clear sheet contents
        .wkbk.Sheets(.sht).Cells.ClearContents
        InitializeForSetDimensions tst, tbl
        tst.Assert tst, .SetNRows(tbl)
        tst.Assert tst, .SetNCols(tbl)
        tst.Assert tst, .SetRngTable(tbl)

        'Check results of method
        tst.Assert tst, .rngTable Is Nothing
    End With
    
    'tst with header-only sheet
    Set tbl = ExcelSteps.New_tbl
    With tbl
        InitializeHomedTbl tbl, tst.wkbkTest
        PopulateTbl tst.wkbkTest, shtTbl

        'Clear data rows but leave header populated
        Range(.wkbk.Sheets(.sht).Rows(2), .wkbk.Sheets(.sht).Rows(4)).ClearContents
        InitializeForSetDimensions tst, tbl
        tst.Assert tst, .SetNRows(tbl)
        tst.Assert tst, .SetNCols(tbl)
        tst.Assert tst, .SetRngTable(tbl)

        'Check results of method
        tst.Assert tst, .rngTable.Address = "$A$1:$C$1"
    End With

    tst.Update tst, procs
End Sub
'-----------------------------------------------------------------------------
' Set final .rngTable (useful for clearing or formatting entire table)
' Populated table
' JDL 5/24/24
'
Sub test_SetRngTable2(procs)
    Dim tst As New Test
    tst.Init tst, "test_SetRngTable2"
    
    'Test-specific code
    '------------------
    Dim tbl As Object
    Set tbl = ExcelSteps.New_tbl
    With tbl
        InitializeHomedTbl tbl, tst.wkbkTest
        PopulateTbl tst.wkbkTest, shtTbl

        InitializeForSetDimensions tst, tbl
        tst.Assert tst, .SetNRows(tbl)
        tst.Assert tst, .SetNCols(tbl)
        tst.Assert tst, .SetRngTable(tbl)

        'Check results of method
        tst.Assert tst, .rngTable.Address = "$A$1:$C$4"
    End With

    tst.Update tst, procs
End Sub
'-----------------------------------------------------------------------------
' Set final .rngTable (useful for clearing or formatting entire table)
' Populated table with header offset -2 from data
' JDL 5/24/24
'
Sub test_SetRngTable3(procs)
    Dim tst As New Test
    tst.Init tst, "test_SetRngTable3"
    
    'Test-specific code
    '------------------
    Dim tbl As Object
    
    'Populate a 3-row table - offset the data an extra row from header
    Set tbl = ExcelSteps.New_tbl
    With tbl
        InitializeHomedTbl tbl, tst.wkbkTest
        
        PopulateTbl tst.wkbkTest, shtTbl
        .wkbk.Sheets(.sht).Rows(2).Insert
        .rowHome = 3
        .iOffsetHeader = -2

        InitializeForSetDimensions tst, tbl
        tst.Assert tst, .SetNRows(tbl)
        tst.Assert tst, .SetNCols(tbl)
        tst.Assert tst, .SetRngTable(tbl)

        'Check results of method
        tst.Assert tst, .rngTable.Address = "$A$1:$C$5"
    End With

    'Populate a 3-row table - offset the data one column (Blank col A)
    Set tbl = ExcelSteps.New_tbl
    With tbl
        InitializeHomedTbl tbl, tst.wkbkTest
        
        PopulateTbl tst.wkbkTest, shtTbl
        .wkbk.Sheets(.sht).Rows(2).Insert
        .wkbk.Sheets(.sht).Columns(1).Insert
        .rowHome = 3
        .colHome = 2
        .iOffsetHeader = -2

        InitializeForSetDimensions tst, tbl
        tst.Assert tst, .SetNRows(tbl)
        tst.Assert tst, .SetNCols(tbl)
        tst.Assert tst, .SetRngTable(tbl)

        'Check results of method
        tst.Assert tst, .rngTable.Address = "$B$1:$D$5"
    End With
    
    tst.Update tst, procs
End Sub
'-----------------------------------------------------------------------------
' Helper sub to run precursor methods to SetNRows and SetNCols
' JDL 5/23/24
'
Sub InitializeForSetDimensions(tst, tbl)
    With tbl
        tst.Assert tst, .SetWkshtAndRanges(tbl)
        tst.Assert tst, .SetIsBlankSheet(tbl)
        tst.Assert tst, .SetIsNoData(tbl)
    End With
End Sub
'-----------------------------------------------------------------------------------------------
' Helper sub to initialize wkbk, home row/col and sht
' JDL 5/23/24
'
Sub InitializeHomedTbl(tbl, wkbk)
    With tbl
        Set .wkbk = wkbk
        .rowHome = 2
        .colHome = 1
        .sht = "SMdl"
        .iOffsetHeader = -1
    End With
End Sub
'-----------------------------------------------------------------------------------------------
'-----------------------------------------------------------------------------------------------
'Testing for tblRowsCols.tblInit() procedure
'-----------------------------------------------------------------------------------------------
'-----------------------------------------------------------------------------------------------
' tblInit Procedure
' Default/homed table
' JDL 5/24/24
'
Sub test_tblInitProcedure1(procs)
    Dim tst As New Test
    tst.Init tst, "test_tblInitProcedure1"

    'tst-specific code
    '------------------
    Dim tbl As Object
    Set tbl = ExcelSteps.New_tbl
    
    'A default table (only sht is specified)
    PopulateTbl2 procs.wkbk_testing, shtTbl
    With tbl
        tst.Assert tst, .Init(tbl, procs.wkbk_testing, sht:=shtTbl)
        tst.Assert tst, .rowHome = 2
        tst.Assert tst, .colHome = 1
    End With
    '------------------
    
    'Update tst results on procs.wksht_results
    tst.Update tst, procs
End Sub
'-----------------------------------------------------------------------------
'tst that tblRowsCols.Init warns and corrects sht argument if doesn't
'match case-sensitive Sheet name
'
'JDL 1/20/22; 3/6/23 revise to point to ExcelSteps; 5/22/24 refactoring
'to test this with tblRowsCols.Init function
'Updated 10/29/25 to add workbook argument to errs.Init call
'
Sub test_tblInitProcedure_sht_CaseError(procs)
    Dim tst As New Test
    tst.Init tst, "test_tblInitProcedure_sht_CaseError"
    
    'Test-specific code
    '------------------
    Dim tbl As Object, i As Integer, msg As String
    Set tbl = ExcelSteps.New_tbl
        
    'Initialize error handling to allow checking warning and error message(s)
    Set ExcelSteps.errs = ExcelSteps.New_ErrorHandling
    ExcelSteps.errs.Init wkbkE:=ExcelSteps.ThisWorkbook
    ExcelSteps.errs.IsShowMsgs = False
    
    PopulateTbl tst.wkbkTest, shtTbl
        
    'Intentionally provision with lower case (e.g. wrong) sheet name
    tst.Assert tst, tbl.Init(tbl, tst.wkbkTest, sht:=LCase(shtTbl))
    
    'Check that warning message was created (not shown if .IsShowMsgs=False)
    msg = ExcelSteps.errs.Msgs_accum
    tst.Assert tst, Left(msg, 87) = "A tblRowsCols Definition's sht argument " & _
        "does not match the case of the workbook's sheet"
    
    'Check that the sht argument was corrected to exactly match the Sheet name
    tst.Assert tst, tbl.sht = shtTbl
    
    tst.Update tst, procs
End Sub
'-----------------------------------------------------------------------------------------------
' Set flag for custom or default/homed table configuration
' JDL 5/21/24; Modified 11/6/24 add case where sht and TblName specified
Sub test_SetIsCustomTbl(procs)
    Dim tst As New Test
    tst.Init tst, "test_SetIsCustomTbl"
    
    'Test-specific code
    '------------------
    Dim tbl As Object
    
    'A default table (only sht is specified)
    Set tbl = ExcelSteps.New_tbl
    With tbl
        tst.Assert tst, tbl.SetIsCustomTbl(tbl, sht:="SMdl")
        tst.Assert tst, (Not .IsCustomTbl)
        tst.Assert tst, .sht = "SMdl"
        tst.Assert tst, .TblName = "SMdl"
    End With
    
    'A default table (only sht and override TblName are specified)
    Set tbl = ExcelSteps.New_tbl
    With tbl
        tst.Assert tst, tbl.SetIsCustomTbl(tbl, sht:="SMdl", TblName:="AltTableName")
        tst.Assert tst, (Not .IsCustomTbl)
        tst.Assert tst, .sht = "SMdl"
        tst.Assert tst, .TblName = "AltTableName"
    End With
    
    'A custom table (TblName specified)
    Set tbl = ExcelSteps.New_tbl
    With tbl
        tst.Assert tst, .SetIsCustomTbl(tbl, TblName:="test_table")
        tst.Assert tst, .IsCustomTbl
        tst.Assert tst, .TblName = "test_table"
    End With

    'A custom table (Defn specified)
    Set tbl = ExcelSteps.New_tbl
    With tbl
        tst.Assert tst, .SetIsCustomTbl(tbl, defn:=defn_test)
        tst.Assert tst, .IsCustomTbl
    End With

    'A custom table (Both specified)
    Set tbl = ExcelSteps.New_tbl
    With tbl
        tst.Assert tst, .SetIsCustomTbl(tbl, TblName:="test_table", defn:=defn_test)
        tst.Assert tst, .IsCustomTbl
    End With

    'A custom table (Defn specified but sht arg overrides default sheet name)
    Set tbl = ExcelSteps.New_tbl
    With tbl
        tst.Assert tst, .SetIsCustomTbl(tbl, defn:="Dummy:xxx", sht:="SMdl")
        tst.Assert tst, .IsCustomTbl
        
        'If defn specified, .sht gets set in SetCustomTblParams
        tst.Assert tst, .sht = "SMdl"
        tst.Assert tst, .defn = "Dummy:xxx"
    End With
    tst.Update tst, procs
End Sub
'-----------------------------------------------------------------------------------------------
' Populate custom table parameters from TableName (Setting) or Defn argument
' (From Setting)
' JDL 5/21/24
'
Sub test_PopulateCustomTblParams1(procs)
    Dim tst As New Test
    tst.Init tst, "test_PopulateCustomTblParams1"
    
    'Test-specific code
    '------------------
    Dim setting_name As String, tbl As Object
    Set tbl = ExcelSteps.New_tbl
    
    'Write the definition to the workbook as a setting and check that it exists
    setting_name = "tbl_test_table"
    ExcelSteps.UpdateSetting tst.wkbkTest, setting_name, defn_test
    
    'Set .IsCustomTbl=True and set .TblName
    With tbl
        Set .wkbk = tst.wkbkTest
        tst.Assert tst, .SetIsCustomTbl(tbl, TblName:="test_table")
        
        'Read Defn from a setting and parse into tbl attributes
        tst.Assert tst, .SetCustomTblParams(tbl)
        tst.Assert tst, .defn = defn_test
        
        'Check individual attributes based on defn_test string
        CheckParsedDefnParams tst, tbl
        tst.Assert tst, .sht = "SMdl"
    End With
    
    'Delete Setting after checks
    ExcelSteps.DeleteSetting tst.wkbkTest, setting_name
    
    tst.Update tst, procs
End Sub
'-----------------------------------------------------------------------------------------------
' Populate custom table parameters from TableName (Setting) or Defn argument
' (From specified Defn string)
' JDL 5/21/24
'
Sub test_PopulateCustomTblParams2(procs)
    Dim tst As New Test
    tst.Init tst, "test_PopulateCustomTblParams2"
    
    Dim tbl As Object
    Set tbl = ExcelSteps.New_tbl
    
    'Ensure no setting that could be confused as having priority over defn
    ExcelSteps.DeleteSetting tst.wkbkTest, "tbl_test_table"
    
    'Set .IsCustomTbl=True and set .TblName
    With tbl
        Set .wkbk = tst.wkbkTest
        .defn = defn_test
        
        'Read Defn from a setting and parse into tbl attributes
        tst.Assert tst, .SetCustomTblParams(tbl)
        
        'Check individual attributes based on defn_test string
        CheckParsedDefnParams tst, tbl
        tst.Assert tst, .sht = "SMdl"
    End With
        
    tst.Update tst, procs
End Sub
'-----------------------------------------------------------------------------------------------
' Populate custom table parameters from TableName (Setting) or Defn argument
' (From specified Defn string with override sheet name argument)
' JDL 5/21/24
'
Sub test_PopulateCustomTblParams3(procs)
    Dim tst As New Test
    tst.Init tst, "test_PopulateCustomTblParams3"
    
    Dim tbl As Object
    Set tbl = ExcelSteps.New_tbl
    
    'Ensure no setting that could be confused as having priority over defn
    ExcelSteps.DeleteSetting tst.wkbkTest, "tbl_test_table"
    
    'Set .IsCustomTbl=True and set .TblName
    With tbl
        Set .wkbk = tst.wkbkTest
        .defn = defn_test
        .sht = "Alt_SheetName"
        
        'Read Defn from a setting and parse into tbl attributes
        tst.Assert tst, .SetCustomTblParams(tbl)
        
        'Check individual attributes based on defn_test string
        CheckParsedDefnParams tst, tbl
        tst.Assert tst, .sht = "Alt_SheetName"
    End With
    
    tst.Update tst, procs
End Sub
'-----------------------------------------------------------------------------------------------
' Helper sub to check individual attributes based on defn_test string
' JDL 5/21/24
'
Sub CheckParsedDefnParams(tst, tbl)
    With tbl
        tst.Assert tst, .rowHome = 6
        tst.Assert tst, .colHome = 2
        tst.Assert tst, .IsSetTblNames = True
        tst.Assert tst, .IsSetColNames = False
        tst.Assert tst, .IsNamePrefix = True
        tst.Assert tst, .IsPrefixSht = False
        tst.Assert tst, .IsSetAryCols = True
        tst.Assert tst, .IsSetColRngs = False
        tst.Assert tst, .iOffsetKeyCol = 1
        tst.Assert tst, .iOffsetHeader = -2
        tst.Assert tst, .nRows = 12
        tst.Assert tst, .nCols = 5
    End With
End Sub
'-----------------------------------------------------------------------------------------------
' Read custom table definition from Settings
' JDL 5/21/24
'
Sub test_ReadDefnSetting(procs)
    Dim tst As New Test
    tst.Init tst, "test_ReadDefnSetting"
    
    'Test-specific code
    '------------------
    Dim setting_name As String, tbl As Object
    Set tbl = ExcelSteps.New_tbl
    
    'Write the definition to the workbook as a setting and check that it exists
    setting_name = "tbl_test_table"
    ExcelSteps.UpdateSetting tst.wkbkTest, setting_name, defn_test
    
    'Set .IsCustomTbl=True and set .TblName
    Set tbl.wkbk = tst.wkbkTest
    tst.Assert tst, tbl.SetIsCustomTbl(tbl, TblName:="test_table")
    
    'Read the definition from Settings sheet
    tst.Assert tst, tbl.ReadDefnSetting(tbl)
    tst.Assert tst, tbl.defn = defn_test
    
    'Delete after check
    ExcelSteps.DeleteSetting tst.wkbkTest, setting_name

    'Repeat with
    tst.Update tst, procs
End Sub
'-----------------------------------------------------------------------------------------------
' Assign default attribute values for homed table
' JDL 5/22/24
'
'
Sub test_SetHomedTblParams(procs)
    Dim tst As New Test
    tst.Init tst, "test_SetHomedTblParams"
    
    'Test-specific code
    '------------------
    Dim tbl As Object
    Set tbl = ExcelSteps.New_tbl
    
    With tbl
        Set .wkbk = tst.wkbkTest
        .sht = "SMdl"
        tst.Assert tst, .SetHomedTblParams(tbl)
        
        'Check assigned values
        tst.Assert tst, .rowHome = 2
        tst.Assert tst, .colHome = 1
        tst.Assert tst, .sht = "SMdl"
        tst.Assert tst, .IsSetTblNames = False
        tst.Assert tst, .IsSetColNames = True
        tst.Assert tst, .IsNamePrefix = False
        tst.Assert tst, .IsPrefixSht = False
        tst.Assert tst, .IsSetAryCols = False
        tst.Assert tst, .IsSetColRngs = False
        tst.Assert tst, .iOffsetKeyCol = 0
        tst.Assert tst, .iOffsetHeader = -1
    End With

    tst.Update tst, procs
End Sub
'-----------------------------------------------------------------------------------------------
' Override attribute values if non-default tblInit arguments are specified
' JDL 5/22/24
'
Sub test_OverrideWithArgs(procs)
    Dim tst As New Test
    tst.Init tst, "test_OverrideWithArgs"
    
    'Test-specific code
    '------------------
    Dim tbl As Object
    Set tbl = ExcelSteps.New_tbl
    
    With tbl
        Set .wkbk = tst.wkbkTest
        .sht = "SMdl"
        .IsCustomTbl = True
        
        'Assign alternate values to test overriding
        Dim rcHome_alt As String, sht_alt As String, IsSetTblNames_alt As Boolean
        Dim IsSetColNames_alt As Boolean, IsNamePrefix_alt As Boolean
        Dim IsPrefixSht_alt As Boolean, IsSetAryCols_alt As Boolean
        Dim IsSetColRngs_alt As Boolean, iOffsetKeyCol_alt As Integer
        Dim iOffsetHeader_alt As Integer, nrows_alt As Integer, ncols_alt As Integer
        Dim NamePrefix_alt As String
    
        sht_alt = "Descriptions"

        rcHome_alt = "4,4"
        IsSetTblNames_alt = True
        IsSetColNames_alt = True
        IsNamePrefix_alt = True
        IsPrefixSht_alt = True
        IsSetAryCols_alt = True
        IsSetColRngs_alt = True
        iOffsetKeyCol_alt = 2
        iOffsetHeader_alt = -2
        nrows_alt = 20
        ncols_alt = 10
        NamePrefix_alt = "AltTblName"

        tst.Assert tst, .SetHomedTblParams(tbl)
        tst.Assert tst, .OverrideWithArgs(tbl, _
            sht:=sht_alt, _
            rcHome:=rcHome_alt, _
            nRows:=nrows_alt, _
            nCols:=ncols_alt, _
            iOffsetKeyCol:=iOffsetKeyCol_alt, _
            iOffsetHeader:=iOffsetHeader_alt, _
            IsSetAryCols:=IsSetAryCols_alt, _
            IsSetColRngs:=IsSetColRngs_alt, _
            IsSetTblNames:=IsSetTblNames_alt, _
            IsSetColNames:=IsSetColNames_alt, _
            IsNamePrefix:=IsNamePrefix_alt, _
            IsPrefixSht:=IsPrefixSht_alt, _
            NamePrefix:=NamePrefix_alt)
        
        'Check assigned values
        tst.Assert tst, .rowHome = 4
        tst.Assert tst, .colHome = 4
        tst.Assert tst, .sht = "Descriptions"
        tst.Assert tst, .IsSetTblNames = IsSetTblNames_alt
        tst.Assert tst, .IsSetColNames = IsSetColNames_alt
        tst.Assert tst, .IsNamePrefix = IsNamePrefix_alt
        tst.Assert tst, .IsPrefixSht = IsPrefixSht_alt
        tst.Assert tst, .IsSetAryCols = IsSetAryCols_alt
        tst.Assert tst, .IsSetColRngs = IsSetColRngs_alt
        tst.Assert tst, .iOffsetKeyCol = iOffsetKeyCol_alt
        tst.Assert tst, .iOffsetHeader = iOffsetHeader_alt
        tst.Assert tst, .NamePrefix = NamePrefix_alt
    End With

    tst.Update tst, procs
End Sub
'-----------------------------------------------------------------------------------------------
'-----------------------------------------------------------------------------------------------
'Testing for Refresh.RefreshRC() (refresh from ExcelSteps sheet)
'-----------------------------------------------------------------------------------------------
'-----------------------------------------------------------------------------
'Refresh rows/columns table - Two-row ExcelSteps sheet
'
'JDL 5/28/24; 10/18/24 refactored for .RefreshRC args and refr refactor
'
Sub test_RefreshTbl3(procs)
    Dim tst As New Test
    tst.Init tst, "test_RefreshTbl3"
    
    'Test-specific code
    '------------------
    Dim refr As Object, tblSteps As Object
    PopulateTbl2 tst.wkbkTest, shtTbl
    
    'Prep Excel Steps (clear previous and replace)
    PrepBlankStepsForTesting tst.wkbkTest, refr, tblSteps
    PopulateStepsTblRefresh tst.wkbkTest, shtTbl
    
    'Check ExcelSteps values
    With tblSteps
        tst.Assert tst, .rngrows.Address = "$2:$21"
        tst.Assert tst, .wksht.Cells(2, 1) = "SMdl"
    End With
    
    With refr
        'set refr required attributes
        .IsReplace = True
        .IsTblFormat = True
        
        'Refresh the populated table (default refresh)
        tst.Assert tst, .RefreshRC(refr, tblSteps, sht:=shtTbl)
    End With
    
    With tst.wkbkTest.Sheets(shtTbl)
        tst.Assert tst, .Cells(2, 5).NumberFormat = "0.000"
        tst.Assert tst, .Cells(2, 7).NumberFormat = "0.00"
        tst.Assert tst, .Cells(2, 7).Formula = "=@Data_2 + @Data_3"
        tst.Assert tst, .Cells(2, 7).Value = 39.6
    End With
    
    tst.Update tst, procs
End Sub
'-----------------------------------------------------------------------------------------
'Refresh rows/columns table - Blank ExcelSteps sheet
'
'JDL 2/21/22; rewrite 3/6/23; 10/18/24 refactored for .RefreshRC args and refr refactor
'
Sub test_RefreshTbl2(procs)
    Dim tst As New Test
    tst.Init tst, "test_RefreshTbl2"
    
    'Test-specific code
    '------------------
    Dim refr As Object, tblSteps As Object
    PopulateTbl2 tst.wkbkTest, shtTbl
            
    'Prep Excel Steps (clear previous and replace)
    PrepBlankStepsForTesting tst.wkbkTest, refr, tblSteps
    
    With refr
        'set refr required attributes
        .IsReplace = True
        .IsTblFormat = True
        
        'Refresh the populated table (default refresh)
        tst.Assert tst, .RefreshRC(refr, tblSteps, sht:=shtTbl)
    End With
    
    CheckRefreshedTable tst
    
    tst.Update tst, procs
End Sub
Sub CheckRefreshedTable(tst)
    Dim rng As Range
    With tst
        .Assert tst, (tst.wkbkTest.Names("Desc").RefersToRange.Parent.Name = shtTbl)
        .Assert tst, (tst.wkbkTest.Names("Desc").RefersToRange.Address = "$A:$A")
        Set rng = Range(tst.wkbkTest.Sheets(shtTbl).Cells(1, 1), tst.wkbkTest.Sheets(shtTbl).Cells(1, 6))
        .Assert tst, rng.Style = "Accent1"
    End With
End Sub
'-----------------------------------------------------------------------------------------
'Prep Excel Steps sheet
'
'JDL 3/6/23 for Refresh Class in ExcelSteps;
'       Modified 10/18/24 to call .InitTbl; simplify args to use default .wkbk and .wkbkS
'
Sub test_PrepExcelStepsSht(procs)
    Dim tst As New Test
    tst.Init tst, "test_PrepExcelStepsSht"
    
    'Test-specific code
    '------------------
    Dim tbl As Object, refr As Object, tblSteps As Object
    Set refr = ExcelSteps.New_Refresh
    Set tblSteps = ExcelSteps.New_tbl
    
    With refr
        tst.Assert tst, .InitTbl(refr, wkbk:=tst.wkbkTest, IsReplace:=True, IsTblFormat:=True)
        tst.Assert tst, .shtS = shtSteps
        
        'For testing, clear previous ExcelSteps sheet if any
        If SheetExists(.wkbk, shtSteps) Then .wkbk.Sheets(shtSteps).Cells.Clear
        
        'ExcelSteps Prep
        tst.Assert tst, .PrepExcelStepsSht(refr, tblSteps)
    End With
        
    With tst
        .Assert tst, tst.wkbkTest.Sheets(shtSteps).Cells(1, 1) = "Sheet"
        .Assert tst, tblSteps.nCols = 9
        
        .Update tst, procs
    End With
End Sub
'-----------------------------------------------------------------------------------------------
'-----------------------------------------------------------------------------------------------
'Testing for tblRowsCols.Provision
'-----------------------------------------------------------------------------------------------
'-----------------------------------------------------------------------------
'Provision empty default table with nrows and ncols specified
'
'JDL 2/14/22
'
Sub test_ProvisionTbl2EmptySpec(procs)
    Dim tst As New Test
    tst.Init tst, "test_ProvisionTbl2EmptySpec"
    
    'Test-specific code
    '------------------
    Dim tbl As Object
    Set tbl = ExcelSteps.New_tbl
    PopulateTbl2 tst.wkbkTest, shtTbl, IsHeader:=False, IsData:=False
    
    'Empty table but with pre-defined nrows and ncols
    tbl.Provision tbl, tst.wkbkTest, False, sht:=shtTbl, nRows:=3, nCols:=4
    
    With tst
        .Assert tst, (tbl.lastcol = 4)
        .Assert tst, (tbl.lastrow = 4)
        .Assert tst, (tbl.cellHome.Address = "$A$2")
        .Assert tst, (tbl.nRows = 3)
        .Assert tst, (tbl.nCols = 4)
        .Assert tst, (tbl.rngrows.Address = "$2:$4")
        .Assert tst, (tbl.rngTable.Address = "$A$1:$D$4")
        .Assert tst, (tbl.rngHeader.Address = "$A$1:$D$1")
    
        .Update tst, procs
    End With
End Sub
'-----------------------------------------------------------------------------
'Provision completely empty default table
'
'JDL 2/14/22; 3/6/23 for ExcelSteps tblRowsCols
'
Sub test_ProvisionTbl2EmptyTbl(procs)
    Dim tst As New Test
    tst.Init tst, "test_ProvisionTbl2EmptyTbl"
    
    'Test-specific code
    '------------------
    Dim tbl As Object
    Set tbl = ExcelSteps.New_tbl
    PopulateTbl2 tst.wkbkTest, shtTbl, IsHeader:=False, IsData:=False
    tbl.Provision tbl, tst.wkbkTest, False, sht:=shtTbl
    
    With tst
        .Assert tst, (tbl.rngTable Is Nothing)
        .Assert tst, (tbl.rngrows Is Nothing)
        .Assert tst, (tbl.rngHeader Is Nothing)
        .Assert tst, (tbl.lastcol = 0)
        .Assert tst, (tbl.lastrow = 0)
        .Assert tst, (tbl.cellHome.Address = "$A$2")
        .Assert tst, (tbl.nRows = 0)
        .Assert tst, (tbl.nCols = 0)
    
        .Update tst, procs
    End With
End Sub
'-----------------------------------------------------------------------------
'Provision rows/columns table header only
'
'JDL 2/14/22; modified 3/6/23 for ExcelSteps tblRowsCols
'
Sub test_ProvisionTbl2HeaderOnly(procs)
    Dim tst As New Test
    tst.Init tst, "test_ProvisionTbl2HeaderOnly"
    
    'Test-specific code
    '------------------
    Dim tbl As Object
    Set tbl = ExcelSteps.New_tbl
    PopulateTbl2 tst.wkbkTest, shtTbl, IsData:=False
    tbl.Provision tbl, tst.wkbkTest, False, sht:=shtTbl
    
    With tst
        .Assert tst, (tbl.rngTable.Address = "$A$1:$F$1")
        .Assert tst, (tbl.rngrows Is Nothing)
        .Assert tst, (tbl.lastcol = 6)
        .Assert tst, (tbl.lastrow = 0)
        .Assert tst, (tbl.rngHeader.Address = "$A$1:$F$1")
        .Assert tst, (tbl.nRows = 0)
        .Assert tst, (tbl.nCols = 6)
    
        .Update tst, procs
    End With
End Sub
'-----------------------------------------------------------------------------
'Provision rows/columns table (new tblRowsCols class)
'
'JDL 2/14/22
'
Sub test_ProvisionTbl2(procs)
    Dim tst As New Test
    tst.Init tst, "test_ProvisionTbl2"
    
    'Test-specific code
    '------------------
    Dim tbl As Object
    Set tbl = ExcelSteps.New_tbl
    
    PopulateTbl2 tst.wkbkTest, shtTbl
    tbl.Provision tbl, tst.wkbkTest, False, sht:=shtTbl
    
    With tst
        .Assert tst, (tbl.rngTable.Address = "$A$1:$F$6")
        .Assert tst, (tbl.lastcol = 6)
        .Assert tst, (tbl.lastrow = 6)
        .Assert tst, (tbl.rngHeader.Address = "$A$1:$F$1")
        .Assert tst, (tbl.nRows = 5)
        .Assert tst, (tbl.nCols = 6)
        .Assert tst, (tbl.rngrows.Address = "$2:$6")
    
        .Update tst, procs
    End With
End Sub
'-----------------------------------------------------------------------------
'Populate a rows/columns table
'
'JDL 2/14/22
'
Sub test_PopulateTbl2(procs)
    Dim tst As New Test
    tst.Init tst, "test_PopulateTbl2"
    
    'Test-specific code
    '------------------
    PopulateTbl2 tst.wkbkTest, shtTbl
    
    With tst
        .Assert tst, (tst.wkbkTest.Sheets(shtTbl).Cells(1, 1) = "Desc")
        .Assert tst, (tst.wkbkTest.Sheets(shtTbl).Cells(1, 6) = "Data_3")
        .Assert tst, (tst.wkbkTest.Sheets(shtTbl).Cells(6, 1) = "D")
        .Assert tst, (tst.wkbkTest.Sheets(shtTbl).Cells(6, 6) = 9.6)
    
        .Update tst, procs
    End With
End Sub
'-----------------------------------------------------------------------------
'This section tests tbl name case sensitivity
' JDL 1/20/22
'-----------------------------------------------------------------------------
'Utility function to correct case sensitivity error with sheet names
'Returns Boolean but also corrects arg spelling to exact match of Sheet name
'JDL 5/22/24
'
Sub test_IsShtCaseErr(procs)
    Dim tst As New Test
    tst.Init tst, "test_IsShtCaseErr"
    
    'tst-specific code
    '------------------
    Dim sht_name As String
    
    'No Error - case of sht_name argument matches sheet name
    sht_name = "SMdl"
    tst.Assert tst, Not IsShtCaseErr(tst.wkbkTest, sht_name)
    tst.Assert tst, sht_name = "SMdl"
    
    'sht_name argument doesn't match sheet name (ByRef correction of sht)
    sht_name = "smdl"
    tst.Assert tst, IsShtCaseErr(tst.wkbkTest, sht_name)
    tst.Assert tst, sht_name = "SMdl"
    '------------------
    
    'Update test results on procs.wksht_results
    tst.Update tst, procs
End Sub

'Provision a rows/columns table
'
'JDL 1/20/22; 3/6/23 revise to point to ExcelSteps; Updated 5/28/24
'
Sub test_ProvisionTbl(procs)
    Dim tst As New Test
    tst.Init tst, "test_ProvisionTbl"
    
    'Test-specific code
    '------------------
    Dim tbl As Object
    Set tbl = ExcelSteps.New_tbl
    
    'Default - no range naming
    PopulateTbl tst.wkbkTest, shtTbl
    With tbl
        tst.Assert tst, .Provision(tbl, tst.wkbkTest, IsFormat:=True, sht:=shtTbl)
        tst.Assert tst, .rngrows.Address = "$2:$4"
        tst.Assert tst, .rngHeader.Address = "$A$1:$C$1"
    End With
    
    'With range naming
    PopulateTbl tst.wkbkTest, shtTbl
    With tbl
        tst.Assert tst, .Provision(tbl, tst.wkbkTest, IsFormat:=True, sht:=shtTbl, _
            IsSetTblNames:=True, IsSetColNames:=True)
        tst.Assert tst, .rngrows.Address = "$2:$4"
        tst.Assert tst, .rngHeader.Address = "$A$1:$C$1"
        tst.Assert tst, tst.wkbkTest.Names("SMdl").RefersTo = "=SMdl!$A:$C"
        tst.Assert tst, tst.wkbkTest.Names("SMdl_Header").RefersTo = "=SMdl!$1:$1"
        tst.Assert tst, tst.wkbkTest.Names("Col_A").RefersTo = "=SMdl!$A:$A"

    End With
    
    tst.Update tst, procs
End Sub
'Populate a rows/columns table
'
'JDL 1/20/22; updated 3/6/23
'
Sub test_PopulateTbl(procs)
    Dim tst As New Test
    tst.Init tst, "test_PopulateTbl"
    
    'Test-specific code
    '------------------
    PopulateTbl tst.wkbkTest, shtTbl
    With tst.wkbkTest.Sheets(shtTbl)
        tst.Assert tst, (.Cells(1, 1) = "Col_A")
        tst.Assert tst, (.Cells(2, 1) = "a")
        tst.Assert tst, (.Cells(4, 3) = 30)
    End With
    
    tst.Update tst, procs
End Sub







