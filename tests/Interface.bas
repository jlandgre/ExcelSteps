Attribute VB_Name = "Interface"
'Version 10/24/24
Option Explicit
'Subs activated by user buttons on Descriptions sheet
Sub CreateSMdl1()
    Dim procs As New Procedures, AllEnabled As Boolean
    procs.Init procs, ThisWorkbook, shtT_temp, "mdlScenario"
    SetApplEnvir False, False, xlCalculationManual
    test_RefreshSMdl1 procs
    ThisWorkbook.Sheets("SMdl").Activate
    SetApplEnvir True, True, xlCalculationAutomatic
End Sub
Sub CreateSMdl4a()
    Dim procs As New Procedures, AllEnabled As Boolean
    procs.Init procs, ThisWorkbook, shtT_temp, "mdlScenario"
    SetApplEnvir False, False, xlCalculationManual
    test_RefreshSMdl4a procs
    ThisWorkbook.Sheets("SMdl").Activate
    SetApplEnvir True, True, xlCalculationAutomatic
    SetApplEnvir True, True, xlCalculationAutomatic
End Sub
Sub CreateSMdl2()
    Dim procs As New Procedures, AllEnabled As Boolean
    procs.Init procs, ThisWorkbook, shtT_temp, "mdlScenario"
    SetApplEnvir False, False, xlCalculationManual
    test_RefreshSMdl2 procs
    ThisWorkbook.Sheets("SMdl").Activate
    SetApplEnvir True, True, xlCalculationAutomatic
End Sub
Sub CreateSMdl3()
    Dim procs As New Procedures, AllEnabled As Boolean
    procs.Init procs, ThisWorkbook, shtT_temp, "mdlScenario"
    SetApplEnvir False, False, xlCalculationManual
    test_RefreshSMdl3 procs
    ThisWorkbook.Sheets("SMdl").Activate
    SetApplEnvir True, True, xlCalculationAutomatic
End Sub
Sub CreateSMdl4()
    Dim procs As New Procedures, AllEnabled As Boolean
    procs.Init procs, ThisWorkbook, shtT_temp, "mdlScenario"
    SetApplEnvir False, False, xlCalculationManual
    test_RefreshSMdl4 procs
    ThisWorkbook.Sheets("SMdl").Activate
    SetApplEnvir True, True, xlCalculationAutomatic
End Sub
Sub CreateSMdl5()
    Dim procs As New Procedures, AllEnabled As Boolean
    procs.Init procs, ThisWorkbook, shtT_temp, "mdlScenario"
    SetApplEnvir False, False, xlCalculationManual
    test_RefreshSMdl5 procs
    ThisWorkbook.Sheets("SMdl").Activate
    SetApplEnvir True, True, xlCalculationAutomatic
End Sub
Sub CreateSMdl6()
    Dim procs As New Procedures, AllEnabled As Boolean
    procs.Init procs, ThisWorkbook, shtT_temp, "mdlScenario"
    SetApplEnvir False, False, xlCalculationManual
    test_RefreshSMdl6 procs
    ThisWorkbook.Sheets("SMdl").Activate
    SetApplEnvir True, True, xlCalculationAutomatic
End Sub
Sub CreateSMdl7()
    Dim procs As New Procedures, AllEnabled As Boolean
    procs.Init procs, ThisWorkbook, shtT_temp, "mdlScenario"
    SetApplEnvir False, False, xlCalculationManual
    test_RefreshSMdl7 procs
    ThisWorkbook.Sheets("SMdl").Activate
    SetApplEnvir True, True, xlCalculationAutomatic
End Sub
Sub CreateSMdl2_Dropdown()
    Dim procs As New Procedures, AllEnabled As Boolean
    procs.Init procs, ThisWorkbook, shtT_temp, "mdlScenario"
    SetApplEnvir False, False, xlCalculationManual
    test_AddDropdownSMdl2 procs
    ThisWorkbook.Sheets("SMdl").Activate
    Cells(1, 1).Select
    SetApplEnvir True, True, xlCalculationAutomatic
End Sub
Sub CreateSMdl5_Dropdown()
    Dim procs As New Procedures, AllEnabled As Boolean
    procs.Init procs, ThisWorkbook, shtT_temp, "mdlScenario"
    SetApplEnvir False, False, xlCalculationManual
    test_AddDropdownSMdl5 procs
    ThisWorkbook.Sheets("SMdl").Activate
    Cells(1, 1).Select
    SetApplEnvir True, True, xlCalculationAutomatic
End Sub
Sub CreateSwapModelsDemo()
    Dim procs As New Procedures, AllEnabled As Boolean
    procs.Init procs, ThisWorkbook, shtT_temp, "mdlScenario"
    SetApplEnvir False, False, xlCalculationManual
    test_SwapModels4 ThisWorkbook, "Tests_SwapModels"
    ThisWorkbook.Sheets("SMdl").Activate
    SetApplEnvir True, True, xlCalculationAutomatic
End Sub
Sub RefreshTblFromRecipe()
    Dim procs As New Procedures, AllEnabled As Boolean
    procs.Init procs, ThisWorkbook, shtT_temp, "tblRowsCols"
    SetApplEnvir False, False, xlCalculationManual
    test_RefreshTbl3 procs
    ThisWorkbook.Sheets("SMdl").Activate
    SetApplEnvir True, True, xlCalculationAutomatic
End Sub

'--------------------------------------------------------------------------------------
' Build a simple pivot with sum analyte and verify category/subcategory totals
' SubCategory as colField
' JDL 3/31/26
'
Sub PivotTableDemo1()
    Dim tst As New Test: tst.Init tst, "test_MakePivotTableProcedure1"
    Set ExcelSteps.errs = Nothing
    Dim tblSrc As Object, pvt As Object, params As Object
    Dim rowFields As Variant, colFields As Variant, analytes As Variant

    If Not PopulateHelper(tblSrc, pvt, params, True) Then GoTo ErrorExit

    rowFields = Array("Category")
    colFields = Array("SubCategory")
    analytes = Array(Array("Amount", xlSum))
    params.Add "isReplaceWithVals", False

    If Not pvt.MakePivotTableProcedure(pvt, tblSrc, rowFields, colFields, "PivotOut", _
        analytes, dictParams:=params) Then GoTo ErrorExit
        
    With ThisWorkbook.Sheets(pvt.shtDest)
        .Activate
        Range(.Columns(2), .Columns(3)).ColumnWidth = 14
    End With
    Exit Sub
ErrorExit:
    MsgBox "Error creating demo Pivot Table"
End Sub
'--------------------------------------------------------------------------------------
' Two rowFields; no colField; Column grand Total enabled
' JDL 3/31/26
'
Sub PivotTableDemo2()
    Dim tst As New Test: tst.Init tst, "test_MakePivotTableProcedure1"
    Set ExcelSteps.errs = Nothing
    Dim tblSrc As Object, pvt As Object, params As Object
    Dim rowFields As Variant, colFields As Variant, analytes As Variant

    If Not PopulateHelper(tblSrc, pvt, params, True) Then GoTo ErrorExit

    rowFields = Array("Category", "SubCategory")
    colFields = vbNullString
    analytes = Array(Array("Amount", xlSum))
    
    params.Add "isColGrand", True
    params.Add "isReplaceWithVals", False

    If Not pvt.MakePivotTableProcedure(pvt, tblSrc, rowFields, colFields, "PivotOut", _
        analytes, dictParams:=params) Then GoTo ErrorExit
        
    ThisWorkbook.Sheets(pvt.shtDest).Activate
    Exit Sub
ErrorExit:
    MsgBox "Error creating demo Pivot Table"
End Sub

'--------------------------------------------------------------------------------------
' Values field in rowFields to summarize multiple metrics by time colFields
' JDL 3/31/26
'
Sub PivotTableDemo3()
    Dim tst As New Test: tst.Init tst, "test_MakePivotTableProcedure1"
    Set ExcelSteps.errs = Nothing
    Dim tblSrc As Object, pvt As Object, params As Object
    Dim rowFields As Variant, colFields As Variant, analytes As Variant

    If Not PopulateHelper(tblSrc, pvt, params, False) Then GoTo ErrorExit

    rowFields = Array("Store", "Prodtype")
    colFields = Array("Week", "Year")
    analytes = Array(Array("Discounts", xlSum), Array("Markdowns", xlSum), Array("COGS", xlSum))

    ' Place Values field in rows between Store and Prodtype (Store, Sigma Values, Prodtype)
    params.Add "DataPivotFieldOrientation", xlRowField
    params.Add "DataPivotFieldPosition", 2
    params.Add "isReplaceWithVals", False

    If Not pvt.MakePivotTableProcedure(pvt, tblSrc, rowFields, colFields, "PivotOut", _
        analytes, dictParams:=params) Then GoTo ErrorExit
        
    ThisWorkbook.Sheets(pvt.shtDest).Activate
    Exit Sub
ErrorExit:
    MsgBox "Error creating demo Pivot Table"
End Sub
Function PopulateHelper(tblSrc, pvt, params, isSimple) As Boolean
    PopulateHelper = True
    
    'Populate data onto SMdl sheet
    If isSimple Then
        PopulatePivotTableSimple ThisWorkbook, shtPivotSrc
    Else
        PopulatePivotTableOTBLike ThisWorkbook, shtPivotSrc
    End If
    
    'Initialize objects
    Set tblSrc = ExcelSteps.New_tbl
    Set pvt = ExcelSteps.New_PivotTable
    Set params = ExcelSteps.New_Dictionary
    
    'Refresh and Provision table
    If Not ExcelSteps.RefreshTblAPI(ThisWorkbook, IsReplace:=True, IsTblFormat:=True, _
        sht:=shtPivotSrc) Then GoTo ErrorExit
    If Not tblSrc.Provision(tblSrc, ThisWorkbook, False, sht:=shtPivotSrc) Then GoTo ErrorExit
    Exit Function
ErrorExit:
    PopulateHelper = False
End Function

