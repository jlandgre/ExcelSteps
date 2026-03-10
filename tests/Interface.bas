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
'Sub test_tbl_refresh()
'    ExcelSteps.RefreshAPI ThisWorkbook, "SMdl4a", IsTbl:=True
'End Sub

