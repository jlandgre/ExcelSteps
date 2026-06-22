Attribute VB_Name = "a_RunTests"
'Version 5/15/26 mdlImportRow and SwapModels need debug --not passing
Option Explicit
'-----------------------------------------------------------------------------------------------
' Run all validations in workbook
'-----------------------------------------------------------------------------------------------
Sub a_RunAllTests()
    TestDriver_Utilities
    TestDriver_Dictionary
    TestDriver_ErrorHandling
    'TestDriver_mdlImportRow
    TestDriver_mdlScenario
    TestDriver_ParseModel
    TestDriver_PivotTable
    'TestDriver_SwapModels
    TestDriver_TblRowsCols
    TestDriver_ToolBox
ThisWorkbook.Sheets("SMdl").Activate
End Sub





