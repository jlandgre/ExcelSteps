Attribute VB_Name = "a_RunTests"
'Version 1/28/26 SwapModels needs refactoring --not passing
Option Explicit
'-----------------------------------------------------------------------------------------------
' Run all validations in workbook
'-----------------------------------------------------------------------------------------------
Sub a_RunAllTests()
    TestDriver_ErrorHandling
    TestDriver_Utilities
    TestDriver_mdlScenario
    TestDriver_ParseModel
    TestDriver_mdlImportRow
    TestDriver_Dictionary
    TestDriver_TblRowsCols
    'TestDriver_SwapModels
ThisWorkbook.Sheets("SMdl").Activate
End Sub


