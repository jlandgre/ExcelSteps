VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Procedures"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Version 5/30/24
Option Explicit
'This class defines project specific groups of tests to run in a tests_xxxx module
'Build it from an Excel code plan by declaring an object for each procedure in the plan

'---------------------------------------------------------------------------------------
' Individual procedure objects (project-specific)
Public tblInit As Object
Public tblSetDimensions As Object
Public tblSetArysNamesRngs As Object
Public tblFormat As Object
Public tblProvision As Object
Public tblRefresh As Object
Public tblRefreshAPI As Object

' Individual procedure objects (project-specific)
Public mdlInit As Object
Public mdlRow As Object
Public mdlVariations As Object
Public mdlDropdowns As Object
Public mdlRefreshSpeed As Object '7/15/25

Public Utilities As Object '10/23/25
Public LiteModelSpeed As Object '10/23/25

Public ParseModel As Object '11/17/25
Public mdlImportRow As Object '1/28/26
Public dictionary As Object '1/29/26
Public ErrorHandling As Object '3/11/26
Public RecordErr As Object '3/12/26

' General attributes for all procedures in the code project
Public wkbk_testing As Workbook
Public wksht_results As Worksheet
Public test_suite_name As String
Public curProcedure As String 'Name of current procedure

'Column ranges on wksht_results
Public colrngProcedure As Range
Public colrngTestName As Range
Public colrngTestResult As Range
Public rngHeader As Range
'---------------------------------------------------------------------------------------
' Instance Procedure objects for each TestGroup
' JDL 5/2/25; Updated 1/28/26
Sub Init(procs, ByVal wkbk, ByVal shtT, test_suite_name)
    With procs
        ' TestDriver_TblRowsCols
        Set .tblInit = New Procedure
        .tblInit.Name = "tblRowsCols.Init"
        Set .tblSetDimensions = New Procedure
        .tblSetDimensions.Name = "tblRowsCols.SetDimensions"
        Set .tblSetArysNamesRngs = New Procedure
        .tblSetArysNamesRngs.Name = "tblRowsCols.SetArysNamesRngs"
        Set .tblFormat = New Procedure
        .tblFormat.Name = "tblRowsCols.Format"
        Set .tblProvision = New Procedure
        .tblProvision.Name = "tblRowsCols.Provision"
        Set .tblRefresh = New Procedure
        .tblRefresh.Name = "tblRowsCols.Refresh"
        Set .tblRefreshAPI = New Procedure
        .tblRefreshAPI.Name = "tbl_RefreshAPI"
        
        ' TestDriver_mdlScenario
        Set .mdlInit = New Procedure
        .mdlInit.Name = "mdlScenario.Init"
        Set .mdlRow = New Procedure
        .mdlRow.Name = "mdlRow class"
        Set .mdlVariations = New Procedure
        .mdlVariations.Name = "mdl Variations"
        Set .mdlDropdowns = New Procedure
        .mdlDropdowns.Name = "Dropdown list capability"
        Set .mdlRefreshSpeed = New Procedure
        .mdlRefreshSpeed.Name = "Refresh speedup"
        
        ' TestDriver_Utilities
        Set .Utilities = New Procedure
        .Utilities.Name = "ExcelSteps Utilities"
        Set .LiteModelSpeed = New Procedure
        .LiteModelSpeed.Name = "Lite Model Prov/Refresh Speedup"
        
        ' TestDriver_ParseModel
        Set .ParseModel = New Procedure
        .ParseModel.Name = "ParseModel"
        
        ' TestDriver_mdlImportRow
        Set .mdlImportRow = New Procedure
        .mdlImportRow.Name = "mdlImportRow Class"
        
        ' TestDriver_Dictionary
        Set .dictionary = New Procedure
        .dictionary.Name = "Dictionary Class"

        ' TestDriver_ErrorHandling
        Set .ErrorHandling = New Procedure
        .ErrorHandling.Name = "ErrorHandling Class"
        Set .RecordErr = New Procedure
        .RecordErr.Name = "ErrorHandling.RecordErr"
        
        ' Set attributes and initialize test results sheet
        .test_suite_name = test_suite_name
        Set .wkbk_testing = wkbk
        
        ' Initialize the Test Result sheet including creating it if it's not present
        .InitTestsSheet procs, shtT

    End With
End Sub
'-----------------------------------------------------------------------------------------------
'Initialize sheet for reporting results (clears previous results)
'Inputs: wkbk [Workbook] contains tests sheet
'        shtT [String] tests sheet name for reporting test results
'
'Created:   11/10/21 JDL; Modified 11/8/24
'
Sub InitTestsSheet(procs, ByVal shtT)
    
    'Add tests sheet if it does not exist and set Worksheet object
    With procs.wkbk_testing
        If Not SheetExists(procs.wkbk_testing, shtT) Then _
            AddSheet procs.wkbk_testing, shtT, .Sheets(.Sheets.Count).Name
        Set procs.wksht_results = .Sheets(shtT)
    End With
    
    With procs
        Set .colrngProcedure = .wksht_results.Columns(1)
        Set .colrngTestName = .wksht_results.Columns(2)
        Set .colrngTestResult = .wksht_results.Columns(3)
        Set .rngHeader = Intersect(.wksht_results.Rows(1), _
        Range(.wksht_results.Columns(1), .wksht_results.Columns(3)))
                
        'Clear the results columns and refresh header strings
        .colrngTestResult.Clear
        .rngHeader = Split("Procedure,Test Name,Test Result", ",")
        
        'Clear previous yellow highlighting of failed results if any
        .colrngTestResult.Interior.Pattern = xlNone
        
        ColWidthAutofit Range(.colrngProcedure, .colrngTestResult)
    End With

End Sub
'-----------------------------------------------------------------------------------------------
' Evaluate overall test results
'
'Inputs:  wkbk [Workbook] workbook containing test results table
'         shtT [String] worksheet name containing test results table
'         sTestSuite [String] name of test suite
'
'   JDL 11/18/21; Modified 10/1/24
Sub EvalOverall(procs)
    Dim w As Variant, IsPass As Boolean, curCell As Range, sMsg As String, rngcol As Range
    
    With procs
        IsPass = True
        Set curCell = Intersect(.wksht_results.Rows(2), .colrngTestName)
        Do While Len(curCell.Value) > 0
                
            'w = curCell.Offset(0, 1).Value
            w = Intersect(curCell.EntireRow, .colrngTestResult).Value
            If Not IsEmpty(w) Then IsPass = IsPass And w
            Set curCell = curCell.Offset(1, 0)
        Loop
        
        'Report overall result to user
        sMsg = procs.test_suite_name & " Testing Failed"
        If IsPass Then sMsg = procs.test_suite_name & " Testing Passed"
        MsgBox sMsg
        
        'Autofit columns
        For Each rngcol In Range(.colrngProcedure, .colrngTestResult)
        With rngcol
            .ColumnWidth = 120
            .AutoFit
            If .ColumnWidth < 120 Then .ColumnWidth = .ColumnWidth + 2
        End With
    Next rngcol

    End With
End Sub




