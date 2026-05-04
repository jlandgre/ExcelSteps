Attribute VB_Name = "Validation"
'Version 5/1/26
Option Explicit

'Global variable (default False for production) can toggle to True from tests workbook
Public IsTest As Boolean

'-----------------------------------------------------------------------------------------------------
' Factory functions below instance add-in objects from a second workbook such as a test suite
' workbook. To call, the second workbook's VBA Project needs to add a Reference to ExcelSteps
' (Tools / References menu in VBA editor), and the second workbook should instance by
' calling these modValidation functions as shown:
'
'Dim tbl as object
'Set tbl = ExcelSteps.new_tblRowsCols
'
'   <<or alternatively>>
'
'Set tbl = Application.Run(sDirPrefix_ExcelSteps & "New_tbl")
'
'   where sDirPrefix_ExcelSteps = "c:\dir1\dir2!" -- path to XLSteps.xlam
'
'JDL 12/15/22
'-----------------------------------------------------------------------------------------------------
Public Function New_mdl() As mdlScenario
    Set New_mdl = New mdlScenario
End Function
Public Function New_tbl() As tblRowsCols
    Set New_tbl = New tblRowsCols
End Function
Public Function New_Refresh() As Refresh
    Set New_Refresh = New Refresh
End Function
Public Function New_ErrorHandling() As ErrorHandling
    Set New_ErrorHandling = New ErrorHandling
End Function
Public Function New_ErrorsMeta() As ErrorsMeta
    Set New_ErrorsMeta = New ErrorsMeta
End Function
Public Function New_mdlRow() As mdlRow
    Set New_mdlRow = New mdlRow
End Function
Public Function New_mdlImportRow() As mdlImportRow
    Set New_mdlImportRow = New mdlImportRow
End Function
Public Function New_ParamBlock() As ParamBlock
    Set New_ParamBlock = New ParamBlock
End Function
Public Function New_Dictionary() As Dictionary
    Set New_Dictionary = New Dictionary
End Function
Public Function New_PivotTable() As PivotTable
    Set New_PivotTable = New PivotTable
End Function
Public Function New_ProjFiles() As ProjFiles
    Set New_ProjFiles = New ProjFiles
End Function
Public Function New_ColInfo() As ColInfo
    Set New_ColInfo = New ColInfo
End Function



