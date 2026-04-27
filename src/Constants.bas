Attribute VB_Name = "Constants"
'Version 4/27/26
'This module is part of the ExcelSteps open source project posted at:
'https://github.com/jlandgre/ExcelSteps/. It is licensed under the MIT open source license

Option Explicit
Public Const Version As String = "version 4/27/26"
Public Const iMinRows As Integer = 10
Public Const sRefreshSuffix As String = "_t"

'Sheet names
Public Const shtSteps As String = "ExcelSteps"
Public Const shtSettings As String = "Settings_"
Public Const shtLists As String = "Lists"
Public Const shtColInfo As String = "colinfo"

'Settings
Public Const setting_shtFrm As String = "ShtNameFrm"
Public Const setting_cBoxFrm1 As String = "cBoxAppShtName"

'Strings for checking name validity
Public Const sXLChars As String = "abcdefghijklmnopqrstuvwxyz0123456789._"
Public Const sXLFirstChars As String = "abcdefghijklmnopqrstuvwxyz"
Public Const sNumbers As String = "0123456789"

'Steps (sStepFunctions is list by function --including comment, width etc.)
Public Const sStepList As String = "Col_Format,Col_Insert,Col_Delete,Col_Rename,Col_AddGroup," _
    & "Col_CondFormat,Col_Dropdown,Tbl_FreezeRow1,Tbl_Sort,Tbl_SplitCols,Delete_FlagRows," _
    & "Delete_RowsWithVal"

'Public Const sStepFunctions As String = "Col_Delete,Col_Rename,Col_Insert,Col_AddGroup,Col_Comment," _
'    & "Col_CondFormat,Col_Dropdown,Col_NumFormat,Col_Width,Tbl_FreezeRow1,Tbl_Sort,Tbl_SplitCols"
    
Public Const sAFormat As String = "Col_Format"
Public Const sADelete As String = "Col_Delete"
Public Const sARename As String = "Col_Rename"
Public Const sAInsert As String = "Col_Insert"
Public Const sAComment As String = "Col_Comment"
Public Const sAGroup As String = "Col_AddGroup"
Public Const sACondFmt As String = "Col_CondFormat"
Public Const sADropdown As String = "Col_Dropdown"
Public Const sANumFmt As String = "Col_NumFormat"
Public Const sAWidth As String = "Col_Width"
'Public Const sANameRows As String = "Tbl_NameRowsBy"
Public Const sASort As String = "Tbl_Sort"
Public Const sAFreezeRow1 As String = "Tbl_FreezeRow1"
Public Const sASplitCols As String = "Tbl_SplitCols"

'Added 1/10/25
Public Const sADelFlagRows As String = "Delete_FlagRows"
Public Const sADelRowsWithVal As String = "Delete_RowsWithVal"

'Menu
Public Const sRefresh As String = "&Refresh Workbook Tables"
Public Const sParseSM As String = "&Parse Scenario Model"
Public Const sAbout As String = "About ExcelSteps"

'Workbook Status: TRUE = auto calculation with events enabled when macros not running
Public Const IsDefaultStatus As Boolean = True

'Workbook Status Setting names
'Public Const sStatusOrig As String = "wkbkstatus_orig"
Public Const sStatusRun As String = "wkbkstatus_run"
Public Const sStatusDefault As String = "wkbkstatus_default"

Public Const ScenHeader As String = "Grp,Subgrp,Description,Variable Names,Units,Number Fmt,Formula/Row Type"
Public Const sLstSettingsHeader As String = "Setting Name,Value"

'Additional Constants used by mdlScenario
Public Const shtTblImp As String = "TblImport"
Public Const ScenHeaderLite As String = "Description,Variable Names,Units"
Public Const sHeaderMdlImport As String = "Model,Grp,Subgrp,Description,Variable Name," _
    & "Units,Number Fmt,Formula/Row Type,Scenario Name,Value"

'Constants related to ExcelSteps
'Public Const sLstSteps As String = "Col_Delete,Col_Insert,Col_AddGroup," _
'    & "Col_Comment,Col_CondFormat,Col_Dropdown,Col_NumFormat,Col_Width," _
'    & "Tbl_FreezeRow1,Tbl_Sort,Tbl_SplitCols,Delete_FlagRows,Delete_RowsWithVal"
Public Const sHeaderSteps As String = "Sheet,Column,Step,Formula/List Name/Sort-by," _
    & "After End or Rename Column,Keep Formulas,Comment,Number Format,Width"
