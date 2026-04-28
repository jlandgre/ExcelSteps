Attribute VB_Name = "VersionHistory"
'This module contains only comments and records the version history
'
' 1/29/20 Added line to refresh form initialization to set focus on OK button
' 3/9/20 Fixed bug that prevented Freeze Row 1 step from working
' 3/12/20 Add logic to delete Pandas index column from ExcelSteps recipe for imported recipe
' 3/23/20 Eliminate individual formatting commands from step list (in lieu of col_Format
' 3/27/20 Fix bug with very wide columns that Autofit to 255 width
' 3/27/20 Fix bug that caused VBA error if Col_AddGroup start column not found
' 3/28/20 Fix bug if number of sheets exceeds number of slots on refresh dialog
' 4/3/20 Fix bug if AddGroup column is hidden by collapsed outline during refresh
' 4/3/20 Fixed bug if column header is also an Excel column name (e.g. "VAL123")
' 4/3/20 Fixed bug with table dimension if column is inserted at right edge of table
' 4/3/20 Added FreezeRow1 to ExcelSteps sheet to deal well with longer recipes
' 4/3/20 Modified step function calls to avoid passing Step class (better enable use as API)
' 4/10/20 Changed modUtilities AddComment sub to rngCell.ClearNotes to deal with
'         distinction between comments and notes (caused VBA error previously
' 5/9/20  Added and performed initial verification of modScenario to handle scenario models
' 1/11/21 Paste in updated versions of modConstants, modRefresh, modScenario, modUtilities
'         - updates to allow rows/columns refresh (as API) in tables that are not homed
'         to Cells(1,1)
' 1/11/21 Bug fixes to modScenario
' 1/11/21 Add rngheaders and rngdata as arguments to RefreshRC() to enable call as API with
'         table pre-defined from a Class Instance
' 7/29/21 Add SplitCols recipe option, Add DeleteCol recipe option
' 7/30/21 Refactor to clsRefresh instead of individual arguments for RefreshRC
' 7/30/21 Use tblRowsCols class instance to hold transformed table description
' 7/30/21 Convert ExcelSteps table description to tblRowsCols class instance
' 7/30/21 Move RefreshRC table formatting to tblRowsCols FormatTbl sub
' 9/15/21 Fix bug so that table range name includes inserted first column (Column A)
' 9/16/21 Add gray shading to StepCondFormat() function to lightly highlight white text values
' 9/23/21 Fix bug with sheet name prefix setting for Rows/Columns refresh
' 9/23/21 Add scenario model refresh to menu
' 9/23/21 Add parse scenario model to menu
' 11/15/21 Fix frmRefresh bug - use loop counter for dual purpose (crash Scenario Model refresh)
' 8/23/22 Integrate updated Scenario Model code (mdlScenario and associated utils) from
'         valRefresh.xlsm
'12/6/22 Add StepRename/rename column instruction
'12/7/22 Add frmRefresh code snippet to unselect if multiple sheets initially selected
'3/1/23 Consolidate standalone validation file with ExcelSteps addin
'3/14/23 SwapModels feature validated to swap rows/cols model into Scenario Model and
'        original Scenario Model into rows/cols
'4/14/23 Bug fix in frmRefresh - mdl.Refresh call needs to return Boolean
'7/10/24 Bug fix in frmRefresh to not omit setting write for last refreshed sheet
'8/6/24 Add wkbkResetStatus call in case of errors in frmRefresh.CmdButOK_Click
'10/24/24 significant refactoring of tests; modify Refresh and RecipeStep for optimized performance
'11/6/24 modify tblRowsCols.SetIsCustomTbl and other methods to enable additional arg combos
'       for sht, TblName and defn to define custom and default tables
'11/6/24 Update Errors_ using Github Copilot prompt to generate table directly from code
'12/12/24 Update mdlScenario.NameMdlColumns() to append .NamePrefix before naming columns
'12/13/24 Ensure previous active sheet gets reactivated after refresh (modInterface.WkbkResetStatus)
'12/13/24 fix logic problems with tblRowsCols.SetIsCustomTbl and .SetCustomTblParams
'1/10/25 Add Delete_Flagged step type; mods to Constants and RecipeStep classes
'1/13/25 fix bug in mdlScenario.SetAttsFromArgs for .sht not a valid .MdlName aka range name
'2/13/25 minor updates to frmRefresh comments to clarify .RefreshRC args used
'2/15/25 Refactor mdlScenario.Init and .ParseMdlScenDefn; eliminate bugs in setting .sht and .mdlName
'5/2/25 Bug fix to FindInRange
'5/6/25 Change to black font for formulas in tables and scenario models (classes RecipeStep and xxx)
'5/9/25 Fix bug in mdlRow for case where Lite model has no instructions on ExcelSteps
'6/15/25 fix bug with RecipeStep.Insert and other subs to replace .Find with FindInRange (partial update)
'7/8/25 fix bug with 6/15 mod; add new utility to modUtilities
'7/21/25 mdl.Refresh speedup (see notes in test suite); Add mdl.iColorGrpRows attribute and code
'7/30/25 Add ParamBlock Class and tests
'8/8/25 Fix bug with mdl.ScenModelLoc and .SetScenModelLoc where multiple models on same sheet
'       cause there to be multiple Scenario variables. Need to use Intersect with mdl.rngRows
'8/12/25 Updates to mdlScenario and modUtilities
'8/15/25 Fix bug with SaveAsCloseOverwrite utility on Mac
'9/12/25 add tblRowsCols.colCur attribute
'9/30/25 add FillRightBySegments utility function;
'10/1/25 modUtilities cleanup and update Errors_
'10/16/25 Add SafeSaveAs to modUtilities to enable saving to iCloud folders with Mac Excel;
'         Update ErrorHandling ShowMessage
'10/23/25 Minor debugging of mdlScenario
'10/24/25 Update SafeSaveAs to work with or without ErrorHandling (errs.IsHandle = True or False)
'11/6/25 Optimized performance (see mdlScenario.Refresh)
'11/17/25 Add mdlScenario IsCustomMdl attr to facilitate ParseMdl option detection; rewrite ParseMdl
'11/19/25 Rewritten/validated modParseSM and modInterface.ParseModelDriver to parse mdl to rows/cols
'12/3/25 Code review/minor updates to many modules; ErrorHandling update; Errors table updated
'12/16/25 minor docstring updates - Add SetExcelStepsVersion to tests Utilities_Testing module
'1/29/26 Cosmetic changes for consistency with use of xlwings vba edit; clean up module names
'3/10/26 Add Dictionary.ParseStringToDictProcedure; update Errors_ table
'3/31/26 Add validated PivotTable class (.MakePivotTableProcedure)
'4/10/26 Fix bug with ParseMdl function
'4/16/26 Fix bug/robustness issue with Dictionary.Add
'4/27/26 Update ErrorHandling (version 4/28/26 has error handling enabled)


