Attribute VB_Name = "UtilitiesMdl"
'Version 10/18/24 - paste from ExcelSteps modConstants
Option Explicit
' Public Const sVersion As String = "8/6/24"
' Public Const iMinRows As Integer = 10
' Public Const sRefreshSuffix As String = "_t"

' 'Sheet names
' Public Const shtSteps As String = "ExcelSteps"
' Public Const shtSettings As String = "Settings_"
' Public Const shtLists As String = "Lists"
' Public Const shtColInfo As String = "colinfo"

' 'Settings
' Public Const setting_shtFrm As String = "ShtNameFrm"
' Public Const setting_cBoxFrm1 As String = "cBoxAppShtName"

' 'Strings for checking name validity
' Public Const sXLChars As String = "abcdefghijklmnopqrstuvwxyz0123456789._"
' Public Const sXLFirstChars As String = "abcdefghijklmnopqrstuvwxyz"
' Public Const sNumbers As String = "0123456789"

' 'Steps (sStepFunctions is list by function --including comment, width etc.)
' Public Const sStepList As String = "Col_Format,Col_Insert,Col_Delete,Col_Rename,Col_AddGroup," _
'     & "Col_CondFormat,Col_Dropdown,Tbl_FreezeRow1,Tbl_Sort,Tbl_SplitCols"
' Public Const sStepFunctions As String = "Col_Delete,Col_Rename,Col_Insert,Col_AddGroup,Col_Comment," _
'     & "Col_CondFormat,Col_Dropdown,Col_NumFormat,Col_Width,Tbl_FreezeRow1,Tbl_Sort,Tbl_SplitCols"
    
' Public Const sAFormat As String = "Col_Format"
' Public Const sADelete As String = "Col_Delete"
' Public Const sARename As String = "Col_Rename"
' Public Const sAInsert As String = "Col_Insert"
' Public Const sAComment As String = "Col_Comment"
' Public Const sAGroup As String = "Col_AddGroup"
' Public Const sACondFmt As String = "Col_CondFormat"
' Public Const sADropdown As String = "Col_Dropdown"
' Public Const sANumFmt As String = "Col_NumFormat"
' Public Const sAWidth As String = "Col_Width"
' 'Public Const sANameRows As String = "Tbl_NameRowsBy"
' Public Const sASort As String = "Tbl_Sort"
' Public Const sAFreezeRow1 As String = "Tbl_FreezeRow1"
' Public Const sASplitCols As String = "Tbl_SplitCols"

' 'Menu
' Public Const sRefresh As String = "&Refresh Workbook Tables"
' Public Const sParseSM As String = "&Parse Scenario Model"
' Public Const sAbout As String = "About ExcelSteps"

' 'Workbook Status: TRUE = auto calculation with events enabled when macros not running
' Public Const IsDefaultStatus As Boolean = True

' 'Workbook Status Setting names
' 'Public Const sStatusOrig As String = "wkbkstatus_orig"
' Public Const sStatusRun As String = "wkbkstatus_run"
' Public Const sStatusDefault As String = "wkbkstatus_default"

' Public Const ScenHeader As String = "Grp,Subgrp,Description,Variable Names,Units,Number Fmt,Formula/Row Type"
' Public Const sLstSettingsHeader As String = "Setting Name,Value"

' 'Additional Constants used by mdlScenario
' Public Const shtTblImp As String = "TblImport"
' Public Const ScenHeaderLite As String = "Description,Variable Names,Units"
' Public Const sHeaderMdlImport As String = "Model,Grp,Subgrp,Description,Variable Name," _
'     & "Units,Number Fmt,Formula/Row Type,Scenario Name,Value"

' 'Constants related to ExcelSteps
' Public Const sLstSteps As String = "Col_Delete,Col_Insert,Col_AddGroup," _
'     & "Col_Comment,Col_CondFormat,Col_Dropdown,Col_NumFormat,Col_Width," _
'     & "Tbl_FreezeRow1,Tbl_Sort,Tbl_SplitCols"
' Public Const sHeaderSteps As String = "Sheet,Column,Step,Formula/List Name/Sort-by," _
'     & "After End or Rename Column,Keep Formulas,Comment,Number Format,Width"
    
'Version 8/29/23 Subs/Functions pasted from ExcelSteps
'-----------------------------------------------------------------------------------------------
' Build Comma-separated list from array
'
' Created: 3/1/21 JDL Modified 11/18/21 Add sDelim optional argument
'
' Function ListFromArray(ary, Optional sDelim, Optional IsFormatted As Boolean) As String
'     Dim val As Variant, lst As String, i As Integer
'     lst = ""
    
'     If IsMissing(IsFormatted) Then IsFormatted = False
'     If IsMissing(sDelim) Then sDelim = ","

'     'Simple delimited list
'     If Not IsFormatted Then
'         For Each val In ary
'             If Len(lst) < 1 Then
'                 lst = CStr(val)
'             Else
'                 lst = lst & sDelim & CStr(val)
'             End If
'         Next val
    
'     'Formatted, comma-separated list
'     Else
'         sDelim = ", "
'         For i = 0 To UBound(ary)
'             val = ary(i)
'             If i = 0 Then
'                 lst = CStr(val)
'             Else
'                 If i = UBound(ary) Then sDelim = " and "
'                 lst = lst & sDelim & CStr(val)
'             End If
'         Next i
'     End If
'     ListFromArray = lst
' End Function
' '-----------------------------------------------------------------------------------------------
' ' Set range for last used cell in a row (works with hidden and column outline cells)
' ' Inputs: cellHome [Range] home (left-most) cell in row to be searched
' '         xlDirection [Integer xlToRight or xlDown enumeration]
' ' finding value in outline-hidden cell. See 6/2/15 response on:
' ' https://stackoverflow.com/questions/20152328/vba-find-function-cant-find-given-value
' '
' ' JDL 12/3/21 (Based on Search Example.xlsm) Modified 12/6/21 to avoid intermittent not
' '
' Function rngLastPopCell(cell, xlDirection)
'     Dim rng As Range, c As Range, IsRowSearch As Boolean, IsColSearch As Boolean
'     Set rngLastPopCell = cell
    
'     'Set search range based on whether column or row search specified
'     If xlDirection = xlToRight Then
'         IsRowSearch = True
'         Set rng = cell.EntireRow
'     ElseIf xlDirection = xlDown Then
'         IsColSearch = True
'         Set rng = cell.EntireColumn
'     Else
'         Exit Function
'     End If
    
'     'To find last populated cell, Search with wrap from first cell
'     Set c = rng.Cells.Find("*", After:=rng.Cells(1), LookIn:=xlFormulas, _
'         SearchDirection:=xlPrevious)
'     If Not c Is Nothing Then Set rngLastPopCell = c
' End Function
' '------------------------------------------------------------------------------------------------
' ' Build multi-cell range containing non-empty cells in intersection of two ranges
' '           (Unlike .End(xlUp) etc., it works with hidden rows or columns)
' ' Inputs: rng1, rng2 [Range] intersecting ranges that form single row or column
' '         rng2 is optional if rng1 is desired row or column range
' '
' ' Modified 1/6/22 - Fix bug: search needs to be limited by rng1 --not entire row/col
' '
' Function BuildMultiCellRange(rng1, Optional rng2) As Range
'     Dim w As Variant, rng As Range, xlDirection, IsEntireRowCol As Boolean
'     If rng1 Is Nothing Then Exit Function
'     Set rng = rng1
'     If Not IsMissing(rng2) Then Set rng = Intersect(rng1, rng2)
    
'     'Set search direction for row or column range
    
'     If rng.Rows.Count = 1 Then
'         xlDirection = xlToRight
'         If rng(rng.Cells.Count).Column = rng.Parent.Columns.Count Then IsEntireRowCol = True
'     ElseIf rng.Columns.Count = 1 Then
'         xlDirection = xlDown
'         If rng(rng.Cells.Count).Row = rng.Parent.Rows.Count Then IsEntireRowCol = True
'     Else
'         'Error condition - return Nothing
'         Exit Function
'     End If
    
'     'If rng is col or row (entire) range, restrict search based on last populated cell
'     If IsEntireRowCol Then Set rng = Range(rng(1), rngLastPopCell(rng, xlDirection))
    
'     For Each w In rng
'         If Len(w) > 0 Then
'             If BuildMultiCellRange Is Nothing Then
'                 Set BuildMultiCellRange = w
'             Else
'                 Set BuildMultiCellRange = Union(BuildMultiCellRange, w)
'             End If
'         End If
'     Next w
' End Function
' '------------------------------------------------------------------------------------------------
' ' Add comment to a cell and deletes previous
' '
' Sub AddComment(rngCell, sTxt)
'     With rngCell
'         If Not .Comment Is Nothing Then .Comment.Delete
'         .AddComment
'         .Comment.Visible = False
'         .Comment.Text Text:=sTxt
'     End With
' End Sub

' '-----------------------------------------------------------------------------
' 'Purpose:   Set a table row range (multirange) based on array of key values
' '
' ' Modified 1/17/22 to avoid unhiding hidden columns
' '
' Sub ColWidthAutofit(rngHeader, Optional iMaxWidth = 80)
'     Dim c As Variant
'     For Each c In rngHeader
'         With c.EntireColumn
'             If Not .Hidden = True Then
'                 .ColumnWidth = 240
'                 .AutoFit
'                 If c.ColumnWidth > iMaxWidth Then c.ColumnWidth = iMaxWidth
'                 .ColumnWidth = .ColumnWidth + 2
'             End If
'         End With
'     Next c
' End Sub

' '-----------------------------------------------------------------------------------------------
' ' Range Find that works with hidden cells and cells in column or row outline
' ' https://www.mrexcel.com/board/threads/vba-cannot-find-in-if-cells-are-hidden-even- /x/
' ' if-xlformulas-is-used.518661/
' '
' ' JDL 5/6/20  Modified 10/12/20 JDL - ByVal to work with class attributes)
' '          Modified 5/24/21 JDL - Return rng.cells(i) instead of rng(i) for within-co; search
' '
' Function FindInRange(ByVal rng As Range, ByVal val) As Range
'   Dim i As Integer, q As String
'   Set FindInRange = Nothing

'   If VarType(val) = vbString Then q = """"
'   On Error Resume Next
'   i = Evaluate("MATCH(" & q & val & q & "," & rng.Address(External:=True) & ",0)")
'   If i > 0 Then Set FindInRange = rng.Cells(i)
' End Function
' '-----------------------------------------------------------------------------------------------
' ' Construct name string in RC format for Excel name creation
' '
' ' JDL 5/7/20 - Modified 9/15/20 to add Multiple Rows and Columns; 1/7/22 reformatted
' '
' Function MakeRefNameString(sht, Optional irow1 = 0, Optional irow2 = 0, Optional icol1 = 0, _
'     Optional icol2 = 0) As String
'     MakeRefNameString = "='" & sht & "'!"
    
'     'Entire row
'     If irow1 > 0 And irow2 = 0 Then
'         MakeRefNameString = MakeRefNameString & "R" & irow1 & ":R" & irow1
    
'     'Entire column
'     ElseIf icol1 > 0 And icol2 = 0 Then
'         MakeRefNameString = MakeRefNameString & "C" & icol1 & ":C" & icol1
    
'     'Multiple rows
'     ElseIf irow1 > 0 And irow2 > 0 And icol1 = 0 And icol2 = 0 Then
'         MakeRefNameString = MakeRefNameString & "R" & irow1 & ":R" & irow2
        
'     'Multiple columns
'     ElseIf irow1 = 0 And irow2 = 0 And icol1 > 0 And icol2 > 0 Then
'         MakeRefNameString = MakeRefNameString & "C" & icol1 & ":C" & icol2
    
'     'Range
'     ElseIf icol1 > 0 And icol2 > 0 And irow1 > 0 And irow2 > 0 Then
        
'         'Single cell
'         If irow1 = irow2 And icol1 = icol2 Then
'             MakeRefNameString = MakeRefNameString & "R" & irow1 & "C" & icol1
            
'         'Block Range
'         Else
'             MakeRefNameString = MakeRefNameString & "R" & irow1 & "C" & icol1 & ":R" _
'                 & irow2 & "C" & icol2
'         End If
'     End If
' End Function
' '-----------------------------------------------------------------------------------------------
' ' Create a valid Excel name from sString
' '
' Function xlName(str) As String
'     Dim i As Integer, s As String
        
'     'Initialize the result string and check each character in str (condense/skip spaces)
'     xlName = ""
'     For i = 1 To Len(str)
'         s = Mid(str, i, 1)
                
'         'If it's an invalid character but not a space, replace it with an underscore
'         If InStr(sXLChars, LCase(s)) < 1 And s <> " " Then
'             xlName = xlName & "_"
            
'         'If it's a valid character, simply add it to the XLName
'         ElseIf s <> " " Then
'             xlName = xlName & s
'         End If
'     Next i
    
'     'If necessary, add  underscore prefix if str starts with a number or is column name
'     If InStr(sXLFirstChars, LCase(Left(xlName, 1))) < 1 Or _
'         IsExcelCellRef(xlName) Then xlName = "_" & xlName
' End Function
' '-----------------------------------------------------------------------------------------------
' ' Return value of specified setting; returns nothing if not found
' '
' ' JDL 4/2/20 Modified: 5/19/21 Add exit if no shtSettings
' '
' Function ReadSetting(wkbk, ByVal sName As String) As Variant
'     Dim c As Range
'     If Not SheetExists(wkbk, shtSettings) Then Exit Function
'     Set c = wkbk.Sheets(shtSettings).Columns(1).Find(sName, LookAt:=xlWhole)
'     If c Is Nothing Then Exit Function
'     ReadSetting = c.Offset(0, 1)
' End Function
' '-----------------------------------------------------------------------------------------------
' ' Update Setting value; create new if not found
' '
' ' JDL 4/2/20  Modified: 4/14/20 Addsheet; 1/10/22 docstring
' '
' Sub UpdateSetting(wkbk, ByVal sName As String, ByVal val As Variant)
'     Dim c As Range
'     If Not SheetExists(wkbk, shtSettings) Then _
'         AddSheet wkbk, shtSettings, wkbk.Sheet(wkbk.Sheets.Count)

'     'Find the existing setting row or add a new setting if  not found
'     With wkbk.Sheets(shtSettings)
'         Set c = .Columns(1).Find(sName, LookAt:=xlWhole)
'         If c Is Nothing Then Set c = .Cells(.Rows.Count, 1).End(xlUp).Offset(1, 0)
'         c = sName
'         c.Offset(0, 1) = val
'         '.Visible = xlVeryHidden
'     End With
' End Sub
' '-----------------------------------------------------------------------------------------------
' ' Performs Sort Step - modified to work with non-homed tables
' ' From modExcelStepsRefresh
' '
' ' Modified JDL 7/8/21
' '
' Sub StepSortBy(sSortBy, rngheaders, rngData)
'     Dim sArySortBy() As String, i As Integer, w As Range

'     sArySortBy = Split(sSortBy, ",")

'     With rngheaders.Parent.Sort.SortFields
'         .Clear
'         For i = LBound(sArySortBy) To UBound(sArySortBy)
'             Set w = rngheaders.Find(sArySortBy(i), LookAt:=xlWhole)

'             'Exit if specified sort field not found in the table
'             If w Is Nothing Then Exit Sub

'             .Add Key:=w, SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortTextAsNumbers
'         Next i
'     End With
'     With rngheaders.Parent.Sort
'         .SetRange Intersect(rngheaders.EntireColumn, rngData)
'         .Header = xlNo
'         .MatchCase = False
'         .Orientation = xlTopToBottom
'         .Apply
'         .SortFields.Clear
'     End With
' End Sub
' '------------------------------------------------------------------------------------------------
' ' True if specified sheet exists
' '
' Function SheetExists(ByVal wkbk As Workbook, ByVal sht As String) As Boolean
'     Dim w As Variant
'     SheetExists = True
'     For Each w In wkbk.Sheets
'         If w.Name = sht Then Exit Function
'     Next w
'     SheetExists = False
' End Function
' '-----------------------------------------------------------------------------------------------
' ' True (and modify ByRef sht arg) if sheet exists but name has case difference
' '   This function deals with the VBA/Excel oddity that Worksheet names are case sensitive for
' '   testing their existence, but wkbk can't have two worksheets with same names except for case.
' '   The function tests for case differences (usually appropriate after first testing that
' '   SheetExists = False) and returns modified sht argument to identically match Worksheet name
' '   --Boolean can be trigger user warning and is test of whether to add a missing Worksheet
' '
' '   JDL 1/20/22
' '
' Function IsShtCaseErr(wkbk, ByRef sht) As Boolean
'     Dim w As Variant
'     For Each w In wkbk.Sheets
'         If (LCase(w.Name) = LCase(sht)) And (w.Name <> sht) Then
'             IsShtCaseErr = True
'             sht = w.Name
'         End If
'     Next w
' End Function
' '-----------------------------------------------------------------------------------------------
' ' Add a new worksheet if doesn't already exist in wkbk
' '
' ' Modified: 12/3/21 Optional shtAfter
' '
' Sub AddSheet(wkbk, shtNew, Optional shtAfter)
'     Dim wkshtNew As Worksheet

'     If wkbk Is Nothing Then Exit Sub
'     If SheetExists(wkbk, shtNew) Then Exit Sub
    
'     'Set shtAfter location if not specified
'     If IsMissing(shtAfter) Then shtAfter = wkbk.Sheets(wkbk.Sheets.Count).Name
'     If Not SheetExists(wkbk, shtAfter) Then shtAfter = wkbk.Sheets(wkbk.Sheets.Count).Name
'     Set wkshtNew = wkbk.Sheets.Add(After:=wkbk.Sheets(shtAfter))
'     wkshtNew.Name = shtNew
' End Sub
' '-----------------------------------------------------------------------------------------------
' ' Delete specified sheet if it exists
' '
' Sub DeleteSheet(wkbk, sht)
'     Application.DisplayAlerts = False
'     If SheetExists(wkbk, sht) Then wkbk.Sheets(sht).Delete
' End Sub
' '-----------------------------------------------------------------------------------------------
' ' Create an Excel name, sName, using sNameString as RefersTo
' '
' Sub MakeXLName(wkbk, sName, sNameString)
'     DeleteXLName wkbk, sName
'     wkbk.Names.Add Name:=sName, RefersToR1C1:=sNameString
' End Sub
' '-----------------------------------------------------------------------------------------------
' ' Delete an Excel name from a workbook
' '
' Sub DeleteXLName(wkbk, sName)
' Dim w As Variant
'     For Each w In wkbk.Names
'         If w.Name = sName Then
'             wkbk.Names(w.Name).Delete
'             Exit Sub
'         End If
'     Next w
' End Sub
' '-----------------------------------------------------------------------------------------------
' ' Delete all workbook names
' ' The .Visible property is used to skip hidden _xlfn.SINGLE name created by dynamic array
' ' glitch from circa 2020 Excel update. See: https://stackoverflow.com/questions/59121799
' '
' ' JDL 11/22/21
' '
' Sub DeleteAllWorkbookNames(wkbk)
'     Dim w As Variant
'     For Each w In wkbk.Names
'         If w.Visible Then w.Delete
'     Next w
' End Sub
' '-----------------------------------------------------------------------------------------------
' ' True if named range exists in the workbook
' '
' Function NameExists(wkbk, sName) As Boolean
' Dim w As Variant
'     NameExists = False
'     For Each w In wkbk.Names
'         If UCase(w.Name) = UCase(sName) Then
'             NameExists = True
'             Exit Function
'         End If
'     Next w
' End Function
' '-----------------------------------------------------------------------------------------------
' ' Determine whether specified value is in array
' '
' ' JDL 1/7/21 - replace verbose version Modified: 3/3/21 handle non-array
' '
' Public Function IsInAry(ary As Variant, val As Variant) As Boolean

'     'Exit if empty array)
'     If UBound(ary) = -1 Then Exit Function

'     IsInAry = Not IsError(Application.Match(val, ary, 0))
' End Function
' '-----------------------------------------------------------------------------------------------
' ' Set a table row range (multirange) based on array of key values
' ' Notes:     Mimics Pandas .loc functionality with multiindex
' '            Suitable for small tables - Works well up to 500 found items in 1000 row table
' '            Bogs down with large n found items (e.g. 2500 found rows in 5000 row table)
' ' Inputs:    tbl [table Class instance]
' '            aryKeyColRanges [array of Ranges] table Class column ranges for key columns
' '            aryKeyValues [array of values (variant)] array of key column values
' ' Validated in Unique Vals_1220.xlsm; validated 4-key 1/27/22 JDL val_Scenario Model.xlsm
' '
' 'Created:   12/17/20 JDL      Modified: 2/17/21
' '
' '
' Function rngMultiKey(tbl, aryKeyColRanges, aryKeyValues) As Range
'     Dim i As Integer, rngCurrent As Range, rngSearch As Range, rngFound As Range

'     'Progressively search/subset across the key columns -- start with entire table
'     Set rngCurrent = tbl.rngrows
'     For i = LBound(aryKeyColRanges) To UBound(aryKeyColRanges)

'         'Stop if previous search returned Nothing
'         If Not rngCurrent Is Nothing Then
'             Set rngSearch = Intersect(rngCurrent, aryKeyColRanges(i))
'             Set rngCurrent = Nothing
'             On Error Resume Next 'in case nothing
'             Set rngCurrent = FindAll(rngSearch, aryKeyValues(i)).EntireRow
'             On Error GoTo 0
'         End If
'     Next i
'     If Not rngCurrent Is Nothing Then Set rngMultiKey = rngCurrent.EntireRow
' End Function




