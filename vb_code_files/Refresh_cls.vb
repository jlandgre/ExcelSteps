'Refresh_cls.vb
'Version 5/6/25
Option Explicit

' This Class handles refreshing a rows/columns table. In ExcelSteps, .RefreshRC() method
' is called from the OK button in frmRefresh UserForm
Public IsMdl As Boolean
Public IsTbl As Boolean
Public wkbk As Workbook 'Model workbook
Public sht As String
Public wkbkS As Workbook 'Workbook containing Steps recipe (default same as model)
Public shtS As String 'Recipe sheet name in wkbkS (default ExcelSteps aka shtSteps

'Internal, loop variables; Usage contingent on mdl or tbl
Public IsCalc As Boolean 'applicable for Scenario Model
Public IsNamePrefix As Boolean

'tblRowsCols attributes (See its header docstring for descriptions)
Public IsReplace As Boolean
Public IsTblFormat As Boolean

'ExcelSteps range of instructions for specified tbl or mdl
Public rngSteps As Range
'------------------------------------------------------------------------------------------------
' Initialize the Class for refreshing mdlScenario
'
' JDL 10/18/24 split original .Init to .InitMdl and .InitTbl
'
Public Function InitMdl(refr, wkbk) As Boolean
    SetErrs InitMdl: If errs.IsHandle Then On Error GoTo ErrorExit
     
    With refr
    
        'Workbook to refresh
        Set .wkbk = wkbk

        'Workbook containing Steps sheets (same as R wkbk but could vary in future)
        Set .wkbkS = wkbk
        .shtS = shtSteps
    End With
    
    'Ensure only one sheet is selected (12/6/22)
    If ActiveWindow.SelectedSheets.Count > 1 Then ActiveWindow.SelectedSheets(1).Select
    Exit Function
    
ErrorExit:
    errs.RecordErr "InitMdl", InitMdl
End Function
'------------------------------------------------------------------------------------------------
' Initialize the Class for refreshing tblRowsCols
'
' Inputs:   shtR [String] sheet name to refresh
'         wkbkR [Workbook] workbook containing table to refresh
'         IsReplace [Boolean] True if refresh replaces original version of table
'         IsTblFormat [Boolean] True if refresh table format
'
' Modified 10/18/24 Refactor to handle custom/non-default tables
'
Public Function InitTbl(refr, wkbk, IsReplace, IsTblFormat) As Boolean
    SetErrs InitTbl: If errs.IsHandle Then On Error GoTo ErrorExit
     
    With refr
    
        'Workbook to refresh
        Set .wkbk = wkbk

        'Workbook containing Steps sheets (same as R wkbk but could vary in future)
        Set .wkbkS = wkbk
        .shtS = shtSteps
                
        .IsReplace = True
        If Not IsMissing(IsReplace) Then .IsReplace = IsReplace
                
        .IsTblFormat = False
        If Not IsMissing(IsTblFormat) Then .IsTblFormat = IsTblFormat
    End With
    
    'Ensure only one sheet is selected (12/6/22)
    If ActiveWindow.SelectedSheets.Count > 1 Then ActiveWindow.SelectedSheets(1).Select
    Exit Function
    
ErrorExit:
    errs.RecordErr "InitTbl", InitTbl
End Function
'------------------------------------------------------------------------------------------------
' Refresh a rows/columns table
' JDL 12/15/22; Mod 3/6/23 add exit if tbl.rngTable Nothing (e.g. blank sheet)
'                   6/9/23 JDL cleanup; 9/30/24 tblSteps instead of tblS
'                   10/18/24 generalize to refresh custom tables
'
Public Function RefreshRC(refr, tblSteps, Optional sht, Optional TblName, Optional Defn, _
    Optional rcHome, Optional nRows, Optional nCols, Optional iOffsetKeyCol, Optional iOffsetHeader, _
    Optional IsSetAryCols, Optional IsSetColRngs, Optional IsSetTblNames, Optional IsSetColNames, _
    Optional IsNamePrefix, Optional IsPrefixSht, Optional NamePrefix) As Boolean
            
    SetErrs RefreshRC: If errs.IsHandle Then On Error GoTo ErrorExit
    Dim tblT As New tblRowsCols, tbl As New tblRowsCols, shtT As String, NamePrefixT As String
    
    With refr
    
        'Prep the ExcelSteps sheet
        If Not .PrepExcelStepsSht(refr, tblSteps) Then GoTo ErrorExit
        
        'Provision the original table; Exit if blank; otherwise save Zoom setting for post-replace
        If Not tbl.Provision(tbl, .wkbk, .IsTblFormat, sht, TblName, Defn, rcHome, nRows, _
            nCols, iOffsetKeyCol, iOffsetHeader, IsSetAryCols, IsSetColRngs, IsSetTblNames, _
            IsSetColNames, IsNamePrefix, IsPrefixSht, NamePrefix) Then GoTo ErrorExit
                                     
        'Exit if blank
        If tbl.rngTable Is Nothing Then Exit Function
        tbl.ShtActivateAndSaveZoom tbl

        'Create copy of table
        If Not .PrepTempSheet(refr, tbl, tblT) Then GoTo ErrorExit
        
        'Name of transformed sheet - Defn/TblName will trigger provision as custom table anyway
        shtT = tblT.sht
        NamePrefixT = tbl.NamePrefix

        'Provision tblT using same inputs as tbl
        If Not tblT.Provision(tblT, .wkbk, .IsTblFormat, shtT, TblName, Defn, rcHome, nRows, _
            nCols, iOffsetKeyCol, iOffsetHeader, IsSetAryCols, IsSetColRngs, IsSetTblNames, _
            IsSetColNames, IsNamePrefix, IsPrefixSht, NamePrefixT) Then GoTo ErrorExit
    End With
                    
    With tblT
        'Extend the header to make range robust to inserted column at right edge of table
        Set .rngHeader = Range(.rngHeader, .rngHeader.Cells(.rngHeader.Cells.Count).Offset(0, 1))
        
        'Check for column name redundancy
        If Not FlagRedundant(.rngHeader) Then GoTo ErrorExit
        
        'Refresh the sheet
        If Not RefreshTblFromRecipe(tblT, tblSteps, tbl.sht) Then GoTo ErrorExit
        If Not .rngRows Is Nothing Then .rngRows.EntireRow.AutoFit

        'If initial column inserted (Column A), Reset header and names
        'Note: Won't work if inserted formulas depend on table ranges and have
        'Keep Formulas FALSE; table and table_header range need to be kept correct
        'as cols are inserted to enable recalc before pasting values into inserted
        'cols. Robust solution is to use tblT class so it has access to cellHome
        'address to adjust if Step_Insert moves (9/15/21)
        If Not .rngHeader Is Nothing Then
            If IsRngDeleted(.cellHome) Then Set .cellHome = .wkbk.Sheets(.sht).Cells(2, 1)
            If .cellHome.Column > 1 Then Set .cellHome = .wkbk.Sheets(.sht).Cells(2, 1)
            If Not .SetDimensions(tblT) Then GoTo ErrorExit
        End If

        'Replace previous .sht
        If refr.IsReplace Then
            DeleteSheet .wkbk, tbl.sht
            .wkbk.Sheets(.sht).Name = tbl.sht
        End If
        
        'Restore original Zoom setting
        .RestoreZoom tblT
        .wksht.Cells(1, 1).Select
    End With
    Exit Function

ErrorExit:
    errs.RecordErr "RefreshRC", RefreshRC
End Function
'------------------------------------------------------------------------------------------------
' Refresh Sheet from ExcelSteps recipe
' JDL 12/16/22  Modified 6/9/23 JDL cleanup; 9/30/24 tblSteps instead of tblSteps
'                                            10/18/24 modify for refactored refr
'                                            10/24/24 change recalc to optimize performance
'
Function RefreshTblFromRecipe(tblT, tblSteps, sht) As Boolean
    SetErrs RefreshTblFromRecipe: If errs.IsHandle Then On Error GoTo ErrorExit

    Dim Step As RecipeStep, rngRecipeActions, r As Range
    
    'Clear and reset the sheet's outline
    If Not ClearAndResetOutline(tblT.wksht) Then GoTo ErrorExit
        
    'Read and iterate through the steps
    Set rngRecipeActions = FindAll(tblSteps.colrngSht, sht)
    If rngRecipeActions Is Nothing Then Exit Function
    
    For Each r In rngRecipeActions
        Set tblSteps.rowCur = r.EntireRow
        Set Step = New RecipeStep
        If Not Step.Read(Step, tblSteps) Then GoTo ErrorExit
        If Not Step.RunActions(Step, tblT, tblSteps) Then GoTo ErrorExit
    Next r
    
    'Recalculate the application and set formatting/values for col_Insert columns
    If Not SetInsertFormatsAndVals(tblT, tblSteps, rngRecipeActions) Then GoTo ErrorExit
    Exit Function
    
ErrorExit:
    errs.RecordErr "RefreshTblFromRecipe", RefreshTblFromRecipe
End Function
'------------------------------------------------------------------------------------------------
' Recalculate and set formatting and values for all Insert step columns ()
' JDL 1/10/25 Fix bug in case colInsert was deleted by previous step
' Modified 5/6/25
'
Function SetInsertFormatsAndVals(tblT, tblSteps, rngRecipeActions) As Boolean
    SetErrs SetInsertFormatsAndVals: If errs.IsHandle Then On Error GoTo ErrorExit

    Dim Step As RecipeStep, r As Range, act As Variant, colInsert As Range
            
    'Recalculate to update data values based on inserted formulas
    Application.Calculate
    
    For Each r In rngRecipeActions
        Set tblSteps.rowCur = r.EntireRow
        
        Set Step = New RecipeStep
        If Not Step.Read(Step, tblSteps) Then GoTo ErrorExit
        
        'For col_Insert columns, either format live calculation or copy/paste values
        If Step.sType = sAInsert Then
            Set colInsert = tblT.rngHeader.Find(Step.sColName, lookat:=xlWhole)
            
            'Conditional in case colInsert was deleted by previous step (1/10/25)
            If Not colInsert Is Nothing Then
                Set colInsert = colInsert.EntireColumn

                With Intersect(colInsert, tblT.rngRows)
                    If Step.IsKeepFormula Then
                        .Style = "Calculation"
                           
                        '5/6/25 change color to black but keep bold and background from Calculation
                        .Font.ColorIndex = xlAutomatic
                
                    Else
                        .Copy
                        .PasteSpecial Paste:=xlPasteValues
                        Application.CutCopyMode = False
                    End If
                End With
            End If

        End If
    Next r
    Exit Function
    
ErrorExit:
    errs.RecordErr "SetInsertFormatsAndVals", SetInsertFormatsAndVals
End Function
'-----------------------------------------------------------------------------------------------
' Prep Temporary copy of rows/columns sheet
'
' Created:   7/30/21 JDL      Modified: 10/11/24 update .Provision NamePrefix arg to current
'
' Inputs: tbl, tblT [Class Instance] description of original and copied tables
'
Function PrepTempSheet(refr, tbl, tblT) As Boolean
    SetErrs PrepTempSheet: If errs.IsHandle Then On Error GoTo ErrorExit
    Dim r As Range
    
    With tbl
    
        'Construct transformed sheet name with deference to Excel length limit
        tblT.sht = Left(.sht, 30 - Len(sRefreshSuffix)) & sRefreshSuffix
        
        'Record original table's zoom setting to restore later
        tblT.ZoomSetting = tbl.ZoomSetting

        'Add temporary worksheet
        DeleteSheet .wkbk, tblT.sht
        AddSheet .wkbk, tblT.sht, .sht
        Set r = .wkbk.Sheets(tblT.sht).Range(.rngHeader.Cells(1).Address)
    
        'Unfilter source table so copy/paste gets everything
        If Not .wksht.AutoFilter Is Nothing Then .wksht.AutoFilter.Range.AutoFilter
    
        'Transfer original table to new sheet
        .rngTable.Copy
        r.PasteSpecial Paste:=xlPasteFormulas
        r.PasteSpecial Paste:=xlPasteComments
        r.PasteSpecial Paste:=xlPasteFormats
    End With
    Exit Function
    
ErrorExit:
    errs.RecordErr "PrepTempSheet", PrepTempSheet
End Function
'------------------------------------------------------------------------------------------------
' Format ExcelSteps sheet Procedure
'
' Modified: 7/14/23 rearrange initialize with If/Endif
'           7/18/23 add IsReformat to force recreate header/formatting from cleared sheet
'           7/27/23 customize rngRows to be from first to last populated rows
'           9/30/24 tblSteps instead of tblS
'           10/18/24 refactor for refr.InitTbl and RefreshAPI and cleanup
'
Function PrepExcelStepsSht(refr, tblSteps, Optional IsReformat = False) As Boolean
    SetErrs PrepExcelStepsSht: If errs.IsHandle Then On Error GoTo ErrorExit
    With refr
    
        'If no Steps sheet, add and initialize before Provision
        If Not SheetExists(.wkbkS, .shtS) Then
            AddSheet .wkbkS, .shtS, .wkbkS.Sheets(wkbkS.Sheets.Count).Name
            IsReformat = True
        
        'Check for existence of header values
        ElseIf .wkbkS.Sheets(.shtS).Cells(1, 1).value <> "Sheet" Then
            IsReformat = True
        End If
        
        If IsReformat Then
            If Not .SetStepsShtHeader(refr) Then GoTo ErrorExit
            If Not .FormatStepsShtWidths(refr) Then GoTo ErrorExit
            If Not .FormatStepsShtColumns(refr) Then GoTo ErrorExit
            If Not tblSteps.Provision(tblSteps, .wkbkS, False, sht:=.shtS, IsSetColRngs:=True) Then GoTo ErrorExit
            If Not .FormatStepsBorders(tblSteps) Then GoTo ErrorExit
            If Not .AddStepsShtDropdowns(tblSteps) Then GoTo ErrorExit
            If Not .StepsShtMiscFormat(tblSteps) Then GoTo ErrorExit
        Else
            If Not tblSteps.Provision(tblSteps, .wkbkS, False, sht:=.shtS, IsSetColRngs:=True) Then GoTo ErrorExit
            
            'Customize rngRows to include last populated row since  may be blank rows
            With tblSteps
                If Not .rngRows Is Nothing Then
                    Set .rngRows = Range(.rngRows.Rows(1), _
                                .colrngCol.Cells(.wksht.Rows.Count, 1).End(xlUp))
                    .nRows = .rngRows.Rows.Count
                End If
            End With
        End If
    End With
    Exit Function
    
ErrorExit:
    errs.RecordErr "PrepExcelStepsSht", PrepExcelStepsSht
End Function
'------------------------------------------------------------------------------------------------
' Set ExcelSteps sheet headers and header format
' JDL 12/15/22   Modified 10/18/24 JDL cleanup
'
Public Function SetStepsShtHeader(refr) As Boolean

    SetErrs SetStepsShtHeader: If errs.IsHandle Then On Error GoTo ErrorExit
    Dim ary As Variant, lstHeaders As String, rngHeader As Range

    'Create array of header strings and populate into first row
    lstHeaders = "Sheet,Column,Step,Formula/List Name/Sort-by," _
        & "After End or Rename Column,Keep Formulas,Comment,Number Format,Width"
    ary = Split(lstHeaders, ",")
    With refr.wkbk.Sheets(refr.shtS)
       Set rngHeader = Range(.Cells(1, 1), .Cells(1, UBound(ary) + 1))
    End With
    rngHeader = ary
      
    'Format the header
    With rngHeader
        .Style = "Accent2"
        .EntireColumn.Font.Size = 9
        .WrapText = True
    End With
    Exit Function
    
ErrorExit:
    errs.RecordErr "SetStepsShtHeader", SetStepsShtHeader
End Function
'------------------------------------------------------------------------------------------------
' Format ExcelSteps Sheet Column Widths
' JDL 12/15/22   Modified 7/18/23 fix bug by change to With refr.wkbk instead of tblSteps.wksht
'
Public Function FormatStepsShtWidths(refr) As Boolean

    SetErrs FormatStepsShtWidths: If errs.IsHandle Then On Error GoTo ErrorExit
    Dim iWidth As Variant, i As Integer

    'Set column widths and column formats
    With refr.wkbk.Sheets(refr.shtS)
        i = 1
        For Each iWidth In Array(12, 12, 12, 32, 12, 7, 12, 7, 7)
            .Columns(i).EntireColumn.ColumnWidth = iWidth
            i = i + 1
        Next iWidth
    End With
    Exit Function
    
ErrorExit:
    errs.RecordErr "FormatStepsShtWidths", FormatStepsShtWidths
End Function
'------------------------------------------------------------------------------------------------
' Format ExcelSteps Sheet Columns
' JDL 12/15/22 Modified 7/18/23 fix bug; change to With refr.wkbk. instead of tblSteps.rngHeader
'
Public Function FormatStepsShtColumns(refr) As Boolean
    SetErrs FormatStepsShtColumns: If errs.IsHandle Then On Error GoTo ErrorExit
    Dim r As Range

    With refr.wkbk.Sheets(refr.shtS).Rows(1)
        Set r = .Find("Number Format", lookat:=xlWhole)
        r.EntireColumn.NumberFormat = "@"
        
        Set r = .Find("Formula/List Name/Sort-by", lookat:=xlWhole)
        r.EntireColumn.NumberFormat = "@"
        r.EntireColumn.WrapText = True
        
        Set r = .Find("Comment", lookat:=xlWhole)
        r.EntireColumn.WrapText = True
    End With
    Exit Function
    
ErrorExit:
    errs.RecordErr "FormatStepsShtColumns", FormatStepsShtColumns
End Function
'------------------------------------------------------------------------------------------------
' Format ExcelSteps Sheet Borders
' JDL 12/15/22   Modified 6/9/23 JDL cleanup; 10/18/24; eliminate refr argument
'
Public Function FormatStepsBorders(tblSteps) As Boolean

    SetErrs FormatStepsBorders: If errs.IsHandle Then On Error GoTo ErrorExit
    Dim r As Range

    'Set rngRows to extend at 20 rows beyond data
    With tblSteps.wksht
        Set r = .Cells(.Rows.Count, 1).End(xlUp)
        If r.Row = 1 Then
            Set tblSteps.rngRows = Range(.Rows(2), .Rows(21))
        Else
            Set tblSteps.rngRows = Range(.Rows(2), r.EntireRow.Offset(20, 0))
        End If
    End With
    
    'Set Borders and font size on the data range
    With tblSteps
        SetBorders .wksht.Cells, xlNone, True
        SetBorders Intersect(.rngRows, .rngHeader.EntireColumn), xlContinuous, True
        Intersect(.rngHeader.EntireColumn, .rngRows).Font.Size = 9
    End With
    Exit Function
    
ErrorExit:
    errs.RecordErr "FormatStepsBorders", FormatStepsBorders
End Function
'------------------------------------------------------------------------------------------------
' Add dropdown list to steps column and Clear previous error messages
' JDL 12/15/22   Modified 6/9/23 JDL cleanup; 10/18/24 eliminate refr argument
'
Public Function AddStepsShtDropdowns(tblSteps) As Boolean
    SetErrs AddStepsShtDropdowns: If errs.IsHandle Then On Error GoTo ErrorExit
    Dim r As Range
    
    With tblSteps
        Set r = .rngHeader.Find("Step", lookat:=xlWhole)
        AddValidationList Intersect(.rngRows, r.EntireColumn), sStepList
        r.EntireColumn.ClearComments
    End With
    Exit Function
    
ErrorExit:
    errs.RecordErr "AddStepsShtDropdowns", AddStepsShtDropdowns
End Function
'------------------------------------------------------------------------------------------------
' Miscellaneous cleanup of ExcelSteps sheet
' JDL 12/15/22   Modified 6/9/23 JDL cleanup; 10/18/24 eliminate refr argument
'
Public Function StepsShtMiscFormat(tblSteps) As Boolean
    SetErrs StepsShtMiscFormat: If errs.IsHandle Then On Error GoTo ErrorExit
    Dim rngCol As Range, c As Variant, Step As New RecipeStep
    
    With tblSteps
    
        'Ensure formula cells display formulas as text - not calculated values
        Set rngCol = .rngHeader.Find("Formula/List Name/Sort-by", lookat:=xlWhole)
        Set rngCol = rngCol.EntireColumn
        For Each c In Intersect(.rngRows, rngCol)
            c.Cells(4).value = c.Cells(4).Formula
        Next c
    
        'Freeze Row1 (RecipeStep method)
        If Not Step.FreezeRow1(Step, tblSteps) Then GoTo ErrorExit
    End With
    Exit Function
    
ErrorExit:
    errs.RecordErr "StepsShtMiscFormat", StepsShtMiscFormat
End Function
