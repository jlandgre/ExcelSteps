'ExcelSteps_RecipeStep_cls.vb
'Version 12/2/25 Minor updates
Option Explicit
'        7/8/25 update of .Find to FindInRange and fix bug with .rngHeader ref
'This class describes an Excel Step as read from the ExcelSteps sheet
Public sSheet As String
Public sColName As String
Public sType As String
Public sFormula As String
Public sCol2 As String
Public sDeleteVal As String
Public IsKeepFormula As Boolean
Public sComment As String
Public NumFmt As String
Public sSortBy As String
Public sLstRngName As String
Public Width As Variant
Public Actions As Collection
'------------------------------------------------------------------------------------------------
'Read RecipeStep Attributes and Create Collection of Actions
'JDL 12/16/22   Modified 6/9/23 JDL cleanup
'
Function Read(Step, tblSteps) As Boolean
    SetErrs Read: If errs.IsHandle Then On Error GoTo ErrorExit
    
    With Step
    
        'Overall parameters
        .sSheet = Intersect(tblSteps.rowCur, tblSteps.colrngSht)
        .sColName = Intersect(tblSteps.rowCur, tblSteps.colrngCol)
        .sType = Intersect(tblSteps.rowCur, tblSteps.colrngStep)

        'Populate a collection with actions involved with this RecipeStep
        Set .Actions = New Collection
        If Not CreateActionsColl(Step, tblSteps) Then GoTo ErrorExit
    
        'Insert is specified by 3 parameters
        If .sType = sAInsert Then
            .sFormula = Intersect(tblSteps.rowCur, tblSteps.colrngStrInput).Formula
            .IsKeepFormula = Intersect(tblSteps.rowCur, tblSteps.colrngKeep)
            .sCol2 = Intersect(tblSteps.rowCur, tblSteps.colrngCol2)
        End If
        
        'Delete rows with specified value
        If .sType = sADelRowsWithVal Then
            .sDeleteVal = Intersect(tblSteps.rowCur, tblSteps.colrngStrInput).Formula
        End If
        
        'Actions besides Insert that require second column name
        If .sType = sAGroup Then .sCol2 = Intersect(tblSteps.rowCur, tblSteps.colrngCol2)
        If .sType = sARename Then .sCol2 = Intersect(tblSteps.rowCur, tblSteps.colrngCol2)
            
        'Actions that require an additional String or list input
        If .sType = sADropdown Then .sLstRngName = Intersect(tblSteps.rowCur, tblSteps.colrngStrInput)
        If .sType = sASort Then .sSortBy = Intersect(tblSteps.rowCur, tblSteps.colrngStrInput)
        
        'Formatting actions - only populate if they are specified
        If IsKeyInColl(.Actions, sAComment) Then .sComment = Intersect(tblSteps.rowCur, tblSteps.colrngComment)
        If IsKeyInColl(.Actions, sANumFmt) Then .NumFmt = Intersect(tblSteps.rowCur, tblSteps.colrngNumFmt)
        If IsKeyInColl(.Actions, sAWidth) Then .Width = Intersect(tblSteps.rowCur, tblSteps.colrngWidth)
    End With
    Exit Function

ErrorExit:
    errs.RecordErr "Read", Read
End Function
'------------------------------------------------------------------------------------------------
' Populate Actions Collection with either single or multiple actions
' JDL 12/16/22   Modified 6/9/23 JDL cleanup
'
Function CreateActionsColl(Step, tblSteps) As Boolean

    SetErrs CreateActionsColl: If errs.IsHandle Then On Error GoTo ErrorExit
    Dim act As Variant, aryColRngs As Variant, idx As Integer
    
    aryColRngs = Array(tblSteps.colrngComment, tblSteps.colrngNumFmt, tblSteps.colrngWidth)
    With Step
    
        'Multi-action Step Types
        If .sType = sAInsert Or .sType = sAFormat Then
            If .sType = sAInsert Then .Actions.Add sAInsert, sAInsert
            For Each act In Array(sAComment, sANumFmt, sAWidth)
                If Len(Intersect(aryColRngs(idx), tblSteps.rowCur)) > 0 Then .Actions.Add act, act
                idx = idx + 1
            Next
            
        'Single action Step Types
        Else
            .Actions.Add .sType
        End If
    End With
    Exit Function

ErrorExit:
    errs.RecordErr "CreateActionsColl", CreateActionsColl
End Function
'------------------------------------------------------------------------------------------------
' Run Step's Actions
' JDL 12/16/22   Modified 6/9/23 JDL cleanup
'
Function RunActions(Step, tbl, tblSteps) As Boolean
    SetErrs RunActions: If errs.IsHandle Then On Error GoTo ErrorExit
    Dim act As Variant

    'Clear previous error message comments
    tblSteps.colrngStep.ClearComments
    
    'Perform each action in Actions Collection
    For Each act In Step.Actions
        If act = sAInsert Then If Not InsertCol(Step, tbl) Then GoTo RecordComment
        If act = sAComment Then If Not HeaderComment(Step, tbl) Then GoTo RecordComment
        If act = sAWidth Then If Not SetWidth(Step, tbl) Then GoTo RecordComment
        If act = sANumFmt Then If Not SetNumFmt(Step, tbl) Then GoTo RecordComment
        If act = sAGroup Then If Not AddGroup(Step, tbl) Then GoTo RecordComment
        If act = sADelete Then If Not Delete(Step, tbl) Then GoTo RecordComment
        If act = sARename Then If Not Rename(Step, tbl) Then GoTo RecordComment
        If act = sAFreezeRow1 Then If Not FreezeRow1(Step, tbl) Then GoTo RecordComment
        If act = sASplitCols Then If Not SplitColumns(Step, tbl) Then GoTo RecordComment
        If act = sADropdown Then If Not AddValLst(Step, tbl) Then GoTo RecordComment
        If act = sASort Then If Not SortBy(Step, tbl) Then GoTo RecordComment
        If act = sACondFmt Then If Not CondFormat(Step, tbl) Then GoTo RecordComment
        
        '1/10/25
        If act = sADelFlagRows Then If Not DeleteFlaggedRows(Step, tbl) Then GoTo RecordComment
        If act = sADelRowsWithVal Then If Not DeleteRowsWithVal(Step, tbl) Then GoTo RecordComment
    Next act
    Exit Function

'Put an error message comment in the recipe's row
RecordComment:
    errs.LookupCommentMsg Intersect(tblSteps.rowCur, tblSteps.colrngStep), "RecipeStep_Comments"
    
    'Reset code to also show a dialog box error message directing to cell comments
    errs.iCodeLocal = 1
    errs.ErrParam = Step.sSheet
    GoTo ErrorExit
    
ErrorExit:
    errs.RecordErr "RunActions", RunActions
End Function
'------------------------------------------------------------------------------------------------
' Delete rows where column contains specified value
'
' JDL 1/10/25;  6/15/25; 7/8/25 Update .Find to FindInRange
'
Function DeleteRowsWithVal(Step, tbl) As Integer

    SetErrs DeleteRowsWithVal: If errs.IsHandle Then On Error GoTo ErrorExit
    Dim r As Range, cellCur As Range, s As String
        
    If tbl.rngRows Is Nothing Then Exit Function
    
    'Exit flag column not specified or if column not found
    With Step
        Set r = FindInRange(tbl.rngHeader, .sColName)
        If errs.IsFail(r Is Nothing, 2) Then GoTo ErrorExit
        
        'Delete value not specified
        If errs.IsFail(Len(.sDeleteVal) < 1, 16) Then GoTo ErrorExit
    End With

    'Initialize cellCur as last cell in flag column
    With tbl.rngRows
        Set cellCur = Intersect(r.EntireColumn, .Rows(.Rows.Count).EntireRow)
    End With
    
    'Loop moving upwards and delete rows flagged with False
    Do While cellCur.Row > 1
            Set cellCur = cellCur.Offset(-1, 0)
        If cellCur.Offset(1, 0) = Step.sDeleteVal Then cellCur.Offset(1, 0).EntireRow.Delete
    Loop
    
    'Deal with special case where all rows were deleted (.rngRows still exists but is undefined)
    On Error Resume Next
    Dim rngAddress As String
    s = tbl.rngRows.Address ' Try to access .Address property
    On Error GoTo 0
    If s = "" Then Set tbl.rngRows = Nothing
    Exit Function
    
ErrorExit:
    DeleteRowsWithVal = False
End Function
'------------------------------------------------------------------------------------------------
' Delete rows flagged with False value in True/False column
'
' JDL 1/10/25; 6/15/25; 7/8/25 Update .Find to FindInRange
'
Function DeleteFlaggedRows(Step, tbl) As Integer

    SetErrs DeleteFlaggedRows: If errs.IsHandle Then On Error GoTo ErrorExit
    Dim r As Range, cellCur As Range, s As String
        
    If tbl.rngRows Is Nothing Then Exit Function
    
    'Exit flag column not specified or if column not found
    With Step
        Set r = FindInRange(tbl.rngHeader, .sColName)
        If errs.IsFail(r Is Nothing, 2) Then GoTo ErrorExit
    End With

    'Initialize cellCur as last cell in flag column
    With tbl.rngRows
        Set cellCur = Intersect(r.EntireColumn, .Rows(.Rows.Count).EntireRow)
    End With
    
    'Loop moving upwards and delete rows flagged with False
    Do While cellCur.Row > 1
            Set cellCur = cellCur.Offset(-1, 0)
        If cellCur.Offset(1, 0) = False Then cellCur.Offset(1, 0).EntireRow.Delete
    Loop
    
    'Deal with special case where all rows flagged and deleted (.rngRows still exists but is undefined)
    On Error Resume Next
    Dim rngAddress As String
    s = tbl.rngRows.Address ' Try to access .Address property
    On Error GoTo 0
    If s = "" Then Set tbl.rngRows = Nothing
    Exit Function
    
ErrorExit:
    DeleteFlaggedRows = False
End Function
'------------------------------------------------------------------------------------------------
' Performs Conditional Format on column
' Modified 6/9/23 JDL cleanup; 6/15/25; 7/8/25 Update .Find to FindInRange
'
Function CondFormat(Step, tbl) As Boolean

    SetErrs CondFormat: If errs.IsHandle Then On Error GoTo ErrorExit
    Dim r As Variant, rngColData As Range
    
    If tbl.rngRows Is Nothing Then Exit Function

    'Check that column is specified and exists
    With Step
        If errs.IsFail(Len(.sColName) < 1, 1) Then GoTo ErrorExit
        Set r = FindInRange(tbl.rngHeader, .sColName)
        If errs.IsFail(r Is Nothing, 2) Then GoTo ErrorExit
    End With
    
    'Apply conditional formatting to hide vals that match previous row's val
    Set rngColData = Intersect(r.EntireColumn, tbl.rngRows)
    If Not PopulateCondFormat(Step, rngColData) Then GoTo ErrorExit
    
    'Set fill color to light gray to highlight white text from conditional formatting
    With rngColData.Interior
        .Pattern = xlSolid
        .ThemeColor = xlThemeColorDark1
        .TintAndShade = -0.1
    End With

    'Set borders around blocks of same values in column
    If Not ApplyBordersByValBlocks(Step, r, rngColData) Then GoTo ErrorExit
    tbl.wksht.Cells(1, 1).Activate
    Exit Function
    
ErrorExit:
    CondFormat = False
End Function
'------------------------------------------------------------------------------------------------
' Populate the column's data rows with conditional format condition to hide duplicate vals
'
' =AND(SUBTOTAL(103,A1)=1, A2=A1) doesn't white out vals where previous row is hidden
' =A2=A1 is a basic formula that works but it whites out vals where prev row hidden/filtered
'
' Called by CondFormat
'
' JDL 12/19/22 (See Digital Transformation with Excel training for more background)
'          Modified 6/9/23 JDL cleanup
'
Function PopulateCondFormat(Step, rngColData) As Boolean
    SetErrs PopulateCondFormat: If errs.IsHandle Then On Error GoTo ErrorExit
    Dim addrA2 As String, addrA1 As String, str As String
    
    'Delete previous conditions if any
    rngColData.EntireColumn.FormatConditions.Delete

    With rngColData.Cells(1)
    
        'Build conditional format formula string in terms of row 2 and row 1 addresses
        addrA2 = .Address(False, False)
        addrA1 = .Offset(-1, 0).Address(False, False)
        str = "=AND(SUBTOTAL(103," & addrA1 & ")=1, " & addrA2 & "=" & addrA1 & ")"
        
        'Apply the condition to the column's data row range
        .FormatConditions.Add Type:=xlExpression, Formula1:=str
        .FormatConditions(1).Font.Color = RGB(255, 255, 255) 'white
        .Copy
        rngColData.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone
    End With
    Exit Function

ErrorExit:
    PopulateCondFormat = False
End Function
'------------------------------------------------------------------------------------------------
' Apply borders to blocks of rows with same value
'
' Called by CondFormat
'
' JDL 12/19/22 (See Digital Transformation with Excel training for more background)
'      Modified 6/9/23 JDL cleanup
'
Function ApplyBordersByValBlocks(Step, r, rngColData) As Boolean
    SetErrs ApplyBordersByValBlocks: If errs.IsHandle Then On Error GoTo ErrorExit
    Dim rngSame As Range
    
    'Apply borders to blocks of rows with same value
    SetBorders r.EntireColumn, xlNone, True
    Set rngSame = r
    For Each r In Union(r, rngColData)
    
        'If at bottom of data or data value changes in next row
        'xxx Error if R or R.Offset has #REF! error
        If Intersect(r.Offset(1, 0), rngColData) Is Nothing Or r <> r.Offset(1, 0) Then
            SetBorders rngSame, xlContinuous, False
            Set rngSame = r.Offset(1, 0)
        
        'Otherwise, Keep going without applying border
        Else
            Set rngSame = Union(rngSame, r.Offset(1, 0))
        End If
    Next r
    Exit Function

ErrorExit:
    ApplyBordersByValBlocks = False
End Function

'------------------------------------------------------------------------------------------------
' Performs Sort Step - modified to handle tables not homed to A1
'
' Modified 6/9/23 JDL cleanup
'
Function SortBy(Step, tbl) As Integer

    SetErrs SortBy: If errs.IsHandle Then On Error GoTo ErrorExit
    Dim r As Range, RSortBy As Range, arySortBy As Variant, i As Integer
        
    If tbl.rngRows Is Nothing Then Exit Function
    
    'Exit sort-by not specified or if column not found
    With Step
        If errs.IsFail(Len(.sSortBy) < 1, 13) Then GoTo ErrorExit
        arySortBy = Split(.sSortBy, ",")
    End With

    'Add each sort column to SortFields; Exit if specified column not found in the table
    With tbl.wksht.Sort.SortFields
        .Clear
        For i = LBound(arySortBy) To UBound(arySortBy)
            Set RSortBy = tbl.rngHeader.Find(arySortBy(i), lookat:=xlWhole)
            If errs.IsFail(RSortBy Is Nothing, 14) Then GoTo ErrorExit
            .Add key:=RSortBy, SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
        Next i
    End With
    
    'Sort the table by the SortFields
    With tbl.rngHeader.Parent.Sort
        .SetRange Intersect(tbl.rngHeader.EntireColumn, tbl.rngRows)
        .Header = xlNo
        .MatchCase = False
        .Orientation = xlTopToBottom
        .Apply
        .SortFields.Clear
    End With
    Exit Function
    
ErrorExit:
    SortBy = False
End Function
'------------------------------------------------------------------------------------------------
'Add dropdown validation list to column's data
'
' Modified 6/9/23 JDL cleanup
'
Function AddValLst(Step, tbl) As Boolean

    SetErrs AddValLst: If errs.IsHandle Then On Error GoTo ErrorExit
    Dim r As Range
    
    If tbl.rngRows Is Nothing Then Exit Function
    
    'Perform checks and exit if error
    With Step
        If errs.IsFail(Len(.sColName) < 1, 1) Then GoTo ErrorExit
        Set r = tbl.rngHeader.Find(.sColName, lookat:=xlWhole)
        If errs.IsFail(r Is Nothing, 2) Then GoTo ErrorExit
        If errs.IsFail(Len(.sLstRngName) < 1, 11) Then GoTo ErrorExit
        If errs.IsFail(Not RangeExists(tbl.wkbk, .sLstRngName), 12) Then GoTo ErrorExit
    
        'Add the dropdown to the column's data rows
        AddValidationList Intersect(r.EntireColumn, tbl.rngRows), "=" & .sLstRngName
    End With
    Exit Function
    
ErrorExit:
    AddValLst = False
End Function
'------------------------------------------------------------------------------------------------
' Split Screen Columns - 7/29/21 (moved to RecipeStep Class 12/19/22)
'
' Modified 6/9/23 JDL cleanup
'
Function SplitColumns(Step, tbl) As Boolean

    SetErrs SplitColumns: If errs.IsHandle Then On Error GoTo ErrorExit
    Dim r As Range
    
    'Check that screen split location specified and column exists
    With Step
        If errs.IsFail(Len(.sColName) < 1, 9) Then GoTo ErrorExit
        Set r = tbl.rngHeader.Find(.sColName, lookat:=xlWhole)
        If errs.IsFail(r Is Nothing, 10) Then GoTo ErrorExit
    End With
    
    'Activate to allow direct call (wksht is already active if called by dialog OK driver)
    tbl.wksht.Activate
    With ActiveWindow
        .FreezePanes = False
        .SplitColumn = r.Column
        .FreezePanes = True
    End With
    Exit Function
    
ErrorExit:
    SplitColumns = False
End Function
'------------------------------------------------------------------------------------------------
' Freeze table Row 1 (calls tblRowsCols method)
'
' Modified 6/9/23 JDL cleanup
'
Function FreezeRow1(Step, tbl) As Boolean
    SetErrs FreezeRow1: If errs.IsHandle Then On Error GoTo ErrorExit
    
    If Not tbl.FreezeRow1(tbl) Then GoTo ErrorExit
    
    Exit Function
ErrorExit:
    FreezeRow1 = False
End Function
'------------------------------------------------------------------------------------------------
' Rename a column
' Modified 1/2/25 correct IsSetColNames attribute
'
Function Rename(Step, tbl) As Boolean
    SetErrs Rename: If errs.IsHandle Then On Error GoTo ErrorExit
    Dim r As Range, prefix As String
    
    'Exit if column or new name are not specified or if column is not in table
    With Step
        If errs.IsFail(Len(.sColName) < 1, 1) Then GoTo ErrorExit
        If errs.IsFail(Len(.sCol2) < 1, 8) Then GoTo ErrorExit
        Set r = tbl.rngHeader.Find(.sColName, lookat:=xlWhole)
        If errs.IsFail(r Is Nothing, 2) Then GoTo ErrorExit
        
        'Rename the header cell and replace its column name
        r.value = .sCol2
        If tbl.IsSetColNames Then
            
            'Delete range name for previous column name
            prefix = ""
            If tbl.IsNamePrefix Then prefix = tbl.NamePrefix & "_"
            DeleteXLName tbl.wkbk, prefix & xlName(.sColName)
            
            'Rename the column with new name
            If Not tbl.NameColumn(tbl, r) Then GoTo ErrorExit
        End If
    End With
    Exit Function
    
ErrorExit:
    Rename = False
End Function
'------------------------------------------------------------------------------------------------
' Delete a column
'
' Modified 6/9/23 JDL cleanup
'
Function Delete(Step, tbl) As Boolean
    SetErrs Delete: If errs.IsHandle Then On Error GoTo ErrorExit
    Dim r As Range
    
    'Exit if column is not specified or is not in table
    With Step
        If errs.IsFail(Len(.sColName) < 1, 1) Then GoTo ErrorExit
        Set r = tbl.rngHeader.Find(.sColName, lookat:=xlWhole)
        If errs.IsFail(r Is Nothing, 2) Then GoTo ErrorExit
        r.EntireColumn.Delete
    End With
    Exit Function
    
ErrorExit:
    Delete = False
End Function
'------------------------------------------------------------------------------------------------
' Add a column grouping to a table
'
' Modified 5/6/25 change to xlThemeColorAccent3
'
Function AddGroup(Step, tbl) As Boolean
    SetErrs AddGroup: If errs.IsHandle Then On Error GoTo ErrorExit
    Dim colStart As Range, colEnd As Range, iColor1 As Integer, iColor As Integer
    
    'Check that column names are specified
    With Step
        If errs.IsFail(Len(.sColName) < 1, 5) Then GoTo ErrorExit
        If errs.IsFail(Len(.sCol2) < 1, 5) Then GoTo ErrorExit

        'Check that specified columns exist and that end column follows start column
        Set colStart = tbl.rngHeader.Find(.sColName, lookat:=xlWhole)
        Set colEnd = tbl.rngHeader.Find(.sCol2, lookat:=xlWhole)
        If errs.IsFail(colStart Is Nothing Or colEnd Is Nothing, 6) Then GoTo ErrorExit
        Set colStart = colStart.Offset(0, 1)
        If errs.IsFail(colEnd.Column < colStart.Column, 7) Then GoTo ErrorExit
            
        'Add the Group
        Range(colStart.EntireColumn, colEnd.EntireColumn).Columns.Group
        
        'Set header color; Toggle group color to Blue if previous group is not blue
        iColor1 = xlThemeColorAccent1
        iColor = xlThemeColorAccent3
        If colStart.Column > 2 Then
            If colStart.Offset(0, -2).Interior.ThemeColor <> iColor1 Then iColor = iColor1
        End If
        Range(colStart.Offset(0, -1), colEnd).Interior.ThemeColor = iColor
    End With
    Exit Function
    
ErrorExit:
    AddGroup = False
End Function
'------------------------------------------------------------------------------------------------
' Add number formatting to a column
'
' Modified 6/9/23 JDL cleanup
'
Function SetNumFmt(Step, tbl) As Boolean
    SetErrs SetNumFmt: If errs.IsHandle Then On Error GoTo ErrorExit
    Dim r As Range
    
    If tbl.rngRows Is Nothing Then Exit Function
    
    'Check that column exists
    With Step
        If errs.IsFail(Len(.sColName) < 1, 1) Then GoTo ErrorExit
        Set r = tbl.rngHeader.Find(.sColName, lookat:=xlWhole)
        If errs.IsFail(r Is Nothing, 2) Then GoTo ErrorExit
        Intersect(r.EntireColumn, tbl.rngRows).NumberFormat = .NumFmt
    End With
    Exit Function
    
ErrorExit:
    SetNumFmt = False
End Function
'------------------------------------------------------------------------------------------------
'Set Column width
'
' Modified 6/9/23 JDL cleanup
'
Function SetWidth(Step, tbl) As Boolean
    SetErrs SetWidth: If errs.IsHandle Then On Error GoTo ErrorExit
    Dim r As Range
    
    'Check that column exists and that specified width is integer and within allowable limits
    With Step
        If errs.IsFail(Len(.sColName) < 1, 1) Then GoTo ErrorExit
        Set r = tbl.rngHeader.Find(.sColName, lookat:=xlWhole)
        If errs.IsFail(r Is Nothing, 2) Then GoTo ErrorExit
        If errs.IsFail(Not IsNumeric(.Width), 4) Then GoTo ErrorExit
        If errs.IsFail(.Width < 0 Or .Width > 255, 4) Then GoTo ErrorExit
        r.EntireColumn.ColumnWidth = .Width
    End With
    Exit Function
    
ErrorExit:
    SetWidth = False
End Function
'------------------------------------------------------------------------------------------------
'Insert header cell comment
'
' Modified 6/9/23 JDL cleanup
'
Function HeaderComment(Step, tbl) As Boolean
    SetErrs HeaderComment: If errs.IsHandle Then On Error GoTo ErrorExit
    Dim r As Range
    
    'Check that column is specified and exists
    With Step
        If errs.IsFail(Len(.sColName) < 1, 1) Then GoTo ErrorExit
        Set r = tbl.rngHeader.Find(.sColName, lookat:=xlWhole)
        If errs.IsFail(r Is Nothing, 2) Then GoTo ErrorExit
        AddComment r, .sComment
    End With
    Exit Function
    
ErrorExit:
    HeaderComment = False
End Function
'------------------------------------------------------------------------------------------------
'Insert column; optionally with specified formula
'
' Modified 10/11/24 Update tbl.IsSetColNames attribute
'          10/24/24 no recalc after insert every column (in lieu of Refresh.SetInsertFormatsAndVals)
'           6/15/25; 7/8/25 Update .Find to FindInRange
'
Function InsertCol(Step, tbl) As Boolean
    SetErrs InsertCol: If errs.IsHandle Then On Error GoTo ErrorExit
    Dim r As Range, colAfter As Range, colInsert As Range
    
    With tbl
        
        'Exit if column is not specified; delete previous version of column
        If errs.IsFail(Len(Step.sColName) < 1, 1) Then GoTo ErrorExit
        Set r = FindInRange(tbl.rngHeader, Step.sColName)
        If Not r Is Nothing Then r.EntireColumn.Delete
        Set r = Nothing
    
        'check Insert After column exists if specified
        If Len(Step.sCol2) > 1 Then
            Set r = FindInRange(tbl.rngHeader, Step.sCol2)
            If errs.IsFail(r Is Nothing, 3) Then GoTo ErrorExit
            Set r = r.Offset(0, 1)
        End If
        
        'No Insert After column, so insert as first table column
        If r Is Nothing Then Set r = .rngHeader.Cells(1)
    
        'Insert the column and clear leftover format
        Set colAfter = r.EntireColumn
        If Not InsertAndClear(Step, colInsert, colAfter, tbl) Then GoTo ErrorExit
                
        If Not .rngRows Is Nothing Then
        
            'Set borders
            SetBorders Intersect(colInsert, .rngRows), xlContinuous, True
            If Len(Step.sFormula) > 0 Then
            
                'Populate formula into data rows (after error check); optionally paste vals
                If Not PopulateFormula(Step, colInsert, tbl) Then GoTo ErrorExit
                If Not FormatFormulaRange(Step, colInsert, tbl) Then GoTo ErrorExit
            End If
        End If

        'Name inserted column
        If .IsSetColNames Then
            If Not .NameColumn(tbl, Intersect(colInsert, .rngHeader)) Then GoTo ErrorExit
        End If
    End With
    Exit Function

ErrorExit:
    InsertCol = False
End Function
'------------------------------------------------------------------------------------------------
' Check formula syntax and populate formula into inserted col's data range
' Called by InsertCol
'
' Modified 6/9/23 JDL cleanup
'
Function InsertAndClear(Step, colInsert, colAfter, tbl) As Boolean
    SetErrs InsertAndClear: If errs.IsHandle Then On Error GoTo ErrorExit
    
    colAfter.Insert
    Set colInsert = colAfter.Offset(0, -1)
    With colInsert
        .ClearFormats
        .Validation.Delete
    If .Column = 1 Then Set tbl.rngHeader = Union(tbl.rngHeader, .Cells(1))
    End With

    'Write header name and format header
    With Intersect(colInsert, tbl.rngHeader)
        .value = Step.sColName
        .Style = "Accent1"
    End With
    Exit Function

ErrorExit:
    InsertAndClear = False
End Function
'------------------------------------------------------------------------------------------------
' Check formula syntax and populate formula into inserted col's data range
' Called by InsertCol
'
' Modified 6/9/23 JDL cleanup
'
Function PopulateFormula(Step, colInsert, tbl) As Boolean
    SetErrs PopulateFormula: If errs.IsHandle Then On Error GoTo ErrorExit
    Dim IsOK As Boolean
    
    With tbl
        IsOK = IsValidFormulaSyntax(Intersect(colInsert, .rngRows.Rows(1)), Step.sFormula)
        If errs.IsFail(Not IsOK, 15) Then GoTo ErrorExit
        Intersect(colInsert, .rngRows).Formula = Step.sFormula
    End With
    Exit Function

ErrorExit:
    PopulateFormula = False
End Function
'------------------------------------------------------------------------------------------------
' Format as Calculation if live formula; otherwise recalc and paste over values
' Called by InsertCol
'
' Modified 5/6/25
'
Function FormatFormulaRange(Step, colInsert, tbl) As Boolean
    SetErrs FormatFormulaRange: If errs.IsHandle Then On Error GoTo ErrorExit
    Dim IsFormulaOK As Boolean

    With Intersect(colInsert, tbl.rngRows)
        If Step.IsKeepFormula Then
            .Style = "Calculation"
            
            '5/6/25 change color to black but keep bold and background from Calculation
            .Font.ColorIndex = xlAutomatic
        Else
            Application.Calculate
            .Copy
            .PasteSpecial Paste:=xlPasteValues
            Application.CutCopyMode = False
        End If
    End With
    Exit Function

ErrorExit:
    FormatFormulaRange = False
End Function