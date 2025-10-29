'ExcelSteps_Utilities.vb
'This module is part of the ExcelSteps open source project posted at:
'https://github.com/jlandgre/ExcelSteps/. It is licensed under the MIT open source license
' version 10/24/25
Option Explicit
'-------------------------------------------------------------------------------------
'-------------------------------------------------------------------------------------
' Open file at path relative to specified workbook (cross-platform compatible)
' JDL 8/1/25; Modified 10/1/25
'
Public Function OpenFile(ByVal fullpath As String, wkbkOpened As Workbook) As Boolean
    SetErrs OpenFile: If errs.IsHandle Then On Error GoTo ErrorExit
    
    ' Check if file exists
    If errs.IsFail(Dir(fullpath) = "", 1, fullpath) Then GoTo ErrorExit
    
    ' Open the workbook
    Set wkbkOpened = Workbooks.Open(fullpath)
    If errs.IsFail(wkbkOpened Is Nothing, 2, fullpath) Then GoTo ErrorExit
    Exit Function
    
ErrorExit:
    errs.RecordErr "OpenFile", OpenFile
End Function
'-------------------------------------------------------------------------------------
' True if specified sheet exists
'
Function SheetExists(ByVal wkbk As Workbook, ByVal sht As String) As Boolean
    Dim w As Variant
    SheetExists = True
    For Each w In wkbk.Sheets
        If w.Name = sht Then Exit Function
    Next w
    SheetExists = False
End Function
'-------------------------------------------------------------------------------------
'Clear sheet's outline and set directions to put summary column left of detail
'
' JDL 12/16/22  Modified 9/5/23 switch to ClearOutline method
'
Function ClearAndResetOutline(wksht) As Boolean
    SetErrs ClearAndResetOutline: If errs.IsHandle Then On Error GoTo ErrorExit

    'ClearColumnOutline wksht
    wksht.Columns.ClearOutline
    With wksht.Outline
        .AutomaticStyles = False
        .SummaryRow = xlAbove
        .SummaryColumn = xlLeft
    End With
    Exit Function
    
ErrorExit:
    errs.RecordErr "ClearAndResetOutline", ClearAndResetOutline
End Function
'-------------------------------------------------------------------------------------
' Clear Column Outline on wksht
'
Sub ClearColumnOutline(wksht)
    wksht.Activate
    Do While True
        wksht.Columns.Ungroup
    Loop
End Sub
'-------------------------------------------------------------------------------------
' Flag redundant values with comments within a range
'
' Modified: 10/1/25 Cleanup
'
Function FlagRedundant(ByVal rng) As Boolean
    SetErrs FlagRedundant: If errs.IsHandle Then On Error GoTo ErrorExit
    Dim rngRedundant As Range, c As Range, r As Range
    
    rng.ClearComments
    For Each r In rng
        Set rngRedundant = RedundantCells(r, rng)
        If Not rngRedundant Is Nothing Then
            For Each c In rngRedundant
                errs.iCodeLocal = 2
                errs.LookupCommentMsg c, errs.Locn
            Next c
            errs.ErrParam = rng.Parent.Name
            errs.iCodeLocal = 1
            GoTo ErrorExit
        End If
    Next r
    Exit Function
    
ErrorExit:
    errs.RecordErr "FlagRedundant", FlagRedundant
End Function
'-------------------------------------------------------------------------------------
' Check whether a name is redundant within a range
'
Function NameIsRedundant(ByVal curCell, ByVal rng) As Boolean
    Dim w As Range, v As Range
    
    NameIsRedundant = False
    Set w = rng.Find(curCell.value, lookat:=xlWhole)
    If Not w.Address = curCell.Address Then
        NameIsRedundant = True
    Else
        Set w = rng.FindNext(w)
        If Not w.Address = curCell.Address Then NameIsRedundant = True
    End If
End Function
'-------------------------------------------------------------------------------------
' Determine whether named range exists in the workbook
'
Function RangeExists(wkbk, sName) As Boolean
Dim w As Variant
    RangeExists = False
    For Each w In wkbk.Names
        If UCase(w.Name) = UCase(sName) Then
            RangeExists = True
            Exit Function
        End If
    Next w
End Function
'-------------------------------------------------------------------------------------
' VBA version of Excel's TRUNC function
'
Function TRUNC(num As Double, digits) As Double
    TRUNC = Fix(num * 10 ^ digits) / 10 ^ digits
End Function
'-------------------------------------------------------------------------------------
' Write a text string or block to a file
' Created: 7/13/21 JDL  Modified 10/1/25 Cleanup
'
Function WriteFile(sFilePath, sText, Optional IsAppend) As Boolean
    SetErrs WriteFile: If errs.IsHandle Then On Error GoTo ErrorExit
    
    If IsMissing(IsAppend) Then IsAppend = False
    If IsAppend Then
        Open sFilePath For Append As #1
    Else
        Open sFilePath For Output As #1
    End If
    Print #1, sText
    Close #1
    Exit Function
    
ErrorExit:
    errs.RecordErr "WriteFile", WriteFile
End Function
'-------------------------------------------------------------------------------------
' Test whether a range was deleted and is in "Object Required" state
' See DGlancy answer here:  https://stackoverflow.com/questions
' /12127311/vba-what-happens-to-range-objects-if-user-deletes-cells
'
' Created:   9/15/21 JDL
'
Function IsRngDeleted(rng) As Boolean
    Dim sAddress As String
    IsRngDeleted = False
    On Error Resume Next
    sAddress = rng.Address
    If err.Number = 424 Then IsRngDeleted = True
End Function
'-------------------------------------------------------------------------------------
' Populate Array of unique values from cells in specified range
' Created:   3/12/22; Modified 10/1/25 clanup
'
Function aryUniqueValsInRange(r)
    Dim aryUnique As Variant, w As Variant
    If r Is Nothing Then Exit Function
    aryUnique = Array()
    For Each w In r.Cells
        If Not IsEmpty(w) Then
            If Not IsInAry(aryUnique, w.value) Then _
                aryUnique = aryAppendVal(aryUnique, w.value)
        End If
    Next w
    aryUniqueValsInRange = aryUnique
End Function
'-------------------------------------------------------------------------------------
' Append a specified val to end of specified array
' Inputs: aryOrig [Variant - Array] array to append onto
'         val [Variant] value to append
'
' JDL Modified 3/3/21 to deal properly empty array as input
'
Function aryAppendVal(aryOrig, val) As Variant
    Dim aryTemp As Variant

    'Initialize an empty array to hold result
    aryTemp = aryOrig
    If Not IsArray(aryTemp) Then aryTemp = Array()

    'If empty array, initialize to one element
    If UBound(aryTemp) = -1 Then
        ReDim Preserve aryTemp(0 To 0)
    Else
        ReDim Preserve aryTemp(0 To UBound(aryTemp) + 1)
    End If

    aryTemp(UBound(aryTemp)) = val
    aryAppendVal = aryTemp
End Function
'-------------------------------------------------------------------------------------
' Combine path with file or folder name with fso.BuildPath
' Modified 2/13/25 for Mac compatibility
'
Function BuildPath(sPath As String, sFile As String) As String
    #If Not Mac Then
        Dim fso As Object
        Set fso = CreateObject("Scripting.FileSystemObject")
        BuildPath = ""
        If Len(sFile) > 0 Then BuildPath = fso.BuildPath(sPath, sFile)
    #Else
        ' Alternative code for Mac (if needed)
        BuildPath = sPath & Application.PathSeparator & sFile
    #End If
End Function
'-------------------------------------------------------------------------------------
' Build a comma-separated list of column headers from colinfo
' Inputs:  wkbk [Workbook] Workbook object containing colinfo
'         sTable [String] tbl name (colinfo selector col name) with col order integers
'
' Created:   10/29/20 JDL      Modified: 7/9/21 Convert ColInfo to tblRowsCols
'                                     2/28/23 Add error trapping and refactor
'                                     8/13/25
'
Function BuildHeaderListFromColInfo(wkbk, ByVal sTable) As String

    SetErrs "driver": If errs.IsHandle Then On Error GoTo ErrorExit
    Dim tblCI As New tblRowsCols, ary As Variant, rngVals As Range, c As Variant
    
    With tblCI
    
        'Provision/sort ColInfo by specified table's column order column (integers)
        If Not .Provision(tblCI, wkbk, False, sht:=shtColInfo) Then GoTo ErrorExit
        StepSortBy sTable, .rngHeader, .rngRows

        'Get the range containing table column order numbers
        Set rngVals = .rngTblHeaderVal(tblCI, sTable)
        Set rngVals = Range(rngVals.Offset(1, 0), rngVals.End(xlDown))
        BuildHeaderListFromColInfo = ListFromArray(Intersect(rngVals.EntireRow, _
            .colrngColName).value)
    End With
    Exit Function
    
ErrorExit:
    errs.RecordErr "BuildHeaderListFromColInfo"
End Function
'-------------------------------------------------------------------------------------
' Add drop-down validation to a range
'
Sub AddValidationList(rng, sFormula)
    With rng.Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, _
        Operator:=xlBetween, Formula1:=sFormula
        .IgnoreBlank = True
        .InCellDropdown = True
        .InputTitle = ""
        .ErrorTitle = ""
        .InputMessage = ""
        .ErrorMessage = ""
        .ShowInput = True
        .ShowError = True
    End With
End Sub
'-------------------------------------------------------------------------------------
' Build Comma-separated list from array
' Created: 3/1/21 JDL Modified 11/18/21 Add sDelim optional argument
'
Function ListFromArray(ary, Optional sDelim, Optional IsFormatted As Boolean) As String
    Dim val As Variant, lst As String, i As Integer
    lst = ""
    
    If IsMissing(IsFormatted) Then IsFormatted = False
    If IsMissing(sDelim) Then sDelim = ","

    'Simple delimited list
    If Not IsFormatted Then
        For Each val In ary
            If Len(lst) < 1 Then
                lst = CStr(val)
            Else
                lst = lst & sDelim & CStr(val)
            End If
        Next val
    
    'Formatted, comma-separated list
    Else
        sDelim = ", "
        For i = 0 To UBound(ary)
            val = ary(i)
            If i = 0 Then
                lst = CStr(val)
            Else
                If i = UBound(ary) Then sDelim = " and "
                lst = lst & sDelim & CStr(val)
            End If
        Next i
    End If
    ListFromArray = lst
End Function
'-------------------------------------------------------------------------------------
' Set range for last used cell in a row or col (works with hidden or col outline cells)
' Inputs: cellHome [Range] home (left-most) cell in row to be searched
'         xlDirection [Integer xlToRight or xlDown enumeration]
' finding value in outline-hidden cell. See 6/2/15 response on:
' https://stackoverflow.com/questions/20152328/vba-find-function-cant-find-given-value
'
' JDL 12/3/21 (Based on Search Example.xlsm) Modified 10/1/25 cleanup
'
Function rngLastPopCell(cell, xlDirection)
    Dim rng As Range, c As Range, IsRowSearch As Boolean, IsColSearch As Boolean
    Set rngLastPopCell = cell
    
    'Set search range based on whether column or row search specified
    If xlDirection = xlToRight Then
        IsRowSearch = True
        Set rng = cell.EntireRow
    ElseIf xlDirection = xlDown Then
        IsColSearch = True
        Set rng = cell.EntireColumn
    Else
        Exit Function
    End If
    
    'To find last populated cell, Search xlPrevious with wrap from first cell
    Set c = rng.Cells.Find("*", After:=rng.Cells(1), LookIn:=xlFormulas, _
        SearchDirection:=xlPrevious)
    If Not c Is Nothing Then Set rngLastPopCell = c
End Function
'-------------------------------------------------------------------------------------
' Build multi-cell range containing non-empty cells in intersection of two ranges
'           (Unlike .End(xlUp) etc., it works with hidden rows or columns)
' Inputs: rng1, rng2 [Range] intersecting ranges that form single row or column
'         rng2 is optional if rng1 is desired row or column range
'
' Modified 1/6/22 - Fix bug: search needs to be limited by rng1 --not entire row/col
'          8/13/25
'
Function BuildMultiCellRange(rng1, Optional rng2) As Range
    Dim w As Variant, rng As Range, xlDirection, IsEntireRowCol As Boolean
    If rng1 Is Nothing Then Exit Function
    Set rng = rng1
    If Not IsMissing(rng2) Then Set rng = Intersect(rng1, rng2)
    
    'Set search direction for row or column range
    
    If rng.Rows.Count = 1 Then
        xlDirection = xlToRight
        If rng(rng.Cells.Count).Column = rng.Parent.Columns.Count Then _
            IsEntireRowCol = True
    ElseIf rng.Columns.Count = 1 Then
        xlDirection = xlDown
        If rng(rng.Cells.Count).Row = rng.Parent.Rows.Count Then _
            IsEntireRowCol = True
    Else
        'Error condition - return Nothing
        Exit Function
    End If
    
    'If rng is entire col or row range, restrict search based on last populated cell
    If IsEntireRowCol Then Set rng = Range(rng(1), rngLastPopCell(rng, xlDirection))
    
    For Each w In rng
        If Len(w) > 0 Then
            If BuildMultiCellRange Is Nothing Then
                Set BuildMultiCellRange = w
            Else
                Set BuildMultiCellRange = Union(BuildMultiCellRange, w)
            End If
        End If
    Next w
End Function
'-------------------------------------------------------------------------------------
' Add comment to a cell and deletes previous
'
Sub AddComment(rngCell, sTxt)
    With rngCell
        If Not .Comment Is Nothing Then .Comment.Delete
        .AddComment
        .Comment.Visible = False
        .Comment.Text Text:=sTxt
    End With
End Sub
'-------------------------------------------------------------------------------------
' return [multi-cell] range of cells in rng with matching values
'
Function RedundantCells(cell, rng) As Range
    Dim c As Range
    For Each c In rng
        If c = cell And c.Address <> cell.Address Then
            If RedundantCells Is Nothing Then Set RedundantCells = cell
            Set RedundantCells = Union(RedundantCells, c)
        End If
    Next c
End Function
'-------------------------------------------------------------------------------------
' Check rng cell contents for validity (non-redundant, suitable as Excel range name)
' Modified: 1/18/22 - Add IsEmptyErr check; 8/13/25; 10/1/25 cleanup
'
Function CheckNames(rng) As Boolean

    SetErrs CheckNames: If errs.IsHandle Then On Error GoTo ErrorExit
    Dim rngRedundant As Range, c As Range, w As Range
    Dim IsRedund As Boolean, IsCellRefErr As Boolean, IsInvalid As Boolean
    Dim IsEmptyErr As Boolean, IsFailCheck As Boolean
    Const Locn As String = "CheckNames"
    
    If rng Is Nothing Then Exit Function
    rng.ClearComments
    
    With errs
        For Each w In rng
        
            'Check name is specified - exit since this error precludes other checks
            ' xxx 5/8/25 not really a valid check if mdl IsLite
            ' --ok for initial var name cell blank
            If .IsFail(IsEmpty(w), 1) Then
                .LookupCommentMsg w, Locn
                
                'Set code for fatal error lookup and report address causing error
                .iCodeLocal = 2
                .ErrParam = w.Address
                GoTo ErrorExit
            End If
        
            'Check variable names for redundancy
            Set rngRedundant = RedundantCells(w, rng)
            If .IsFail(Not rngRedundant Is Nothing, 3) Then
                IsFailCheck = True
                For Each c In rngRedundant
                    .LookupCommentMsg c, Locn, IsReinitialize:=False
                Next c
                
                'Reinitialize errs to allow reporting fatal error message
                .Init IsHandle:=.IsHandle
                
            'Check whether name is Excel cell reference or invalid name
            ElseIf .IsFail(IsExcelCellRef(w), 4) Or _
                    .IsFail(Not IsValidExcelName(w), 5) Then
                IsFailCheck = True
                .LookupCommentMsg w, Locn
            End If
        Next w
        
        'If failed for any reason, set fatal error message and error exit
        If .IsFail(IsFailCheck, 6) Then
            .ErrParam = rng.Parent.Name
            GoTo ErrorExit
        End If
    End With
    Exit Function
    
ErrorExit:
    errs.RecordErr Locn, CheckNames
End Function
'-----------------------------------------------------------------------------------
' Set a table row range (multirange) based on array of key values
' Modified 1/17/22 to avoid unhiding hidden columns
'          3/14/23 to include MinWidth; 9/29/25 add CountA so no action if empty
'
Sub ColWidthAutofit(rngHeader, Optional iMaxWidth = 80, Optional iMinWidth = 30)
    Dim c As Variant
    For Each c In rngHeader
        With c.EntireColumn
            If (Not .Hidden = True) And _
                (Not WorksheetFunction.CountA(c.EntireColumn) = 0) Then
                .ColumnWidth = 240
                .AutoFit
                If c.ColumnWidth > iMaxWidth Then c.ColumnWidth = iMaxWidth
                If c.ColumnWidth < iMinWidth Then c.ColumnWidth = iMinWidth
                .ColumnWidth = .ColumnWidth + 2
            End If
        End With
    Next c
End Sub
'-------------------------------------------------------------------------------------
' Set a table row range (multirange) based on array of key values
' Notes:    Mimics Pandas .loc functionality with multiindex
'           Suitable for small tables - Works well up to 500 found items in 1000 rows
'           Bogs down with large number of found items (e.g. 2500 found in 5000 rows)
' Inputs:   tbl [table Class instance]
'           aryKeyColRanges [array of Ranges] table Class column ranges for key cols
'           aryKeyValues [array of values (variant)] array of key column values
' Built/validated in Unique Vals_1220.xlsm
'
' Created:   12/17/20 JDL      Modified: 8/13/25
'
Function KeyColRng(tbl, aryKeyColRanges, aryKeyValues) As Range
    Dim i As Integer, rngCurrent As Range, rngSearch As Range, rngFound As Range
    
    'Progressively search/subset across the key columns -- start with entire table
    Set rngCurrent = tbl.rngRows
    For i = LBound(aryKeyColRanges) To UBound(aryKeyColRanges)
    
        'Stop if previous search returned Nothing
        If Not rngCurrent Is Nothing Then
            Set rngSearch = Intersect(rngCurrent, aryKeyColRanges(i))
            Set rngCurrent = Nothing
            On Error Resume Next 'in case nothing
            Set rngCurrent = FindAll(rngSearch, aryKeyValues(i)).EntireRow
            On Error GoTo 0
        End If
    Next i

    If Not rngCurrent Is Nothing Then Set KeyColRng = rngCurrent.EntireRow
End Function
'-------------------------------------------------------------------------------------
' Range Find that works with hidden cells and cells in column or row outline
' https://www.mrexcel.com/board/threads/vba-cannot-find-in-if-cells-are-hidden-even-
' /x/if-xlformulas-is-used.518661/
'
' JDL 5/6/20  Modified 5/2/25 to iterate over rng.Areas (fix bug with non-contiguous)
'             Minor comment updates 8/13/25
'
Function FindInRange(ByVal rng As Range, ByVal val) As Range
    Dim area As Range, i As Integer, q As String
    Set FindInRange = Nothing

    ' If val is a string, wrap it in
    If VarType(val) = vbString Then q = """"

    ' Iterate over contiguous areas for reliability of Evaluate with MATCH
    For Each area In rng.Areas
        On Error Resume Next
        i = Evaluate("MATCH(" & q & val & q & "," & area.Address(External:=True) _
            & ",0)")
        On Error GoTo 0
        If i > 0 Then
            Set FindInRange = area.Cells(i)
            Exit Function
        End If
    Next area
End Function
'-------------------------------------------------------------------------------------
'Build multirange with all occurrences of specified value (case sensitive)
'Inspired by http://www.cpearson.com/excel/findall.aspx
' (JDL Find_All_Pearson_Version.xlsm)
'
'JDL 12/14/22; 8/13/25
'
Function FindAll(rngSearch, FindWhat As Variant)
    Dim c As Range, cFirst As Range, iRow As Long, iCol As Long
    
    'Initialize with first found cell
    Set FindAll = rngSearch.Find(FindWhat, lookat:=xlWhole)
    If FindAll Is Nothing Then Exit Function
    Set cFirst = FindAll
    Set c = FindAll
    
    'Loop until search returns to first, found cell
    Do Until False
        Set c = rngSearch.FindNext(After:=c)
        If (c.Address = cFirst.Address) Then Exit Do
        Set FindAll = Application.Union(FindAll, c)
    Loop
End Function
'-------------------------------------------------------------------------------------
' Construct name string in RC format for Excel name creation
' JDL 5/7/20 - Modified 9/15/20 to add Multiple Rows and Columns; 10/1/25 cleanup
'
Function MakeRefNameString(sht, Optional irow1 = 0, Optional irow2 = 0, _
    Optional icol1 = 0, Optional icol2 = 0) As String
    MakeRefNameString = "='" & sht & "'!"
    
    'Entire row
    If irow1 > 0 And irow2 = 0 Then
        MakeRefNameString = MakeRefNameString & "R" & irow1 & ":R" & irow1
    
    'Entire column
    ElseIf icol1 > 0 And icol2 = 0 Then
        MakeRefNameString = MakeRefNameString & "C" & icol1 & ":C" & icol1
    
    'Multiple rows
    ElseIf irow1 > 0 And irow2 > 0 And icol1 = 0 And icol2 = 0 Then
        MakeRefNameString = MakeRefNameString & "R" & irow1 & ":R" & irow2
        
    'Multiple columns
    ElseIf irow1 = 0 And irow2 = 0 And icol1 > 0 And icol2 > 0 Then
        MakeRefNameString = MakeRefNameString & "C" & icol1 & ":C" & icol2
    
    'Range
    ElseIf icol1 > 0 And icol2 > 0 And irow1 > 0 And irow2 > 0 Then
        
        'Single cell
        If irow1 = irow2 And icol1 = icol2 Then
            MakeRefNameString = MakeRefNameString & "R" & irow1 & "C" & icol1
            
        'Block Range
        Else
            MakeRefNameString = MakeRefNameString & "R" & irow1 & "C" & icol1 & ":R" _
                & irow2 & "C" & icol2
        End If
    End If
End Function
'-------------------------------------------------------------------------------------
' Create a valid Excel name from string
'
Function xlName(str) As String
    Dim i As Integer, s As String
        
    'Initialize result string and check each character in str (condense/skip spaces)
    xlName = ""
    For i = 1 To Len(str)
        s = Mid(str, i, 1)
                
        'If it's an invalid character but not a space, replace it with an underscore
        If InStr(sXLChars, LCase(s)) < 1 And s <> " " Then
            xlName = xlName & "_"
            
        'If it's a valid character, simply add it to the XLName
        ElseIf s <> " " Then
            xlName = xlName & s
        End If
    Next i
    
    'If necessary, add  underscore prefix if str starts with a number or is column name
    If InStr(sXLFirstChars, LCase(Left(xlName, 1))) < 1 Or _
        IsExcelCellRef(xlName) Then xlName = "_" & xlName
End Function
'-------------------------------------------------------------------------------------
' Return value of specified setting; returns nothing if not found
' JDL 4/2/20 Modified: 5/19/21 Add exit if no shtSettings; 10/1/25 cleanup
'
Function ReadSetting(wkbk, ByVal sName As String) As Variant
    Dim c As Range
    If Not SheetExists(wkbk, shtSettings) Then Exit Function
    Set c = wkbk.Sheets(shtSettings).Columns(1).Find(sName, lookat:=xlWhole)
    If c Is Nothing Then Exit Function
    ReadSetting = c.Offset(0, 1)
End Function
'-------------------------------------------------------------------------------------
' Update Setting value; create new if not found
' JDL 4/2/20  Modified: 4/14/20 Addsheet; 1/10/22 docstring; 10/1/25 cleanup
'
Sub UpdateSetting(wkbk, ByVal sName As String, ByVal val As Variant)
    Dim c As Range
    If Not SheetExists(wkbk, shtSettings) Then _
        AddSheet wkbk, shtSettings, wkbk.Sheet(wkbk.Sheets.Count)

    'Find the existing setting row or add a new setting if  not found
    With wkbk.Sheets(shtSettings)
        Set c = .Columns(1).Find(sName, lookat:=xlWhole)
        If c Is Nothing Then Set c = .Cells(.Rows.Count, 1).End(xlUp).Offset(1, 0)
        c = sName
        c.Offset(0, 1) = val
        '.Visible = xlVeryHidden
    End With
End Sub
'-------------------------------------------------------------------------------------
'   DeleteSetting - Deletes a setting from Settings sheet (if found there)
'   JDL 5/1/20  Modified: 7/13/21 Exit if no shtSettings; 10/1/25 cleanup
'
Function DeleteSetting(wkbk, sName As String) As Variant
    Dim c As Range
    DeleteSetting = False
    If Not SheetExists(wkbk, shtSettings) Then Exit Function
    
    Set c = wkbk.Sheets(shtSettings).Columns(1).Find(sName, lookat:=xlWhole)
    If c Is Nothing Then Exit Function
    c.EntireRow.Delete
    DeleteSetting = True
End Function
'-------------------------------------------------------------------------------------
' Append a value to an existing setting (creates new setting if no previous)
' Created:   8/24/20 JDL      Modified: 8/26/20; 10/1/25 cleanup
'
Sub AppendSetting(wkbk, sName As String, val As Variant)
    Dim c As Range

    'Find the existing setting row or add a new setting if  not found
    With wkbk.Sheets(shtSettings)
        Set c = .Columns(1).Find(sName, lookat:=xlWhole)
        If c Is Nothing Then Set c = .Cells(.Rows.Count, 1).End(xlUp).Offset(1, 0)
        c = sName
        c.Offset(0, 1) = c.Offset(0, 1) & val
    End With
End Sub
'-------------------------------------------------------------------------------------
' Set a range for key column rows that match a value in a sorted table
' Created:   9/23/20 JDL      Modified: 12/16/21 StepSortBy as Subroutine; 8/13/25
'
Function rngKeycolRows(tbl, colrng, val) As Range
    Dim cellFirst As Range, cellLast As Range, filter As New clsFilter
    With tbl
        
        'Store and Remove table filtering
        filter.CaptureExisting filter, .rngTable, .rngHeader
        .wkbk.Sheets(.sht).AutoFilterMode = False
        
        'Sort the table by colRng
        StepSortBy Intersect(.rngHeader, colrng.EntireColumn), .rngHeader, .rngRows

        'Start search for val in first cell of colRng
        Set cellFirst = colrng.Find(val, After:=colrng.Cells(colrng.Cells.Count), _
            lookat:=xlWhole, SearchDirection:=xlNext)
        If cellFirst Is Nothing Then Exit Function
        Set rngKeycolRows = cellFirst

        'Search previous from first cell of colRng (starts in last cell)
        Set cellLast = colrng.Find(val, lookat:=xlWhole, SearchDirection:=xlPrevious)
        If cellLast Is Nothing Then Exit Function
        Set rngKeycolRows = Range(cellFirst, cellLast).EntireRow
        
        'Replace previous filtering
        If filter.IsShtFilter Then filter.AddToTable filter
    End With

Exit Function
ErrorExit:
    MsgBox "Error setting row range for lookup"
    Set rngKeycolRows = Nothing
End Function
'-------------------------------------------------------------------------------------
' Performs Sort Step - modified to work with non-homed tables
' From modExcelStepsRefresh
'
' Modified JDL 7/8/21
'
Sub StepSortBy(sSortBy, rngheaders, rngData)
    Dim sArySortBy() As String, i As Integer, w As Range

    sArySortBy = Split(sSortBy, ",")

    With rngheaders.Parent.Sort.SortFields
        .Clear
        For i = LBound(sArySortBy) To UBound(sArySortBy)
            Set w = rngheaders.Find(sArySortBy(i), lookat:=xlWhole)

            'Exit if specified sort field not found in the table
            If w Is Nothing Then Exit Sub

            .Add key:=w, SortOn:=xlSortOnValues, Order:=xlAscending, _
                DataOption:=xlSortTextAsNumbers
        Next i
    End With
    With rngheaders.Parent.Sort
        .SetRange Intersect(rngheaders.EntireColumn, rngData)
        .Header = xlNo
        .MatchCase = False
        .Orientation = xlTopToBottom
        .Apply
        .SortFields.Clear
    End With
End Sub
'-------------------------------------------------------------------------------------
' Return the range corresponding to the extent between a pre-defined cell and the
'           last populated cell in either row or column direction
' Inputs:   rng1 [Range] range defining a row or column --cell defining start of
'            rngToExtent search (Intersect(.cellHome.Entirerow, .colrng.entirecolumn)
'           IsRows [Boolean] True if rngToExtent is one or more rows; False if columns
' Outputs:  rng of either rows (IsRows = True) or columns representing the extent
'           between rng1 and last populated cell; default is rng1 row or column if
'           last populated is not beyond rng1
'
' 1/26/21 JDL   Modified: 8/13/25
'
Function rngToExtent(rng1, IsRows) As Range
    Dim xlDirection As Integer, rng2 As Range
    
    'Exit if rng1 is not single row or column
    If (Not IsRows And rng1.Columns.Count > 1) Or _
        (IsRows And rng1.Rows.Count > 1) Then Exit Function
    
    'Initialize the direction based on whether row or column
    If IsRows Then xlDirection = xlDown
    If Not IsRows Then xlDirection = xlToRight
    
    'Find the last populated cell
    Set rng2 = rngLastPopCell(rng1, xlDirection)
    
    'Initialize; reset if last populated cell is beyond rng1
    If IsRows Then
        Set rngToExtent = rng1.EntireRow
        If rng2.Row >= rng1.Row Then _
            Set rngToExtent = Range(rng1.EntireRow, rng2.EntireRow)
    Else
        Set rngToExtent = rng1.EntireColumn
        If rng2.Column >= rng1.Column Then _
            Set rngToExtent = Range(rng1.EntireColumn, rng2.EntireColumn)
    End If
End Function
'-------------------------------------------------------------------------------------
' Return the range to last populated cell in a row or column
'           Uses slower .Offset() method to work where outline may be present
' Inputs:   rng [Range] range defining a row or column --must be single row or column
'                       depending on xlDirection
'           IsRowSearch [Boolean] True if search across rows (return single-col range
' Outputs:  rng representing the extent from rng(1) to last populated row/column
'           outputs Nothing if first cell is empty
'
' 4/7/21   JDL; Modified 8/13/25
'
Function rngToExtentOffset(ByVal rng As Range, ByVal IsRowSearch As Boolean) As Range
    Dim IsColSearch As Boolean, cellHome As Range, cellCur As Range
    Dim i As Integer, j As Integer

    If Not IsRowSearch Then IsColSearch = True

    'Exit if rng is not single row or column
    If IsColSearch And rng.Columns.Count > 1 Then Exit Function
    If IsRowSearch And rng.Rows.Count > 1 Then Exit Function

    'Home cell is the first cell in rng
    Set cellHome = rng.Cells(1)
    If IsEmpty(cellHome) Then Exit Function

    'Move to right or down to find last populated cell
    Set cellCur = cellHome
    i = 1
    j = 1
    If IsColSearch Then i = 0
    If IsRowSearch Then j = 0
    Do While Not IsEmpty(cellCur.Offset(i, j))
        Set cellCur = cellCur.Offset(i, j)
    Loop
    Set rngToExtentOffset = Range(cellHome, cellCur)
End Function
'-------------------------------------------------------------------------------------
' Set borders of the specified range
'   linetype can be xlNone, xlContinuous or other valid LineStyle constants
'   bInterior controls whether interior borders are affected
'
Sub SetBorders(rng, linetype, bInterior)
    Dim i As Variant, ary As Variant
    rng.Borders(xlDiagonalDown).LineStyle = xlNone
    rng.Borders(xlDiagonalUp).LineStyle = xlNone
        
    '7,8,9,10 are edges xlEdgeLeft, Top, Bottom and Right
    '11 and 12 are xlInsideVertical and xlInsideHorizontal
    
    ary = Array(7, 8, 9, 10, 11, 12)
    If Not bInterior Then ary = Array(7, 8, 9, 10)
    For Each i In ary
        With rng.Borders(i)
            .LineStyle = linetype
            If linetype <> xlNone Then
                .ThemeColor = 3
                .TintAndShade = -0.25
                .Weight = xlThin
            End If
        End With
    Next i
End Sub
'-------------------------------------------------------------------------------------
' Check whether a string is a valid Excel name
'
Function IsValidExcelName(str) As Boolean
    Dim i As Integer, s As String
        
    'Check each character for validity in Excel names
    IsValidExcelName = True
    For i = 1 To Len(str)
        s = Mid(str, i, 1)
        If InStr(sXLChars, LCase(s)) < 1 Then
            IsValidExcelName = False
            Exit For
        End If
    Next i
    
    'Check whether first character is valid
    If InStr(sXLFirstChars, LCase(Left(str, 1))) < 1 And _
        LCase(Left(str, 1)) <> "_" Then IsValidExcelName = False
End Function
'-------------------------------------------------------------------------------------
' Test whether string like "ABC27" is a reserved, Excel cell reference
' JDL Modified 11/11/24 to deal with case where numeric suffix is all zeros
'
Function IsExcelCellRef(str) As Boolean
    Dim pos As Integer, IsLetterPrefix As Boolean, IsNumberSuffix As Boolean
    Dim IsPrefix As Boolean, s As String, IsAllZeros As Boolean
    IsExcelCellRef = True
    
    'Test whether three letter or less prefix
    pos = 1
    IsPrefix = True
    IsLetterPrefix = True
    While pos <= 4 And pos <= Len(str) And IsPrefix
        If InStr(sXLFirstChars, LCase(Mid(str, pos, 1))) < 1 Then IsPrefix = False
        pos = pos + 1
    Wend
    pos = pos - 1
    If pos < 2 Then IsLetterPrefix = False

    'Test whether suffix is all digits
    If IsLetterPrefix Then
        IsNumberSuffix = True
        IsAllZeros = True
        
        While pos <= Len(str) And IsNumberSuffix
            s = LCase(Mid(str, pos, 1))
            
            'Check that number is non-zero
            If CStr(s) <> "0" Then IsAllZeros = False
            
            'Check if non-number found
            If InStr(sNumbers, s) < 1 Then
                IsNumberSuffix = False
            End If
            pos = pos + 1
        Wend
    End If
    If Not IsLetterPrefix Or Not IsNumberSuffix Or IsAllZeros Then IsExcelCellRef = False
End Function
'-------------------------------------------------------------------------------------
' True (and modify ByRef sht arg) if sheet exists but name has case difference
'   This function deals with the VBA/Excel oddity that Worksheet names are case
'   sensitive for testing their existence, but wkbk can't have two worksheets with
'   same names except for case. The function tests for case differences (usually
'   appropriate after first testing that SheetExists = False) and returns modified
'   sht argument to identically match Worksheet name --Boolean can be trigger user
'   warning and is test of whether to add a missing Worksheet
'
'   JDL 1/20/22; Modified 8/13/25 docstring
'
Function IsShtCaseErr(wkbk, ByRef sht) As Boolean
    Dim w As Variant
    For Each w In wkbk.Sheets
        If (LCase(w.Name) = LCase(sht)) And (w.Name <> sht) Then
            IsShtCaseErr = True
            sht = w.Name
        End If
    Next w
End Function
'-------------------------------------------------------------------------------------
' Add a new worksheet if doesn't already exist in wkbk
' Modified: 12/3/21 Optional shtAfter; 10/1/25 cleanup
'
Sub AddSheet(wkbk, shtNew, Optional shtAfter)
    Dim wkshtNew As Worksheet

    If wkbk Is Nothing Then Exit Sub
    If SheetExists(wkbk, shtNew) Then Exit Sub
    
    'Set shtAfter location if not specified
    If IsMissing(shtAfter) Then shtAfter = wkbk.Sheets(wkbk.Sheets.Count).Name
    If Not SheetExists(wkbk, shtAfter) Then shtAfter = _
        wkbk.Sheets(wkbk.Sheets.Count).Name
    Set wkshtNew = wkbk.Sheets.Add(After:=wkbk.Sheets(shtAfter))
    wkshtNew.Name = shtNew
End Sub
'-------------------------------------------------------------------------------------
' Delete specified sheet if it exists
'
Sub DeleteSheet(wkbk, sht)
    Application.DisplayAlerts = False
    If SheetExists(wkbk, sht) Then wkbk.Sheets(sht).Delete
End Sub
'-------------------------------------------------------------------------------------
' Create an Excel name, sName, using sNameString as RefersTo
'
Sub MakeXLName(wkbk, sName, sNameString)
    'DeleteXLName wkbk, sName
    wkbk.Names.Add Name:=sName, RefersToR1C1:=sNameString
End Sub
'-------------------------------------------------------------------------------------
' Delete an Excel name from a workbook
'
Sub DeleteXLName(wkbk, sName)
Dim w As Variant
    For Each w In wkbk.Names
        If w.Name = sName Then
            wkbk.Names(w.Name).Delete
            Exit Sub
        End If
    Next w
End Sub
'-------------------------------------------------------------------------------------
' Delete all workbook names
' .Visible property is skips hidden _xlfn.SINGLE name created by dynamic array
' glitch from circa 2020 Excel. See: https://stackoverflow.com/questions/59121799
'
' JDL 11/22/21
'
Sub DeleteAllWorkbookNames(wkbk)
    Dim w As Variant
    For Each w In wkbk.Names
        If w.Visible Then w.Delete
    Next w
End Sub
'-------------------------------------------------------------------------------------
' Delete Unused names from workbook
'
Sub DeleteUnusedNames(wkbk)
    Dim w As Variant
    For Each w In wkbk.Names
        If InStr(w.RefersToR1C1, "#REF!") > 0 Then w.Delete
    Next w
End Sub
'-------------------------------------------------------------------------------------
' True if named range exists in the workbook
'
Function NameExists(wkbk, sName) As Boolean
Dim w As Variant
    NameExists = False
    For Each w In wkbk.Names
        If UCase(w.Name) = UCase(sName) Then
            NameExists = True
            Exit Function
        End If
    Next w
End Function
'-------------------------------------------------------------------------------------
' Determine whether specified value is in array
' JDL 1/7/21 - replace verbose Modified: 3/3/21 handle non-array; 10/1/25 cleanup
'
Public Function IsInAry(ary As Variant, val As Variant) As Boolean

    'Exit if empty array)
    If UBound(ary) = -1 Then Exit Function

    IsInAry = Not IsError(Application.Match(val, ary, 0))
End Function
'-------------------------------------------------------------------------------------
' Set a table row range (multirange) based on array of key values
' Notes:     Mimics Pandas .loc functionality with multiindex
'            Suitable for small tables - Works well up to 500 found items in 1000 row table
'            Bogs down with large n found items (e.g. 2500 found rows in 5000 row table)
' Inputs:    tbl [table Class instance]
'            aryKeyColRanges [array of Ranges] table Class column ranges for key columns
'            aryKeyValues [array of values (variant)] array of key column values
' Validated in Unique Vals_1220.xlsm; validated 4-key 1/27/22 JDL val_Scenario Model.xlsm
'
'Created:   12/17/20 JDL      Modified: 2/17/21
'
'
Function rngMultiKey(tbl, aryKeyColRanges, aryKeyValues) As Range
    Dim i As Integer, rngCurrent As Range, rngSearch As Range, rngFound As Range

    'Progressively search/subset across the key columns -- start with entire table
    Set rngCurrent = tbl.rngRows
    For i = LBound(aryKeyColRanges) To UBound(aryKeyColRanges)

        'Stop if previous search returned Nothing
        If Not rngCurrent Is Nothing Then
            Set rngSearch = Intersect(rngCurrent, aryKeyColRanges(i))
            Set rngCurrent = Nothing
            On Error Resume Next 'in case nothing
            Set rngCurrent = FindAll(rngSearch, aryKeyValues(i)).EntireRow
            On Error GoTo 0
        End If
    Next i
    If Not rngCurrent Is Nothing Then Set rngMultiKey = rngCurrent.EntireRow
End Function
'-------------------------------------------------------------------------------------
' Set a table row range (multirange) based on array of keys (non tblRowsCols Class version)
' Notes:    Mimics Pandas .loc functionality with multiindex
'         Suitable for small tables - Works well up to 500 found items in 1000 row table
'         Bogs down with large number of found items (e.g. 2500 found rows in 5000 row table)
' Inputs    rngRows [Range] range of data rows in table (entire rows)
'         aryKeyColRanges [array of Ranges] column ranges for key columns
'         aryKeyValues [array of values (variant)] array of key column values
'
' Created:   12/17/20 JDL      Modified: 12/21/22 non tblRowsCols version for ErrorHandling
'
Function rngMultiKeyBasic(rngRows, aryKeyColRanges, aryKeyValues) As Range
    Dim i As Integer, rngCurrent As Range, rngSearch As Range, rngFound As Range

    'Progressively search/subset across the key columns -- start with entire table
    Set rngCurrent = rngRows
    For i = LBound(aryKeyColRanges) To UBound(aryKeyColRanges)

        'Stop if previous search returned Nothing
        If Not rngCurrent Is Nothing Then
            Set rngSearch = Intersect(rngCurrent, aryKeyColRanges(i))
            Set rngCurrent = Nothing
            On Error Resume Next 'in case nothing
            Set rngCurrent = FindAll(rngSearch, aryKeyValues(i)).EntireRow
            On Error GoTo 0
        End If
    Next i
    If Not rngCurrent Is Nothing Then Set rngMultiKeyBasic = rngCurrent.EntireRow
End Function
'-------------------------------------------------------------------------------------
' True if sheet has column outline
' JDL 9/21/22; 10/1/25 cleanup
'
Function HasColOutlining(wkbk, sht) As Boolean
    Dim c As Range
    HasColOutlining = False
    For Each c In wkbk.Sheets(sht).UsedRange.Rows(1).EntireColumn
        If c.OutlineLevel > 1 Then
            HasColOutlining = True
            Exit Function
        End If
    Next c
End Function
'-------------------------------------------------------------------------------------
'True if specified key is in Collection
'JDL 12/16/22; 10/1/25 cleanup
'
Public Function IsKeyInColl(coll, key) As Boolean
Dim v As Variant
On Error GoTo err
    IsKeyInColl = True
    v = coll(key)
    Exit Function
err:
    IsKeyInColl = False
End Function
'-------------------------------------------------------------------------------------
' Check if formula string contains syntax error by inserting formula into a check cell
' JDL 12/16/22; 10/1/25 cleanup
'
Function IsValidFormulaSyntax(rngCheckCell, sFormula)
    IsValidFormulaSyntax = True
    On Error Resume Next
    With rngCheckCell
        .Formula = sFormula
        If err.Number = 1004 Then IsValidFormulaSyntax = False
        .ClearContents
    End With
End Function
'-----------------------------------------------------------------------------
' Set Application environment for testing
' JDL 1/4/22 (added 10/7/24)
Sub SetApplEnvir(IsEvents, IsScreenUpdate, xlCalc)
    With Application
        .EnableEvents = IsEvents
        .ScreenUpdating = IsScreenUpdate
        .Calculation = xlCalc
    End With
End Sub
'-----------------------------------------------------------------------------------------
' Find end of contiguous populated range in a direction (used in client proj 8/25)
' JDL 6/10/25
'
Public Function FindEndPopulatedRange(cStart As Range, direct As xlDirection) As Range
    Dim r As Long, c As Long

    With cStart.Worksheet
        Select Case direct
            Case xlToRight
                c = cStart.Column
                Do While .Cells(cStart.Row, c).value <> ""
                    c = c + 1
                Loop
                Set FindEndPopulatedRange = .Cells(cStart.Row, c - 1)
            Case xlDown
                r = cStart.Row
                Do While .Cells(r, cStart.Column).value <> ""
                    r = r + 1
                Loop
                Set FindEndPopulatedRange = .Cells(r - 1, cStart.Column)
            Case Else
                Set FindEndPopulatedRange = cStart
        End Select
    End With
End Function

'-----------------------------------------------------------------------------------------
' Delete all visible range names that start with specified prefix
' JDL 7/8/25
'
Public Function DeleteRngNamesWithPrefix(wkbk, prefix) As Boolean
    SetErrs DeleteRngNamesWithPrefix: If errs.IsHandle Then On Error GoTo ErrorExit
    Dim nm As Name, i As Long, iLenPrefix As Long
    
    ' Loop through wkbk range names
    iLenPrefix = Len(prefix)
    For i = wkbk.Names.Count To 1 Step -1
        Set nm = wkbk.Names(i)
        
        ' Delete if name starts with prefix and is visible (not hidden)
        If Left(nm.Name, iLenPrefix) = prefix And nm.Visible Then nm.Delete
    Next i
    Exit Function
    
ErrorExit:
    errs.RecordErr "DeleteRngNamesWithPrefix", DeleteRngNamesWithPrefix
End Function
'-------------------------------------------------------------------------------------
' Save workbook with overwrite capability and optional close
' JDL 8/4/25; Updated 8/15/25 add file format setting
'
Public Function SaveAsCloseOverwrite(ByVal wkbk As Workbook, _
            ByVal filepath As String, Optional IsSave As Boolean = True, _
            Optional IsClose As Boolean = True) As Boolean
    SetErrs SaveAsCloseOverwrite: If errs.IsHandle Then On Error GoTo ErrorExit
    Dim fileFormat As Long
    
    If wkbk Is Nothing Then Exit Function
        
    ' If saving, disable alerts to prevent overwrite prompt
    If IsSave Then
        Application.DisplayAlerts = False
        
        ' Determine file format based on extension for cross-platform compatibility
        If InStr(LCase(filepath), ".xlsx") > 0 Then
            fileFormat = 51 ' xlOpenXMLWorkbook
        ElseIf InStr(LCase(filepath), ".xlsm") > 0 Then
            fileFormat = 52 ' xlOpenXMLWorkbookMacroEnabled
        ElseIf InStr(LCase(filepath), ".xls") > 0 Then
            fileFormat = -4143 ' xlWorkbookNormal
        ElseIf InStr(LCase(filepath), ".csv") > 0 Then
            fileFormat = 6 ' xlCSV
        Else
            fileFormat = 51 ' Default to xlsx format
        End If
        
        wkbk.SaveAs filepath, fileFormat:=fileFormat
        Application.DisplayAlerts = True
    End If
    
    If IsClose Then
        wkbk.Close False
        Set wkbk = Nothing
    End If
    Exit Function
    
ErrorExit:
    ' Ensure alerts re-enabled if error occurs
    Application.DisplayAlerts = True
    errs.RecordErr "SaveAsCloseOverwrite", SaveAsCloseOverwrite
End Function
'-------------------------------------------------------------------------------------
' Converts path separators in pathname to OS-appropriate separator
' JDL 8/13/25
'
Public Function CreatePathForOS(pathname As String) As Boolean
    SetErrs CreatePathForOS: If errs.IsHandle Then On Error GoTo ErrorExit
    Dim i As Integer, currentChar As String, resultPath As String
    
    ' Handle empty string
    If Len(pathname) = 0 Then
        pathname = ""
        Exit Function
    End If
    
    ' Replace all forward slashes and backslashes with OS path separator
    resultPath = pathname
    resultPath = Replace(resultPath, "/", Application.PathSeparator)
    resultPath = Replace(resultPath, "\", Application.PathSeparator)
    
    ' Update the original ByRef pathname parameter
    pathname = resultPath
    Exit Function
    
ErrorExit:
    errs.RecordErr "CreatePathForOS", CreatePathForOS
End Function
'-------------------------------------------------------------------------------------
' Converts path separators in pathname to OS-appropriate separator
' Cross-platform file picker.
' Returns a full path as a String, or "" if the user cancels.
' On Windows -> normal path. On Mac -> POSIX path (e.g., /Users/you/Desktop/file.txt).
' JDL 6/25


Public Function PickFile(Optional ByVal prompt As String = "Choose a fileƒ", _
    Optional ByVal initialFolder As String = "") As String

    #If Mac Then
        ' --- macOS path via AppleScriptTask ---
        ' IMPORTANT: You must install the AppleScript shown below as:
        '   ~/Library/Application Scripts/com.microsoft.Excel/ExcelFileDialogs.scpt
        '
        ' We pass "prompt|initialFolder" as one string; the script splits it.
        Dim arg As String
        arg = prompt & "|" & initialFolder
        On Error Resume Next
        PickFile = AppleScriptTask("ExcelFileDialogs.scpt", "chooseFile", arg)
        On Error GoTo 0
    
    #Else
        Dim fd As Object
        Set fd = Application.FileDialog(msoFileDialogFilePicker)
    
        With fd
            .Title = prompt
            .AllowMultiSelect = False
    
            ' Optional: show all files
            .Filters.Clear
            .Filters.Add "All Files", "*.*"
    
            ' Seed the starting folder (trailing slash ok/not required)
            If Len(initialFolder) > 0 Then
                Dim startPath As String
                startPath = initialFolder
                If Right$(startPath, 1) <> "\" And Right$(startPath, 1) <> "/" Then
                    startPath = startPath & Application.PathSeparator
                End If
                .InitialFileName = startPath
            End If
    
            If .Show = -1 Then
                PickFile = .SelectedItems(1)
            Else
                PickFile = ""
            End If
        End With
    
        Set fd = Nothing
    #End If
End Function
' Demo sub for PickFile
Public Sub DemoPickFile()
    Dim p As String
    p = PickFile("Select a file to import", ThisWorkbook.Path & Application.PathSeparator)
    If Len(p) = 0 Then
        MsgBox "Cancelled.", vbInformation
    Else
        MsgBox "You picked:" & vbCrLf & p, vbInformation
    End If
End Sub
'-------------------------------------------------------------------------------------
' Fill Right values in a single row range until next populated cell in each segment
' Leaves initial blank cells blank and only fills between populated cells
' JDL 9/30/25
'
Public Function FillRightBySegments(rowRng As Range) As Boolean
    SetErrs FillRightBySegments: If errs.IsHandle Then On Error GoTo ErrorExit
    Dim cell As Range, startCell As Range, lastValue As Variant
    
    ' Validate input - must be a single row
    If errs.IsFail(rowRng.Rows.Count <> 1, 1) Then GoTo ErrorExit
    
    ' Skip if range is empty or has only one cell
    If rowRng.Columns.Count <= 1 Then Exit Function
    If WorksheetFunction.CountA(rowRng) = 0 Then Exit Function
    
    ' Find first populated cell to start from
    Set startCell = Nothing
    For Each cell In rowRng.Cells
        If Not IsEmpty(cell.Value2) Then
            Set startCell = cell
            lastValue = cell.Value2
            Exit For
        End If
    Next cell
    
    ' Exit if no populated cells found
    If startCell Is Nothing Then Exit Function
    
    ' Fill Right only until next populated cell in each segment
    For Each cell In rowRng.Cells
        With cell
            ' Found next populated cell - fill from startCell to cell before this one
            If Not IsEmpty(.Value2) And .Address <> startCell.Address Then
                If .Column > startCell.Column + 1 Then _
                    startCell.Resize(1, .Column - startCell.Column).FillRight
                Set startCell = cell
                lastValue = .Value2
            
            ' Last cell in range - fill remaining if needed
            ElseIf .Address = rowRng.Cells(1, rowRng.Columns.Count).Address Then
                If IsEmpty(.Value2) And .Column > startCell.Column Then _
                    startCell.Resize(1, .Column - startCell.Column + 1).FillRight
            End If
        End With
    Next cell
    Exit Function
    
ErrorExit:
    errs.RecordErr "FillRightBySegments", FillRightBySegments
End Function
'-------------------------------------------------------------------------------------
' Convert literal "vbCrLf" text in string to actual line break
' JDL 10/16/25
'
Public Function ConvertVbCrLfToConcat(ByVal sInput As String) As String
    Dim sResult As String
    
    ' Replace all occurrences of " vbCrLf " with actual line break
    sResult = Replace(sInput, " vbCrLf ", vbCrLf)
    
    ConvertVbCrLfToConcat = sResult
End Function '-------------------------------------------------------------------------------------
'-------------------------------------------------------------------------------------
' SafeSaveAs
'
' Needed for Mac Excel to save files into iCloud folders where need to
' establish write permission before saving. Normal wb.SaveAs works with no issues in
' Windows Excel and in Mac Excel saving to local folders
'-------------------------------------------------------------------------------------
'-------------------------------------------------------------------------------------
' SafeSaveAs: SaveCopyAs with iCloud access handling
'
' wkbk - Workbook to save
' pf - Full path including filename
'
' JDL 10/14/25
'
Public Function SafeSaveAs(wkbk As Workbook, ByVal pf As String) As Boolean
    SetErrs SafeSaveAs: If errs.IsHandle Then On Error GoTo ErrorExit
    Dim folder As String, fname As String, sep As String, isMac As Boolean
    Dim isICloud As Boolean
    
    If Not SetSep(sep) Then GoTo ErrorExit
    If Not SetFolderFile(pf, folder, fname, sep) Then GoTo ErrorExit
    If errs.IsFail(Len(Dir(folder, vbDirectory)) = 0, 1, pf) Then GoTo ErrorExit
    
    ' Detect platform and cloud storage
    isMac = (InStr(1, Application.OperatingSystem, "Macintosh", vbTextCompare) > 0)
    isICloud = IsICloudPath(pf)
    
    ' Mac + iCloud: ensure we have write access (Can migrate err.Raise to use errs.IsFail)
    If isMac And isICloud Then
        If errs.IsFail(Not EnsureFolderAccess(folder), 1, pf) Then GoTo ErrorExit
    End If
    
    ' Save the file
    Application.DisplayAlerts = False
    wkbk.SaveCopyAs fileName:=pf
    Application.DisplayAlerts = True
    Exit Function
    
ErrorExit:
    Application.DisplayAlerts = True
    errs.RecordErr "SafeSaveAs", SafeSaveAs
End Function
'-------------------------------------------------------------------------------------
' Get platform-appropriate path separator
' JDL 10/14/25
'
Private Function SetSep(ByRef sep As String) As Boolean
    SetErrs SetSep: If errs.IsHandle Then On Error GoTo ErrorExit
    
    sep = Application.PathSeparator
    If InStr(1, Application.OperatingSystem, "Macintosh", vbTextCompare) > 0 Then _
        sep = "/"
    Exit Function
    
ErrorExit:
    errs.RecordErr "SetSep", SetSep
End Function
'-------------------------------------------------------------------------------------
' Split path into folder and filename
' JDL 10/14/25
'
Private Function SetFolderFile(ByVal pf As String, ByRef folder As String, _
    ByRef fname As String, ByVal sep As String) As Boolean
    SetErrs SetFolderFile: If errs.IsHandle Then On Error GoTo ErrorExit
    Dim pos As Long
    
    pos = InStrRev(pf, sep)
    If errs.IsFail(pos = 0, 1, pf) Then GoTo ErrorExit
    
    folder = Left$(pf, pos - 1)
    fname = Mid$(pf, pos + 1)
    
    If errs.IsFail(Len(folder) = 0 Or Len(fname) = 0, 2, pf) Then GoTo ErrorExit
    Exit Function
    
ErrorExit:
    errs.RecordErr "SetFolderFile", SetFolderFile
End Function
'-------------------------------------------------------------------------------------
' Ensure write access to folder, requesting permission if needed
' JDL 10/14/25; Modified 10/24/25
'
Private Function EnsureFolderAccess(ByVal folder As String) As Boolean
    SetErrs EnsureFolderAccess: If errs.IsHandle Then On Error GoTo ErrorExit
    
    ' Check if we already have access
    If CanWriteToFolder(folder) Then Exit Function
        
    'File access error occurred (err.number = 75)
    If err.Number = 75 Then
        'Debug.Print err.Number
        err.Clear
        
    'Other error occurred and needs errs to handle (bypasses message)
    ElseIf err.Number <> 0 And errs.IsHandle Then
        GoTo ErrorExit
    End If
    
    ' Need to request access
    Call errs.ShowMessage("EnsureFolderAccess", 1, vbInformation)
    
    If Not RequestFolderAccess(folder) Then GoTo ErrorExit
    Exit Function
    
ErrorExit:
    errs.RecordErr "EnsureFolderAccess", EnsureFolderAccess
End Function
'-------------------------------------------------------------------------------------
' Test if we can write to a folder (grants sandbox access)
' JDL 10/14/25; Modified 10/24/25
'
Private Function CanWriteToFolder(ByVal folder As String) As Boolean
    SetErrs CanWriteToFolder: If errs.IsHandle Then On Error GoTo ErrorExit
    Dim ff As Integer, probe As String, oldProbe As String
    
    oldProbe = Dir(folder & "/excel_probe*.tmp")
    Do While Len(oldProbe) > 0
        Kill folder & "/" & oldProbe
        oldProbe = Dir()
    Loop
    
    ' Now create new probe file
    probe = folder & "/excel_probe_" & Format(Now, "yyyymmdd_hhnnss") & ".tmp"
    ff = FreeFile
    
    'Trap file access error and attempt to write to probe file
    On Error Resume Next
    Open probe For Output As #ff
    
    If err.Number = 75 Then
        CanWriteToFolder = False
        Exit Function
    ElseIf err.Number <> 0 Then
        GoTo ErrorExit
    End If
    
    'Reset to default error handling
    On Error GoTo 0
    If errs.IsHandle Then On Error GoTo ErrorExit
    
    'Write to probe file, close and delete
    Print #ff, "ok"
    Close #ff
    Kill probe
    Exit Function
    
ErrorExit:
    On Error Resume Next
    Close #ff
    Kill probe
    On Error GoTo 0
    errs.RecordErr "CanWriteToFolder", CanWriteToFolder
End Function
'-------------------------------------------------------------------------------------
' Show native folder picker to grant access
' JDL 10/14/25
'
Private Function RequestFolderAccess(ByVal folder As String) As Boolean
    SetErrs RequestFolderAccess: If errs.IsHandle Then On Error GoTo ErrorExit
    Dim selectedFolder As Variant, dlg As FileDialog
    
    #If Mac Then
        selectedFolder = MacScript("POSIX path of (choose folder with prompt " & _
            """Grant Excel access to this iCloud folder:"" default location """ & _
            folder & """)")
    #Else
        Set dlg = Application.FileDialog(msoFileDialogFolderPicker)
        With dlg
            .Title = "Grant Excel access to this folder"
            .InitialFileName = folder
            If errs.IsFail(.Show <> -1, 1) Then GoTo ErrorExit
        End With
    #End If
    Exit Function
    
ErrorExit:
    errs.RecordErr "RequestFolderAccess", RequestFolderAccess
End Function

'-------------------------------------------------------------------------------------
' Detect iCloud Drive paths
' JDL 10/14/25
'
Private Function IsICloudPath(ByVal pf As String) As Boolean
    IsICloudPath = (InStr(1, pf, "/Library/Mobile Documents/", vbTextCompare) > 0)
End Function