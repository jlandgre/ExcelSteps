Attribute VB_Name = "Utilities_Testing"
'Version 7/26/23 - add AryFromRowRng, AryFromColRng functions
Option Explicit
Public Const bErrorHandle As Boolean = False
'-----------------------------------------------------------------------------------------------
' Sub wkbkResetStatus
'
Sub wkbkResetStatus(bRefreshStart, wkbk, xCalculation, shtCurrent, shtRC, ZoomSetting)
    
    Application.ScreenUpdating = Not bRefreshStart
    If bRefreshStart Then
        xCalculation = Application.Calculation
        shtCurrent = wkbk.ActiveSheet.Name
        ZoomSetting = ActiveWindow.Zoom
        If SheetExists(wkbk, shtRC) Then wkbk.Sheets(shtRC).Activate
        Application.Calculation = xlCalculationManual
    Else
        Application.Calculation = xCalculation
        If SheetExists(wkbk, shtCurrent) Then wkbk.Sheets(shtCurrent).Activate
        ActiveWindow.Zoom = ZoomSetting
    End If
End Sub
'-----------------------------------------------------------------------------------------------
' Set ScreenUpdating and EnableEvents status
'
Sub wkbkResetSimple(IsStart)
    With Application
        .ScreenUpdating = Not IsStart
        .EnableEvents = Not IsStart
    End With
End Sub

'-----------------------------------------------------------------------------------------------
' RngColWidthFormat
'
Sub RngColWidthFormat(ByVal rngCell As Range)

    'Exit if column is empty
    If WorksheetFunction.CountA(rngCell.EntireColumn) = 0 Then Exit Sub
    
    With rngCell.EntireColumn
        .ColumnWidth = 120
        .AutoFit
        If .ColumnWidth < 120 Then .ColumnWidth = .ColumnWidth + 2
        
    End With
End Sub
'-----------------------------------------------------------------------------
'Purpose: Build a comma-separated list of a column range's contents
'
'Created: 11/23/21 JDL
'
Function LstColContents(wkbk, sht, iCol)
    Dim rng As Range
    With wkbk.Sheets(sht)
        Set rng = Range(.Cells(1, iCol), rngLastPopCell(.Cells(1, iCol), xlDown))
    LstColContents = ListFromArray(rng)
    End With
End Function
'-----------------------------------------------------------------------------------------------
'Purpose:   Determine whether specified range is a row range
'
Function IsRowRange(rng) As Boolean
    IsRowRange = rng.Address = rng.EntireRow.Address
End Function
'-----------------------------------------------------------------------------------------------
'Purpose:   Determine whether specified range is a column range
'
Function IsColRng(rng) As Boolean
    IsColRng = rng.Address = rng.EntireColumn.Address
End Function
'-----------------------------------------------------------------------------------------------
'Purpose:   Determine whether specified range is cell or block of cells (not column or row)
'
Function IsCellRng(rng) As Boolean
    IsCellRng = Not IsRowRange(rng) And Not IsColRng(rng)
End Function
'-----------------------------------------------------------------------------------------------
Sub ShadeYellow(rng)
    rng.Interior.Pattern = xlSolid
    rng.Interior.Color = 65535
End Sub
'-----------------------------------------------------------------------------------------------
'Purpose:  Reveal all worksheet cells by clearing outline, turning off filter and unhiding
'
'Inputs:    wksht [Worksheet] Worksheet to reveal
'
'Created:   12/1/21 JDL      Modified:
'
Sub RevealWkshtCells(wksht)
    With wksht
        .AutoFilterMode = False
        .Cells.ClearOutline
        .Cells.EntireColumn.Hidden = False
        .Cells.EntireRow.Hidden = False
    End With
End Sub
'-----------------------------------------------------------------------------
'Purpose: Set Application properties for testing
'
'Created:   1/4/22 JDL      Modified:
'
Sub SetApplEnvir(IsEvents, IsScreenUpdate, xlCalc)
    With Application
        .EnableEvents = IsEvents
        .ScreenUpdating = IsScreenUpdate
        .Calculation = xlCalc
    End With
End Sub
'-----------------------------------------------------------------------------------------------
'Purpose: Count and delete stray worksheets if they are empty (Sheet1 etc.)
'
'Modified: 1/20/22 Used for testing tblRowsCols
'
Function iCountAndDeleteStraySheets(wkbk) As Integer
    Dim wksht As Variant
    iCountAndDeleteStraySheets = 0
    For Each wksht In wkbk.Sheets
        If LCase(Left(wksht.Name, 5)) = "sheet" Then
            If wksht.UsedRange.Address = "$A$1" And IsEmpty(wksht.Cells(1, 1)) Then
                DeleteSheet wkbk, wksht.Name
                iCountAndDeleteStraySheets = iCountAndDeleteStraySheets + 1
            End If
        End If
    Next wksht
End Function

'-----------------------------------------------------------------------------------------------
'Purpose:   Test whether a range was deleted and is in "Object Required" state
'
'Inputs:    rng [Range] range to test
'
'Created:   9/15/21 JDL      Modified:
'
'See DGlancy answer here:
'https://stackoverflow.com/questions/12127311/vba-what-happens-to-range-objects-if-user-deletes-cells
'
Function IsRngDeleted(rng) As Boolean
    Dim sAddress As String
    IsRngDeleted = False
    On Error Resume Next
    sAddress = rng.Address
    If Err.Number = 424 Then IsRngDeleted = True
End Function

'
'Function RangeExists - Determines whether named range exists in the workbook
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
'-----------------------------------------------------------------------------------------------
'Purpose: Append a specified val to end of specified array
'
'Inputs: aryOrig [Variant - Array] array to append onto
'        val [Variant] value to append
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
'-----------------------------------------------------------------------------------------------
'Purpose:   Populate Array of unique values from cells in specified range
'
'Inputs:    R [Range] range containing cells for unique values
'
'Created:   3/12/22
'
Function aryUniqueValsInRange(r)
    Dim aryUnique As Variant, w As Variant
    If r Is Nothing Then Exit Function
    aryUnique = Array()
    For Each w In r.Cells
        If Not IsEmpty(w) Then
            If Not IsInAry(aryUnique, w.Value) Then aryUnique = aryAppendVal(aryUnique, w.Value)
        End If
    Next w
    aryUniqueValsInRange = aryUnique
End Function

'-----------------------------------------------------------------------------------------------
' Create a uni-dimension array from a specified Range row
' JDL 7/26/23
Function AryFromRowRng(rngRow, iRow) As Variant
    Dim i As Integer, aryRow As Variant, aryTemp() As Variant
    aryRow = rngRow.Value
    ReDim aryTemp(0 To UBound(aryRow, 2) - 1)
    
    For i = LBound(aryRow, 2) To UBound(aryRow, 2)
        aryTemp(i - 1) = aryRow(1, i)
    Next i
    AryFromRowRng = aryTemp
End Function
'-----------------------------------------------------------------------------------------------
' Create a uni-dimensional array from a specified Range column
' JDL 7/26/23
Function AryFromColRng(rngcol, iCol) As Variant
    Dim i As Integer, aryCol As Variant, aryTemp() As Variant
    aryCol = rngcol.Value
    ReDim aryTemp(0 To UBound(aryCol, 1) - 1)
    
    For i = LBound(aryCol, 1) To UBound(aryCol, 1)
        aryTemp(i - 1) = aryCol(i, 1)
    Next i
    AryFromColRng = aryTemp
End Function

'-----------------------------------------------------------------------------------------------
' Lookup and return value at intersection of table cell row and column range
' Inputs:   rngCell [Range] single row range (usually single cell from key column)
'           rngCol [Range] table column range
'           iShift [Integer] optional shift row value
'
'Created: 6/20 JDL
'
Function TableLoc(rngCell, rngcol, Optional ishift = 0) As Variant
    TableLoc = Intersect(rngCell.Offset(ishift, 0).EntireRow, rngcol)
End Function
'-----------------------------------------------------------------------------------------------
' Set Add-in version comment -updates the add-in's Comments field to match internal Version constant
' JDL 12/16/25
Sub Set_ExcelStepsVersion()
    Dim wkbk As Workbook
    
    Set wkbk = Application.Workbooks("XLSteps.xlam")
    wkbk.BuiltinDocumentProperties("Comments").Value = ExcelSteps.Version
End Sub
'-----------------------------------------------------------------------------------------------
' Consolidate code_plan.csv and ExcelSteps_Code_Plan.xlsx into Code_Plan.xlsx
' JDL 3/9/26
'
Sub ExportCodePlanToExcel()
    Dim wkbkCodePlan As Workbook, wkbkCSV As Workbook, wkbkExisting As Workbook
    Dim pathDocs As String, pathCSV As String, pathExisting As String, pathOutput As String
    Dim sep As String, wksht As Worksheet, rngHeaders As Range
    Dim headers As Variant
    
    sep = Application.PathSeparator
    pathDocs = ThisWorkbook.Path & sep & ".." & sep & "docs"
    pathCSV = pathDocs & sep & "code_plan.csv"
    pathExisting = pathDocs & sep & "ExcelSteps_code_plan.xlsx"
    pathOutput = pathDocs & sep & "code_Plan.xlsx"
    
    ' Create new workbook
    Set wkbkCodePlan = Workbooks.Add
    
    ' Open and copy code_plan.csv if it exists
    If Len(Dir$(pathCSV)) > 0 Then
        If Not ExcelSteps.OpenFile(pathCSV, wkbkCSV) Then GoTo ErrorExit
        wkbkCSV.Sheets(1).Copy Before:=wkbkCodePlan.Sheets(1)
        wkbkCodePlan.Sheets(1).Name = "plan"
        wkbkCSV.Close SaveChanges:=False
    Else
        ' Create blank sheet with headers
        Set wksht = wkbkCodePlan.Sheets(1)
        wksht.Name = "plan"
        headers = Split("Module;Use_Case;Procedure;Method;Docstring;Arguments;" & _
                       "Code writing instructions;Testing Considerations", ";")
        Set rngHeaders = wksht.Range("A1").Resize(1, UBound(headers) + 1)
        rngHeaders.Value = headers
    End If
    
    ' Open and copy ExcelSteps sheet from ExcelSteps_Code_Plan.xlsx
    If Not ExcelSteps.OpenFile(pathExisting, wkbkExisting) Then GoTo ErrorExit
    wkbkExisting.Sheets("ExcelSteps").Copy After:=wkbkCodePlan.Sheets(wkbkCodePlan.Sheets.Count)
    wkbkExisting.Close SaveChanges:=False
    
    ' Delete default blank sheet if it exists
    Application.DisplayAlerts = False
    On Error Resume Next
    wkbkCodePlan.Sheets("Sheet1").Delete
    On Error GoTo ErrorExit
    Application.DisplayAlerts = True
    
    ' Save as Code_Plan.xlsx
    If Not ExcelSteps.SaveAsCloseOverwrite(wkbkCodePlan, pathOutput, _
                                           IsSave:=True, IsClose:=True) Then GoTo ErrorExit
    Exit Sub

ErrorExit:
    Application.DisplayAlerts = True
    MsgBox "Error ExportCodePlanToExcel"
End Sub

'-----------------------------------------------------------------------------------------------
' Split Code_Plan.xlsx into code_plan.csv and ExcelSteps_Code_Plan.xlsx
' JDL 3/9/26
'
Sub ExportCodePlanFromExcel()
    Dim wkbkCodePlan As Workbook, wkbkTmp As Workbook
    Dim pathDocs As String, pathCodePlan As String, pathCSV As String, pathExisting As String
    Dim sep As String

    sep = Application.PathSeparator
    pathDocs = ThisWorkbook.Path & sep & ".." & sep & "docs"
    pathCodePlan = pathDocs & sep & "code_plan.xlsx"
    pathCSV = pathDocs & sep & "code_plan.csv"
    pathExisting = pathDocs & sep & "ExcelSteps_code_plan.xlsx"

    If Not ExcelSteps.OpenFile(pathCodePlan, wkbkCodePlan) Then GoTo ErrorExit

    ' Export code_plan sheet to CSV by copying the sheet into a temp workbook.
    wkbkCodePlan.Sheets("plan").Copy
    Set wkbkTmp = ActiveWorkbook
    If Not ExcelSteps.SaveAsCloseOverwrite(wkbkTmp, pathCSV, _
                                           IsSave:=True, IsClose:=True) Then GoTo ErrorExit

    ' Export ExcelSteps sheet to a standalone workbook.
    wkbkCodePlan.Sheets("ExcelSteps").Copy
    Set wkbkTmp = ActiveWorkbook
    If Not ExcelSteps.SaveAsCloseOverwrite(wkbkTmp, pathExisting, _
                                           IsSave:=True, IsClose:=True) Then GoTo ErrorExit

    wkbkCodePlan.Close SaveChanges:=False
    Exit Sub

ErrorExit:
    On Error Resume Next
    If Not wkbkCodePlan Is Nothing Then wkbkCodePlan.Close SaveChanges:=False
    MsgBox "Error ExportCodePlanFromExcel"
End Sub


