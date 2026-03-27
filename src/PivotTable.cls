VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "PivotTable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Version 3/27/26
Option Explicit

Public tblSrc As Object
Public tblOut As Object
Public wkbkDest As Workbook
Public wkshtDest As Worksheet
Public shtDest As String
Public pvtCache As PivotCache
Public pvtTable As Object
Public rngTableOut As Range
Public rngDataOut As Range
'--------------------------------------------------------------------------------------
' Build a pivot table from rows/columns source data and convert to values output
' JDL 3/27/26
'
Public Function MakePivotTableProcedure(pvt, tblSrc, rowFields, colFields, shtDest, _
        analytes, Optional filters = vbNullString, _
        Optional sortOrder = vbNullString, _
        Optional isRowGrand As Variant = vbNullString, _
    Optional isColGrand As Variant = vbNullString, _
    Optional IsDestNewWorkbook As Boolean = False) As Boolean
    SetErrs MakePivotTableProcedure: If errs.IsHandle Then On Error GoTo ErrorExit

    If Not InitPivotTable(pvt, tblSrc, shtDest) Then GoTo ErrorExit
    If Not CreatePivotCacheAndTable(pvt, IsDestNewWorkbook) Then GoTo ErrorExit
    If Not ConfigurePivotFields(pvt, rowFields, colFields, analytes, filters) Then GoTo ErrorExit
    If Not ApplySortOrder(pvt, sortOrder) Then GoTo ErrorExit
    If Not FormatPivotTable(pvt, isRowGrand, isColGrand) Then GoTo ErrorExit
    If Not ConvertPivotToValues(pvt) Then GoTo ErrorExit
    If Not SetOutputRanges(pvt) Then GoTo ErrorExit
    Exit Function

ErrorExit:
    errs.RecordErr "MakePivotTableProcedure", MakePivotTableProcedure
End Function
'--------------------------------------------------------------------------------------
' Initialize attributes used by MakePivotTableProcedure
' JDL 3/27/26
'
Public Function InitPivotTable(pvt, tblSrc, shtDest) As Boolean
    SetErrs InitPivotTable: If errs.IsHandle Then On Error GoTo ErrorExit

    If errs.IsFail(tblSrc Is Nothing, 1) Then GoTo ErrorExit
    If errs.IsFail(tblSrc.rngTable Is Nothing, 2) Then GoTo ErrorExit

    Set pvt.tblSrc = tblSrc
    Set pvt.tblOut = New tblRowsCols

    pvt.shtDest = shtDest
    If Len(pvt.shtDest) < 1 Then pvt.shtDest = "PivotData"
    Exit Function

ErrorExit:
    errs.RecordErr "InitPivotTable", InitPivotTable
End Function
'--------------------------------------------------------------------------------------
' Create destination workbook and pivot cache/table
' JDL 3/27/26
'
Public Function CreatePivotCacheAndTable(pvt, Optional IsDestNewWorkbook As Boolean = False) As Boolean
    SetErrs CreatePivotCacheAndTable: If errs.IsHandle Then On Error GoTo ErrorExit
    Dim i As Long

    With pvt
        If IsDestNewWorkbook Then
            Set .wkbkDest = Application.Workbooks.Add
            Set .wkshtDest = .wkbkDest.Sheets(1)
            .wkshtDest.Name = .shtDest

            Application.DisplayAlerts = False
            For i = .wkbkDest.Sheets.Count To 2 Step -1
                .wkbkDest.Sheets(i).Delete
            Next i
            Application.DisplayAlerts = True
        Else
            Set .wkbkDest = .tblSrc.wkbk
            If errs.IsFail(.wkbkDest Is Nothing, 1) Then GoTo ErrorExit
            If errs.IsFail(LCase(.tblSrc.sht) = LCase(.shtDest), 2) Then GoTo ErrorExit

            Application.DisplayAlerts = False
            If SheetExists(.wkbkDest, .shtDest) Then .wkbkDest.Sheets(.shtDest).Delete
            Application.DisplayAlerts = True

            Set .wkshtDest = .wkbkDest.Worksheets.Add(After:=.wkbkDest.Sheets(.wkbkDest.Sheets.Count))
            .wkshtDest.Name = .shtDest
        End If

        Set .pvtCache = .wkbkDest.PivotCaches.Create(SourceType:=xlDatabase, _
            SourceData:=.tblSrc.rngTable)
        Set .pvtTable = .pvtCache.CreatePivotTable(TableDestination:=.wkshtDest.Cells(1, 1))
    End With
    Exit Function

ErrorExit:
    errs.RecordErr "CreatePivotCacheAndTable", CreatePivotCacheAndTable
End Function
'--------------------------------------------------------------------------------------
' Configure row, column, page and data fields
' JDL 3/27/26
'
Public Function ConfigurePivotFields(pvt, rowFields, colFields, analytes, filters) As Boolean
    SetErrs ConfigurePivotFields: If errs.IsHandle Then On Error GoTo ErrorExit

    If Not AddFieldGroup(pvt, rowFields, xlRowField) Then GoTo ErrorExit
    If HasFieldSpecs(colFields) Then
        If Not AddFieldGroup(pvt, colFields, xlColumnField) Then GoTo ErrorExit
    End If
    If Not AddFilterFields(pvt, filters) Then GoTo ErrorExit
    If Not AddAnalytes(pvt, analytes) Then GoTo ErrorExit
    Exit Function

ErrorExit:
    errs.RecordErr "ConfigurePivotFields", ConfigurePivotFields
End Function
'--------------------------------------------------------------------------------------
' Check whether a row/column field argument contains one or more non-blank field names
' JDL 3/27/26
'
Private Function HasFieldSpecs(fields) As Boolean
    Dim i As Long, j As Long
    Dim fieldSpec As Variant, fieldName As String

    HasFieldSpecs = False
    If IsArray(fields) Then
        For i = LBound(fields) To UBound(fields)
            fieldSpec = fields(i)
            If IsArray(fieldSpec) Then
                j = LBound(fieldSpec)
                fieldName = CStr(fieldSpec(j))
            Else
                fieldName = CStr(fieldSpec)
            End If

            If Len(Trim(fieldName)) > 0 Then
                HasFieldSpecs = True
                Exit Function
            End If
        Next i
        Exit Function
    End If

    If Len(Trim(CStr(fields))) > 0 Then HasFieldSpecs = True
End Function
'--------------------------------------------------------------------------------------
' Add row or column field group from string or array of strings
' JDL 3/27/26
'
Private Function AddFieldGroup(pvt, fields, ByVal fieldOrientation As XlPivotFieldOrientation) As Boolean
    SetErrs AddFieldGroup: If errs.IsHandle Then On Error GoTo ErrorExit
    Dim i As Long, j As Long, fieldName As String, nFields As Long
    Dim fieldSpec As Variant, isSubtotalsOn As Boolean, hasSubtotalSpec As Boolean
    Dim pvtField As PivotField

    If IsArray(fields) Then
        nFields = UBound(fields) - LBound(fields) + 1
        For i = LBound(fields) To UBound(fields)
            fieldSpec = fields(i)
            hasSubtotalSpec = False
            isSubtotalsOn = True

            If IsArray(fieldSpec) Then
                j = LBound(fieldSpec)
                fieldName = CStr(fieldSpec(j))
                If UBound(fieldSpec) >= j + 1 Then
                    hasSubtotalSpec = True
                    isSubtotalsOn = CBool(fieldSpec(j + 1))
                End If
            Else
                fieldName = CStr(fieldSpec)
            End If

            If Len(fieldName) < 1 Then GoTo NextField
            Set pvtField = pvt.pvtTable.PivotFields(fieldName)
            With pvtField
                .Orientation = fieldOrientation
                .Position = i - LBound(fields) + 1
            End With
            If hasSubtotalSpec Then
                If Not SetFieldSubtotals(pvtField, isSubtotalsOn) Then GoTo ErrorExit
            End If
NextField:
        Next i
    Else
        fieldName = CStr(fields)
        If Len(fieldName) < 1 Then GoTo ErrorExit
        Set pvtField = pvt.pvtTable.PivotFields(fieldName)
        With pvtField
            .Orientation = fieldOrientation
            .Position = 1
        End With
    End If
    Exit Function

ErrorExit:
    errs.RecordErr "AddFieldGroup", AddFieldGroup
End Function
'--------------------------------------------------------------------------------------
' Toggle pivot field subtotals by setting Automatic (index 1) on/off and clearing others
' JDL 3/27/26
'
Private Function SetFieldSubtotals(pvtField As PivotField, ByVal IsOn As Boolean) As Boolean
    SetErrs SetFieldSubtotals: If errs.IsHandle Then On Error GoTo ErrorExit
    Dim i As Long

    For i = 1 To 12
        pvtField.Subtotals(i) = False
    Next i
    If IsOn Then pvtField.Subtotals(1) = True
    Exit Function

ErrorExit:
    errs.RecordErr "SetFieldSubtotals", SetFieldSubtotals
End Function
'--------------------------------------------------------------------------------------
' Add optional page filter fields; each entry may be "FieldName" or Array(FieldName, Value)
' JDL 3/27/26
'
Private Function AddFilterFields(pvt, filters) As Boolean
    SetErrs AddFilterFields: If errs.IsHandle Then On Error GoTo ErrorExit
    Dim i As Long, nPos As Long
    Dim filterItem As Variant, fieldName As String, fieldValue As String

    If IsArray(filters) Then
        nPos = 1
        For i = LBound(filters) To UBound(filters)
            filterItem = filters(i)
            If IsArray(filterItem) Then
                fieldName = CStr(filterItem(LBound(filterItem)))
                If UBound(filterItem) >= LBound(filterItem) + 1 Then
                    fieldValue = CStr(filterItem(LBound(filterItem) + 1))
                Else
                    fieldValue = vbNullString
                End If
            Else
                fieldName = CStr(filterItem)
                fieldValue = vbNullString
            End If

            If Len(fieldName) < 1 Then GoTo NextFilter
            With pvt.pvtTable.PivotFields(fieldName)
                .Orientation = xlPageField
                .Position = nPos
                If Len(fieldValue) > 0 Then .CurrentPage = fieldValue
            End With
            nPos = nPos + 1
NextFilter:
        Next i
    End If
    Exit Function

ErrorExit:
    errs.RecordErr "AddFilterFields", AddFilterFields
End Function
'--------------------------------------------------------------------------------------
' Add analyte definitions to data area
' JDL 3/27/26
'
Private Function AddAnalytes(pvt, analytes) As Boolean
    SetErrs AddAnalytes: If errs.IsHandle Then On Error GoTo ErrorExit
    Dim i As Long, j As Long, fieldName As String, caption As String
    Dim xFunc As XlConsolidationFunction, analyteDef As Variant

    If errs.IsFail(Not IsArray(analytes), 1) Then GoTo ErrorExit

    For i = LBound(analytes) To UBound(analytes)
        analyteDef = analytes(i)
        xFunc = xlSum

        If IsArray(analyteDef) Then
            j = LBound(analyteDef)
            fieldName = CStr(analyteDef(j))
            If UBound(analyteDef) >= j + 1 Then xFunc = CLng(analyteDef(j + 1))
            If UBound(analyteDef) >= j + 2 Then
                caption = CStr(analyteDef(j + 2))
            Else
                caption = GetDefaultCaption(fieldName, xFunc)
            End If
        Else
            fieldName = CStr(analyteDef)
            caption = GetDefaultCaption(fieldName, xFunc)
        End If

        If errs.IsFail(Len(fieldName) < 1, 2) Then GoTo ErrorExit
        pvt.pvtTable.AddDataField pvt.pvtTable.PivotFields(fieldName), caption, xFunc
    Next i
    Exit Function

ErrorExit:
    errs.RecordErr "AddAnalytes", AddAnalytes
End Function
'--------------------------------------------------------------------------------------
' Return default data-field caption by consolidation function
' JDL 3/27/26
'
Private Function GetDefaultCaption(ByVal fieldName As String, _
        ByVal xFunc As XlConsolidationFunction) As String
    Select Case xFunc
        Case xlCount
            GetDefaultCaption = "Count of " & fieldName
        Case Else
            GetDefaultCaption = "Sum of " & fieldName
    End Select
End Function
'--------------------------------------------------------------------------------------
' Apply optional sort order to first column field
' JDL 3/27/26
'
Public Function ApplySortOrder(pvt, sortOrder) As Boolean
    SetErrs ApplySortOrder: If errs.IsHandle Then On Error GoTo ErrorExit
    Dim xOrder As XlSortOrder, sOrder As String
    Dim pvtField As PivotField, dataFieldName As String
    Dim nErr As Long

    sOrder = LCase(Trim(CStr(sortOrder)))
    If Len(sOrder) < 1 Then Exit Function

    Select Case sOrder
        Case "asc", "ascending"
            xOrder = xlAscending
        Case "desc", "descending"
            xOrder = xlDescending
        Case Else
            Exit Function
    End Select

    If pvt.pvtTable.ColumnFields.Count > 0 Then
        Set pvtField = pvt.pvtTable.ColumnFields(1)
        dataFieldName = vbNullString

        ' DataFields can throw when the pivot has no value field; skip sort in that case.
        On Error Resume Next
        dataFieldName = pvt.pvtTable.DataFields(1).Name
        nErr = Err.Number
        Err.Clear
        If errs.IsHandle Then
            On Error GoTo ErrorExit
        Else
            On Error GoTo 0
        End If

        If nErr <> 0 Or Len(dataFieldName) < 1 Then Exit Function
        pvtField.AutoSort xOrder, dataFieldName
    End If
    Exit Function

ErrorExit:
    errs.RecordErr "ApplySortOrder", ApplySortOrder
End Function
'--------------------------------------------------------------------------------------
' Apply standard formatting and optional grand-total settings
' JDL 3/27/26
'
Public Function FormatPivotTable(pvt, isRowGrand, isColGrand) As Boolean
    SetErrs FormatPivotTable: If errs.IsHandle Then On Error GoTo ErrorExit

    With pvt.pvtTable
        If Not IsEmpty(isRowGrand) And CStr(isRowGrand) <> vbNullString Then _
            .RowGrand = CBool(isRowGrand)
        If Not IsEmpty(isColGrand) And CStr(isColGrand) <> vbNullString Then _
            .ColumnGrand = CBool(isColGrand)
        .PreserveFormatting = True
        .RepeatAllLabels xlRepeatLabels
    End With

    If Not pvt.pvtTable.DataBodyRange Is Nothing Then _
        pvt.pvtTable.DataBodyRange.NumberFormat = "0.00"
    Exit Function

ErrorExit:
    errs.RecordErr "FormatPivotTable", FormatPivotTable
End Function
'--------------------------------------------------------------------------------------
' Freeze pivot results by converting UsedRange to values
' JDL 3/27/26
'
Public Function ConvertPivotToValues(pvt) As Boolean
    SetErrs ConvertPivotToValues: If errs.IsHandle Then On Error GoTo ErrorExit

    If Not pvt.pvtTable.DataBodyRange Is Nothing Then Set pvt.rngDataOut = pvt.pvtTable.DataBodyRange

    With pvt.wkshtDest.UsedRange
        .Copy
        .Cells(1, 1).PasteSpecial xlPasteValues
    End With
    Application.CutCopyMode = False
    Exit Function

ErrorExit:
    errs.RecordErr "ConvertPivotToValues", ConvertPivotToValues
End Function
'--------------------------------------------------------------------------------------
' Set output range pointers for downstream callers
' JDL 3/27/26
'
Public Function SetOutputRanges(pvt) As Boolean
    SetErrs SetOutputRanges: If errs.IsHandle Then On Error GoTo ErrorExit

    Set pvt.rngTableOut = pvt.wkshtDest.UsedRange
    If Not pvt.rngDataOut Is Nothing Then _
        Set pvt.rngDataOut = pvt.wkshtDest.Range(pvt.rngDataOut.Address)

    If Not pvt.tblOut.Init(pvt.tblOut, pvt.wkbkDest, pvt.shtDest) Then GoTo ErrorExit
    Exit Function

ErrorExit:
    errs.RecordErr "SetOutputRanges", SetOutputRanges
End Function
