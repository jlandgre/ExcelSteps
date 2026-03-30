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
Public IsDestNewWorkbook As Boolean
Public dictSubT As Object
Public dictParams As Object
Public rowFields As String
Public colFields As String
'--------------------------------------------------------------------------------------
' Build a pivot table from rows/columns source data and convert to values output
' JDL 3/27/26
'
Public Function MakePivotTableProcedure(pvt, tblSrc, rowFields, colFields, shtDest, _
        analytes, Optional filters = vbNullString, _
        Optional sortOrder = vbNullString, _
        Optional dictSubT As Object = Nothing, _
        Optional dictParams As Object = Nothing) As Boolean
    SetErrs MakePivotTableProcedure: If errs.IsHandle Then On Error GoTo ErrorExit

    If Not InitPivotTable(pvt, tblSrc, shtDest, dictSubT, dictParams) Then GoTo ErrorExit
    If Not CreatePivotCacheAndTable(pvt) Then GoTo ErrorExit
    If Not ValidateFieldSpecs(pvt, rowFields, colFields) Then GoTo ErrorExit
    If Not ValidateAnalytes(pvt, analytes) Then GoTo ErrorExit
    If Not ConfigurePivotFields(pvt, analytes, filters) Then GoTo ErrorExit
    If Not ApplySortOrder(pvt, sortOrder) Then GoTo ErrorExit
    If Not FormatPivotTable(pvt) Then GoTo ErrorExit
    ' If Not ConvertPivotToValues(pvt) Then GoTo ErrorExit
    ' If Not SetOutputRanges(pvt) Then GoTo ErrorExit
    Exit Function

ErrorExit:
    errs.RecordErr "MakePivotTableProcedure", MakePivotTableProcedure
End Function
'--------------------------------------------------------------------------------------
' Initialize attributes used by MakePivotTableProcedure
' JDL 3/27/26
'
Public Function InitPivotTable(pvt, tblSrc, shtDest, _
    Optional dictSubT As Object = Nothing, _
    Optional dictParams As Object = Nothing) As Boolean
    SetErrs InitPivotTable: If errs.IsHandle Then On Error GoTo ErrorExit

    If errs.IsFail(tblSrc Is Nothing, 1) Then GoTo ErrorExit
    If errs.IsFail(tblSrc.rngTable Is Nothing, 2) Then GoTo ErrorExit

    With pvt
        Set .tblSrc = tblSrc
        Set .tblOut = New tblRowsCols
        Set .dictSubT = New Dictionary
        Set .dictParams = New Dictionary
        If Not dictSubT Is Nothing Then Set .dictSubT = dictSubT
        If Not dictParams Is Nothing Then Set .dictParams = dictParams

        ' Set attribute for whether pivot table is in new workbook (default False)
        .IsDestNewWorkbook = False
        If .dictParams.Exists("IsDestNewWorkbook") Then _
            .IsDestNewWorkbook = CBool(.dictParams.Item("IsDestNewWorkbook"))

        ' Set Pivot table sheet name
        .shtDest = shtDest
        If Len(.shtDest) < 1 Then .shtDest = "PivotData"

        ' For same-workbook output, destination sheet must differ from source sheet.
        If Not .IsDestNewWorkbook Then _
            If errs.IsFail(LCase(.tblSrc.sht) = LCase(.shtDest), 3) Then GoTo ErrorExit
    End With
    Exit Function

ErrorExit:
    errs.RecordErr "InitPivotTable", InitPivotTable
End Function

'--------------------------------------------------------------------------------------
' Create destination workbook and pivot cache/table
' JDL 3/27/26
'
Public Function CreatePivotCacheAndTable(pvt) As Boolean
    SetErrs CreatePivotCacheAndTable: If errs.IsHandle Then On Error GoTo ErrorExit

    With pvt
        ' Destination workbook is different than source workbook
        If .IsDestNewWorkbook Then
            Set .wkbkDest = Application.Workbooks.Add
            Set .wkshtDest = .wkbkDest.Sheets(1)
            If Not DeleteExtraneousShts(.wkbkDest, .shtDest) Then GoTo ErrorExit

        'Destination sheet is in same workbook as source sheet
        Else
            Set .wkbkDest = .tblSrc.wkbk

            ' Delete existing destination sheet if present to avoid error
            if SheetExists(.wkbkDest, .shtDest) Then DeleteSheet .wkbkDest, .shtDest

            With pvt.wkbkDest
                Set pvt.wkshtDest = .Worksheets.Add(After:=pvt.tblSrc.wksht)
            End With
        End If

        ' Name the destination sheet per input argument
        .wkshtDest.Name = .shtDest

        ' Create the pivot cache and table
        Set .pvtCache = _
            .wkbkDest.PivotCaches.Create(SourceType:=xlDatabase, SourceData:=.tblSrc.rngTable)
        Set .pvtTable = _
            .pvtCache.CreatePivotTable(TableDestination:=.wkshtDest.Cells(1, 1))
    End With
    Exit Function

ErrorExit:
    errs.RecordErr "CreatePivotCacheAndTable", CreatePivotCacheAndTable
End Function


'--------------------------------------------------------------------------------------
' Validate row/column field specs and normalize into class CSV attributes
' JDL 3/30/26
'
Public Function ValidateFieldSpecs(pvt, rowFields, colFields) As Boolean
    SetErrs ValidateFieldSpecs: If errs.IsHandle Then On Error GoTo ErrorExit
    Dim rowFieldsCsv As String, colFieldsCsv As String

    'Convert row and col field specs to splittable comma-separated strings
    If Not NormalizeFieldSpecs(rowFields, False, rowFieldsCsv) Then GoTo ErrorExit
    If Not NormalizeFieldSpecs(colFields, True, colFieldsCsv) Then GoTo ErrorExit

    With pvt

        'Check specified names exist and don't overlap between row and column groups
        If Not ValidateFieldNamesExist(.tblSrc, rowFieldsCsv) Then GoTo ErrorExit
        If Not ValidateFieldNamesExist(.tblSrc, colFieldsCsv) Then GoTo ErrorExit
        If Not ValidateNoFieldOverlap(rowFieldsCsv, colFieldsCsv) Then GoTo ErrorExit

        'Set class attributes for normalized field specs
        .rowFields = rowFieldsCsv
        .colFields = colFieldsCsv
    End With
    Exit Function

ErrorExit:
    errs.RecordErr "ValidateFieldSpecs", ValidateFieldSpecs
End Function

'--------------------------------------------------------------------------------------
' Validate analytes as two-item arrays: (fieldName, xFunc)
' JDL 3/30/26
'
Public Function ValidateAnalytes(pvt, analytes) As Boolean
    SetErrs ValidateAnalytes: If errs.IsHandle Then On Error GoTo ErrorExit
    Dim fieldName As String, analyteFieldsCsv As String
    Dim analyteDef As Variant

    If errs.IsFail(Not IsArray(analytes), 1) Then GoTo ErrorExit
    analyteFieldsCsv = vbNullString

    For Each analyteDef In analytes
        If errs.IsFail(Not IsArray(analyteDef), 2) Then GoTo ErrorExit

        If errs.IsFail(LBound(analyteDef) <> 0, 3) Then GoTo ErrorExit
        If errs.IsFail(UBound(analyteDef) <> 1, 4) Then GoTo ErrorExit

        fieldName = Trim(CStr(analyteDef(0)))
        If errs.IsFail(Len(fieldName) < 1, 5) Then GoTo ErrorExit
        If errs.IsFail(Not IsNumeric(analyteDef(1)), 6) Then GoTo ErrorExit

        If Len(analyteFieldsCsv) > 0 Then analyteFieldsCsv = analyteFieldsCsv & ","
        analyteFieldsCsv = analyteFieldsCsv & fieldName
    Next analyteDef

    If Not ValidateFieldNamesExist(pvt.tblSrc, analyteFieldsCsv) Then
        errs.iCodeLocal = 6
        GoTo ErrorExit
    End If
    Exit Function

ErrorExit:
    errs.RecordErr "ValidateAnalytes", ValidateAnalytes
End Function

'--------------------------------------------------------------------------------------
' Normalize row/column field spec into comma-separated string
' JDL 3/30/26
'
Private Function NormalizeFieldSpecs(fields, ByVal IsCols As Boolean, _
        fieldSpecsCsv As String) As Boolean
    SetErrs NormalizeFieldSpecs: If errs.IsHandle Then On Error GoTo ErrorExit
    Dim i As Long, fieldName As String
    Dim fieldItem As Variant, sOut As String

    fieldSpecsCsv = vbNullString
    sOut = vbNullString

    If IsArray(fields) Then
        For i = LBound(fields) To UBound(fields)
            fieldItem = fields(i)
            If errs.IsFail(IsArray(fieldItem), 1) Then GoTo ErrorExit

            fieldName = Trim(CStr(fieldItem))
            If errs.IsFail(Len(fieldName) < 1, 2) Then GoTo ErrorExit
            If Len(sOut) > 0 Then sOut = sOut & ","
            sOut = sOut & fieldName
        Next i

    ElseIf VarType(fields) = vbString Then
        fieldName = Trim(CStr(fields))
        If Not IsCols Then
            If errs.IsFail(Len(fieldName) < 1, 3) Then GoTo ErrorExit
        End If
        sOut = fieldName

    Else
        If errs.IsFail(True, 4) Then GoTo ErrorExit
    End If

    fieldSpecsCsv = sOut
    Exit Function

ErrorExit:
    errs.RecordErr "NormalizeFieldSpecs", NormalizeFieldSpecs
End Function

'--------------------------------------------------------------------------------------
' Ensure all specified field names exist in source table headers
' JDL 3/30/26
'
Private Function ValidateFieldNamesExist(tblSrc, ByVal fieldSpecsCsv As String) As Boolean
    SetErrs ValidateFieldNamesExist: If errs.IsHandle Then On Error GoTo ErrorExit
    Dim dictHeaders As Object, rngHeaders As Range
    Dim arrFields As Variant, i As Long, fieldName As String
    Dim headerCell As Range

    If Len(fieldSpecsCsv) < 1 Then Exit Function

    Set dictHeaders = New Dictionary
    If errs.IsFail(tblSrc.rngHeader Is Nothing, 1) Then GoTo ErrorExit

    For Each headerCell In tblSrc.rngHeader.Cells
        fieldName = Trim(CStr(headerCell.Value))
        If Len(fieldName) > 0 Then
            If Not dictHeaders.Exists(LCase(fieldName)) Then _
                dictHeaders.Add LCase(fieldName), True
        End If
    Next headerCell

    arrFields = Split(fieldSpecsCsv, ",")
    For i = LBound(arrFields) To UBound(arrFields)
        fieldName = Trim(CStr(arrFields(i)))
        If errs.IsFail(Not dictHeaders.Exists(LCase(fieldName)), 2) Then GoTo ErrorExit
    Next i
    Exit Function

ErrorExit:
    errs.RecordErr "ValidateFieldNamesExist", ValidateFieldNamesExist
End Function

'--------------------------------------------------------------------------------------
' Ensure same field is not assigned to both row and column groups
' JDL 3/30/26
'
Private Function ValidateNoFieldOverlap(ByVal rowFieldsCsv As String, _
        ByVal colFieldsCsv As String) As Boolean
    SetErrs ValidateNoFieldOverlap: If errs.IsHandle Then On Error GoTo ErrorExit
    Dim dictRow As Object
    Dim rowArr As Variant, colArr As Variant
    Dim i As Long, fieldName As String

    If Len(rowFieldsCsv) < 1 Or Len(colFieldsCsv) < 1 Then Exit Function

    Set dictRow = New Dictionary

    rowArr = Split(rowFieldsCsv, ",")
    For i = LBound(rowArr) To UBound(rowArr)
        fieldName = LCase(Trim(CStr(rowArr(i))))
        If Len(fieldName) > 0 Then
            If Not dictRow.Exists(fieldName) Then dictRow.Add fieldName, True
        End If
    Next i

    colArr = Split(colFieldsCsv, ",")
    For i = LBound(colArr) To UBound(colArr)
        fieldName = LCase(Trim(CStr(colArr(i))))
        If errs.IsFail(dictRow.Exists(fieldName), 1) Then GoTo ErrorExit
    Next i
    Exit Function

ErrorExit:
    errs.RecordErr "ValidateNoFieldOverlap", ValidateNoFieldOverlap
End Function

'--------------------------------------------------------------------------------------
' Configure row, column, page and data fields
' JDL 3/27/26
'
Public Function ConfigurePivotFields(pvt, analytes, filters) As Boolean
    SetErrs ConfigurePivotFields: If errs.IsHandle Then On Error GoTo ErrorExit

    With pvt
        ' Add row fields (at least one required)
        If Not AddFieldGroup(pvt, .rowFields, xlRowField) Then GoTo ErrorExit

        ' Add column fields if any (optional)
        If Not AddFieldGroup(pvt, .colFields, xlColumnField) Then GoTo ErrorExit
    End With

    ' Add page filters if any (optional) and analytes (at least one required)
    If Not AddFilterFields(pvt, filters) Then GoTo ErrorExit
    If Not AddAnalytes(pvt, analytes) Then GoTo ErrorExit
    Exit Function

ErrorExit:
    errs.RecordErr "ConfigurePivotFields", ConfigurePivotFields
End Function

'--------------------------------------------------------------------------------------
' Add row or column field group from array of strings or individual string
' JDL 3/27/26
'
Private Function AddFieldGroup(pvt, fieldsCsv, ByVal fieldOrientation As XlPivotFieldOrientation) As Boolean
    SetErrs AddFieldGroup: If errs.IsHandle Then On Error GoTo ErrorExit
    Dim i As Long, fieldName As Variant

    If Len(CStr(fieldsCsv)) < 1 Then Exit Function

    i = 1
    For Each fieldName In Split(CStr(fieldsCsv), ",")
        With pvt.pvtTable.PivotFields(fieldName)
            .Orientation = fieldOrientation
            .Position = i
        End With
        If Not SetFieldSubtotals(pvt, fieldName) Then GoTo ErrorExit
        i = i + 1
    Next fieldName
    Exit Function

ErrorExit:
    errs.RecordErr "AddFieldGroup", AddFieldGroup
End Function

'--------------------------------------------------------------------------------------
' Toggle pivot field subtotals by field-name override from pvt.dictSubT
' JDL 3/27/26
'
Private Function SetFieldSubtotals(pvt, ByVal fieldName As String) As Boolean
    SetErrs SetFieldSubtotals: If errs.IsHandle Then On Error GoTo ErrorExit
    Dim i As Long

    ' Default to no subtotals
    With pvt.pvtTable.PivotFields(fieldName)
        For i = 1 To 12
            .Subtotals(i) = False
        Next i

        ' Override with dictSubT value if specified
        If Not pvt.dictSubT Is Nothing Then
            If pvt.dictSubT.Exists(fieldName) Then _
                .Subtotals(1) = CBool(pvt.dictSubT.Item(fieldName))
        End If
    End With
    Exit Function

ErrorExit:
    errs.RecordErr "SetFieldSubtotals", SetFieldSubtotals
End Function

'--------------------------------------------------------------------------------------
' Add optional page filter fields; entry may be "FieldName" or Array(FieldName, Value)
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
' JDL 3/27/26; Updated 3/30/26
'
Private Function AddAnalytes(pvt, analytes) As Boolean
    SetErrs AddAnalytes: If errs.IsHandle Then On Error GoTo ErrorExit
    Dim fieldName As String, caption As String
    Dim xFunc As XlConsolidationFunction, analyteDef As Variant

    For Each analyteDef In analytes
        fieldName = CStr(analyteDef(0))
        xFunc = CLng(analyteDef(1))
        caption = fieldName & "_"

        On Error Resume Next
        With pvt.pvtTable
             .AddDataField .PivotFields(fieldName), caption, xFunc
        End With
        If errs.IsFail(err.Number > 0, 1, fieldName) Then
            On Error GoTo 0
            GoTo ErrorExit
        End If

        ' Reset error handling on each iteration to avoid unintended error bypass
        err.Clear
        On Error GoTo 0
        If errs.IsHandle Then On Error GoTo ErrorExit
    Next analyteDef
    Exit Function

ErrorExit:
    errs.RecordErr "AddAnalytes", AddAnalytes
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

    ' xlAscending=1; xlDescending=2
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

        ' DataFields can err when pivot has no value field; skip sort in that case.
        On Error Resume Next
        dataFieldName = pvt.pvtTable.DataFields(1).Name
        nErr = err.Number
        err.Clear
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
' Apply standard formatting and optional parameter overrides
' JDL 3/27/26; Modified 3/30/26
'
Public Function FormatPivotTable(pvt) As Boolean
    SetErrs FormatPivotTable: If errs.IsHandle Then On Error GoTo ErrorExit
    Dim rowAxisLayout As XlLayoutRowType, repeatLabels As Long
    Dim isRowGrand As Boolean, isColGrand As Boolean
    Dim hasDataPivotOrientation As Boolean, hasDataPivotPosition As Boolean
    Dim dataPivotOrientation As XlPivotFieldOrientation, dataPivotPosition As Long

    'Set defaults
    rowAxisLayout = xlTabularRow
    repeatLabels = xlRepeatLabels
    isRowGrand = False
    isColGrand = False
    hasDataPivotOrientation = False
    hasDataPivotPosition = False
    dataPivotPosition = 1

    'Check for non-default inputs specified by dictParams
    If Not ResolveFormatParams(pvt.dictParams, rowAxisLayout, repeatLabels, _
            isRowGrand, isColGrand, hasDataPivotOrientation, dataPivotOrientation, _
            hasDataPivotPosition, dataPivotPosition) Then GoTo ErrorExit

    'Set pivot table formatting
    With pvt.pvtTable
        If hasDataPivotOrientation Then
            If .DataFields.Count > 1 Then
                .DataPivotField.Orientation = dataPivotOrientation
                If hasDataPivotPosition Then .DataPivotField.Position = dataPivotPosition
            End If
        End If
        .rowAxisLayout rowAxisLayout
        .RowGrand = isRowGrand
        .ColumnGrand = isColGrand
        .PreserveFormatting = True
        .RepeatAllLabels repeatLabels
    End With
    Exit Function

ErrorExit:
    errs.RecordErr "FormatPivotTable", FormatPivotTable
End Function

'--------------------------------------------------------------------------------------
' Resolve optional formatting parameter overrides from dictionary input
' JDL 3/27/26
'
Private Function ResolveFormatParams(dictParams As Object, rowAxisLayout As XlLayoutRowType, _
    repeatLabels As Long, isRowGrand As Boolean, isColGrand As Boolean, _
    hasDataPivotOrientation As Boolean, dataPivotOrientation As XlPivotFieldOrientation, _
    hasDataPivotPosition As Boolean, dataPivotPosition As Long) As Boolean
    SetErrs ResolveFormatParams: If errs.IsHandle Then On Error GoTo ErrorExit

    If dictParams Is Nothing Then Exit Function

    With dictParams
        If .Exists("RowAxisLayout") Then rowAxisLayout = CLng(.Item("RowAxisLayout"))
        If .Exists("RepeatAllLabels") Then repeatLabels = CLng(.Item("RepeatAllLabels"))
        If .Exists("isRowGrand") Then isRowGrand = CBool(.Item("isRowGrand"))
        If .Exists("isColGrand") Then isColGrand = CBool(.Item("isColGrand"))
        If .Exists("DataPivotFieldOrientation") Then
            hasDataPivotOrientation = True
            dataPivotOrientation = CLng(.Item("DataPivotFieldOrientation"))
        End If
        If .Exists("DataPivotFieldPosition") Then
            hasDataPivotPosition = True
            dataPivotPosition = CLng(.Item("DataPivotFieldPosition"))
        End If
    End With
    Exit Function

ErrorExit:
    errs.RecordErr "ResolveFormatParams", ResolveFormatParams
End Function

'--------------------------------------------------------------------------------------
' Freeze pivot results by converting UsedRange to values
' JDL 3/27/26; Modified 3/30/26
'
Public Function ConvertPivotToValues(pvt) As Boolean
    SetErrs ConvertPivotToValues: If errs.IsHandle Then On Error GoTo ErrorExit

    With pvt.pvtTable
        If Not .DataBodyRange Is Nothing Then Set pvt.rngDataOut = .DataBodyRange
    End With
    
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
' JDL 3/27/26; Modified 3/30/26
'
Public Function SetOutputRanges(pvt) As Boolean
    SetErrs SetOutputRanges: If errs.IsHandle Then On Error GoTo ErrorExit

    With pvt
        Set .rngTableOut = .wkshtDest.UsedRange
        If Not .rngDataOut Is Nothing Then Set .rngDataOut = _
            .wkshtDest.Range(.rngDataOut.Address)
        If Not .tblOut.Init(.tblOut, .wkbkDest, .shtDest) Then GoTo ErrorExit
    End With
    Exit Function

ErrorExit:
    errs.RecordErr "SetOutputRanges", SetOutputRanges
End Function


