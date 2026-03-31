Attribute VB_Name = "tests_PivotTable"
Option Explicit
Public Const shtPivotSrc As String = "SMdl"
'Version 3/31/26
'--------------------------------------------------------------------------------------
' PivotTable Class Testing
Sub TestDriver_PivotTable()
    Dim procs As New Procedures, AllEnabled As Boolean

    With procs
        .Init procs, ThisWorkbook, "Tests_PivotTable", "Tests_PivotTable"
        SetApplEnvir False, False, xlCalculationManual

        'Enable testing of all or individual procedures
        AllEnabled = False
        .PivotTable.Enabled = False
    End With
    
    'Single test
    procs.curProcedure = procs.PivotTable.Name
    test_MakePivotTableProcedure3 procs

    'Setup procedure group
    With procs.PivotTable
        If .Enabled Or AllEnabled Then
            procs.curProcedure = .Name
            test_InitPivotTable procs
            test_CreatePivotCacheAndTable procs
            test_ValidateFieldSpecs procs
            test_ValidateAnalytes procs
            test_ConfigurePivotFields procs
            test_ApplySortOrder procs
            test_FormatPivotTable procs
            test_ConvertPivotToValues procs
            test_SetOutputRanges procs
            test_MakePivotTableProcedure1 procs
            test_MakePivotTableProcedure2 procs
            test_MakePivotTableProcedure3 procs
        End If
    End With

    procs.EvalOverall procs
    SetApplEnvir True, True, xlCalculationAutomatic
End Sub
'--------------------------------------------------------------------------------------
'--------------------------------------------------------------------------------------
' procs.PivotTable
'--------------------------------------------------------------------------------------
'--------------------------------------------------------------------------------------
' Initialize PivotTable class attributes from source table and destination sheet
' JDL 3/27/26
'
Sub test_InitPivotTable(procs)
    Dim tst As New Test: tst.Init tst, "test_InitPivotTable"
    Set ExcelSteps.errs = Nothing
    Dim tblSrc As Object, pvt As Object

    Set tblSrc = ExcelSteps.New_tbl
    Set pvt = ExcelSteps.New_PivotTable

    With tst
        PopulatePivotTableSimple ThisWorkbook, shtPivotSrc
        .Assert tst, tblSrc.Provision(tblSrc, ThisWorkbook, False, sht:=shtPivotSrc, _
            IsSetColNames:=True)
        .Assert tst, pvt.InitPivotTable(pvt, tblSrc, "PivotOut")
        .Assert tst, Not pvt.tblSrc Is Nothing
        .Assert tst, Not pvt.tblOut Is Nothing
        .Assert tst, pvt.shtDest = "PivotOut"
        .Update tst, procs
    End With
End Sub
'--------------------------------------------------------------------------------------
' Create destination workbook and pivot cache/table objects
' JDL 3/27/26; Modified 3/29/26
'
Sub test_CreatePivotCacheAndTable(procs)
    Dim tst As New Test: tst.Init tst, "test_CreatePivotCacheAndTable"
    Set ExcelSteps.errs = Nothing
    Dim tblSrc As Object: Set tblSrc = ExcelSteps.New_tbl
    Dim pvt As Object: Set pvt = ExcelSteps.New_PivotTable
    Dim i As Integer, ary As Variant

    With tst
        PopulatePivotTableSimple ThisWorkbook, shtPivotSrc
        .Assert tst, tblSrc.Provision(tblSrc, ThisWorkbook, False, sht:=shtPivotSrc, _
            IsSetColNames:=True)
            
        'Check creation of PivotCache and PivotTable, wkbk and sht
        .Assert tst, pvt.InitPivotTable(pvt, tblSrc, "PivotOut")
        .Assert tst, pvt.CreatePivotCacheAndTable(pvt)
        .Assert tst, Not pvt.wkbkDest Is Nothing
        .Assert tst, Not pvt.wkshtDest Is Nothing
        .Assert tst, Not pvt.pvtCache Is Nothing
        .Assert tst, Not pvt.pvttable Is Nothing
        
        'Pivot Table is created with PivotField for each column
        .Assert tst, pvt.pvttable.PivotFields.Count = 3
        
        'Check assignment of PivotField names; .Orientation initialized to 0
        For i = 1 To pvt.pvttable.PivotFields.Count
            ary = Array("Category", "SubCategory", "Amount")
            .Assert tst, pvt.pvttable.PivotFields(i) = ary(i - 1)
            .Assert tst, pvt.pvttable.PivotFields(i).Orientation = 0
        Next i
        
        .Assert tst, pvt.wkshtDest.Name = "PivotOut"
        DeletePivotSht pvt
        .Update tst, procs
    End With
End Sub
'--------------------------------------------------------------------------------------
' Validate row/column field specs and error handling edge cases
' JDL 3/30/26
'
Sub test_ValidateFieldSpecs(procs)
    Dim tst As New Test: tst.Init tst, "test_ValidateFieldSpecs"
    Set ExcelSteps.errs = Nothing
    Dim tblSrc As Object: Set tblSrc = ExcelSteps.New_tbl
    Dim pvt As Object: Set pvt = ExcelSteps.New_PivotTable
    Dim rowFields As Variant, colFields As Variant

    With tst
        PopulatePivotTableSimple ThisWorkbook, shtPivotSrc
        .Assert tst, tblSrc.Provision(tblSrc, ThisWorkbook, False, sht:=shtPivotSrc, _
            IsSetColNames:=True)

        'Initialize up through cache/table creation
        .Assert tst, pvt.InitPivotTable(pvt, tblSrc, "PivotOut")
        .Assert tst, pvt.CreatePivotCacheAndTable(pvt)

        'Valid specs: array rowFields, string colFields
        Set ExcelSteps.errs = Nothing
        rowFields = Array("Category")
        colFields = "SubCategory"
        .Assert tst, pvt.ValidateFieldSpecs(pvt, rowFields, colFields)
        .Assert tst, pvt.rowFields = "Category"
        .Assert tst, pvt.colFields = "SubCategory"

        'Invalid rowFields type
        Set ExcelSteps.errs = Nothing
        rowFields = 42
        colFields = "SubCategory"
        .Assert tst, Not pvt.ValidateFieldSpecs(pvt, rowFields, colFields)
        .Assert tst, ExcelSteps.errs.Locn = "ValidateFieldSpecs"
        .Assert tst, ExcelSteps.errs.iCodeLocal = 4

        'Invalid empty rowFields string
        Set ExcelSteps.errs = Nothing
        rowFields = vbNullString
        colFields = "SubCategory"
        .Assert tst, Not pvt.ValidateFieldSpecs(pvt, rowFields, colFields)
        .Assert tst, ExcelSteps.errs.Locn = "ValidateFieldSpecs"
        .Assert tst, ExcelSteps.errs.iCodeLocal = 3

        'Invalid colFields type
        Set ExcelSteps.errs = Nothing
        rowFields = "Category"
        colFields = 77
        .Assert tst, Not pvt.ValidateFieldSpecs(pvt, rowFields, colFields)
        .Assert tst, ExcelSteps.errs.Locn = "ValidateFieldSpecs"
        .Assert tst, ExcelSteps.errs.iCodeLocal = 4

        'Unknown field name not found in source headers
        Set ExcelSteps.errs = Nothing
        rowFields = Array("Category", "MissingField")
        colFields = "SubCategory"
        .Assert tst, Not pvt.ValidateFieldSpecs(pvt, rowFields, colFields)
        .Assert tst, ExcelSteps.errs.Locn = "ValidateFieldSpecs"
        .Assert tst, ExcelSteps.errs.iCodeLocal = 2

        'Field overlap between row and column groups
        Set ExcelSteps.errs = Nothing
        rowFields = Array("Category", "SubCategory")
        colFields = "SubCategory"
        .Assert tst, Not pvt.ValidateFieldSpecs(pvt, rowFields, colFields)
        .Assert tst, ExcelSteps.errs.Locn = "ValidateFieldSpecs"
        .Assert tst, ExcelSteps.errs.iCodeLocal = 1

        DeletePivotSht pvt
        .Update tst, procs
    End With
End Sub
'--------------------------------------------------------------------------------------
' Validate analytes spec and error handling edge cases
' JDL 3/30/26
'
Sub test_ValidateAnalytes(procs)
    Dim tst As New Test: tst.Init tst, "test_ValidateAnalytes"
    Set ExcelSteps.errs = Nothing
    Dim tblSrc As Object: Set tblSrc = ExcelSteps.New_tbl
    Dim pvt As Object: Set pvt = ExcelSteps.New_PivotTable
    Dim analytes As Variant

    With tst
        PopulatePivotTableSimple ThisWorkbook, shtPivotSrc
        .Assert tst, tblSrc.Provision(tblSrc, ThisWorkbook, False, sht:=shtPivotSrc, _
            IsSetColNames:=True)
        .Assert tst, pvt.InitPivotTable(pvt, tblSrc, "PivotOut")
        .Assert tst, pvt.CreatePivotCacheAndTable(pvt)

        'Valid analytes: two-item arrays (fieldName, xFunc)
        Set ExcelSteps.errs = Nothing
        analytes = Array(Array("Amount", xlSum))
        .Assert tst, pvt.ValidateAnalytes(pvt, analytes)

        'Invalid analytes root type
        Set ExcelSteps.errs = Nothing
        analytes = "Amount"
        .Assert tst, Not pvt.ValidateAnalytes(pvt, analytes)
        .Assert tst, ExcelSteps.errs.Locn = "ValidateAnalytes"
        .Assert tst, ExcelSteps.errs.iCodeLocal = 1

        'Invalid analyte item type
        Set ExcelSteps.errs = Nothing
        analytes = Array("Amount")
        .Assert tst, Not pvt.ValidateAnalytes(pvt, analytes)
        .Assert tst, ExcelSteps.errs.Locn = "ValidateAnalytes"
        .Assert tst, ExcelSteps.errs.iCodeLocal = 2

        'Invalid analyte item count (must be exactly two)
        Set ExcelSteps.errs = Nothing
        analytes = Array(Array("Amount"))
        .Assert tst, Not pvt.ValidateAnalytes(pvt, analytes)
        .Assert tst, ExcelSteps.errs.Locn = "ValidateAnalytes"
        .Assert tst, ExcelSteps.errs.iCodeLocal = 4

        'Invalid analyte field name blank
        Set ExcelSteps.errs = Nothing
        analytes = Array(Array(vbNullString, xlSum))
        .Assert tst, Not pvt.ValidateAnalytes(pvt, analytes)
        .Assert tst, ExcelSteps.errs.Locn = "ValidateAnalytes"
        .Assert tst, ExcelSteps.errs.iCodeLocal = 5

        'Invalid analyte xFunc type
        Set ExcelSteps.errs = Nothing
        analytes = Array(Array("Amount", "badfunc"))
        .Assert tst, Not pvt.ValidateAnalytes(pvt, analytes)
        .Assert tst, ExcelSteps.errs.Locn = "ValidateAnalytes"
        .Assert tst, ExcelSteps.errs.iCodeLocal = 6

        'Invalid analyte field missing in source headers
        Set ExcelSteps.errs = Nothing
        analytes = Array(Array("MissingField", xlSum))
        .Assert tst, Not pvt.ValidateAnalytes(pvt, analytes)
        .Assert tst, ExcelSteps.errs.Locn = "ValidateAnalytes"
        .Assert tst, ExcelSteps.errs.iCodeLocal = 7

        DeletePivotSht pvt
        .Update tst, procs
    End With
End Sub
'--------------------------------------------------------------------------------------
' Configure row/column/data fields on created pivot table
' JDL 3/27/26
'
Sub test_ConfigurePivotFields(procs)
    Dim tst As New Test: tst.Init tst, "test_ConfigurePivotFields"
    Set ExcelSteps.errs = Nothing
    Dim tblSrc As Object: Set tblSrc = ExcelSteps.New_tbl
    Dim pvt As Object: Set pvt = ExcelSteps.New_PivotTable
    Dim rowFields As Variant, colFields As Variant, analytes As Variant

    With tst
        PopulatePivotTableSimple ThisWorkbook, shtPivotSrc
        .Assert tst, tblSrc.Provision(tblSrc, ThisWorkbook, False, sht:=shtPivotSrc, _
            IsSetColNames:=True)
        .Assert tst, pvt.InitPivotTable(pvt, tblSrc, "PivotOut")
        .Assert tst, pvt.CreatePivotCacheAndTable(pvt)

        rowFields = Array("Category")
        colFields = Array("SubCategory")

        'We specify (field name, aggregation) for each analyte (inner array)
        analytes = Array(Array("Amount", xlSum))

        .Assert tst, pvt.ValidateFieldSpecs(pvt, rowFields, colFields)
        .Assert tst, pvt.ConfigurePivotFields(pvt, analytes, vbNullString)
        
        'Check SubCategory assigned to columns (xlColumnField=2)
        .Assert tst, CStr(pvt.wkshtDest.Cells(2, 2).Value2) = "X"
        .Assert tst, CStr(pvt.wkshtDest.Cells(2, 3).Value2) = "Y"
        .Assert tst, pvt.pvttable.PivotFields(2).Orientation = xlColumnField
        
        'Check Category assigned to rows (xlRowField=1)
        .Assert tst, CStr(pvt.wkshtDest.Cells(3, 1).Value2) = "A"
        .Assert tst, CStr(pvt.wkshtDest.Cells(4, 1).Value2) = "B"
        .Assert tst, pvt.pvttable.PivotFields(1).Orientation = xlRowField
        DeletePivotSht pvt
        .Update tst, procs
    End With
End Sub
'--------------------------------------------------------------------------------------
' Apply sort order to configured column field
' JDL 3/27/26
'
Sub test_ApplySortOrder(procs)
    Dim tst As New Test: tst.Init tst, "test_ApplySortOrder"
    Set ExcelSteps.errs = Nothing
    Dim tblSrc As Object: Set tblSrc = ExcelSteps.New_tbl
    Dim pvt As Object: Set pvt = ExcelSteps.New_PivotTable
    Dim rowFields As Variant, colFields As Variant, analytes As Variant

    With tst
        PopulatePivotTableSimple ThisWorkbook, shtPivotSrc
        .Assert tst, tblSrc.Provision(tblSrc, ThisWorkbook, False, sht:=shtPivotSrc, _
            IsSetColNames:=True)
        .Assert tst, pvt.InitPivotTable(pvt, tblSrc, "PivotOut")
        .Assert tst, pvt.CreatePivotCacheAndTable(pvt)

        rowFields = Array("Category")
        colFields = Array("SubCategory")
        analytes = Array(Array("Amount", xlSum))

        .Assert tst, pvt.ValidateFieldSpecs(pvt, rowFields, colFields)
        .Assert tst, pvt.ConfigurePivotFields(pvt, analytes, vbNullString)
        .Assert tst, pvt.ApplySortOrder(pvt, "desc")
        
        'Check descending order
        .Assert tst, pvt.wkshtDest.Cells(2, 2).Value2 = "Y"
        DeletePivotSht pvt
        .Update tst, procs
    End With
End Sub
'--------------------------------------------------------------------------------------
' Apply pivot formatting and grand total toggles
' (No colFields so Sum by Category+SubCategory combos)
' JDL 3/27/26
'
Sub test_FormatPivotTable(procs)
    Dim tst As New Test: tst.Init tst, "test_FormatPivotTable"
    Set ExcelSteps.errs = Nothing
    
    Dim tblSrc As Object: Set tblSrc = ExcelSteps.New_tbl
    Dim pvt As Object: Set pvt = ExcelSteps.New_PivotTable
    Dim rowFields As Variant, colFields As Variant, analytes As Variant
    Dim cellGrand As Range

    With tst
        PopulatePivotTableSimple ThisWorkbook, shtPivotSrc
        .Assert tst, tblSrc.Provision(tblSrc, ThisWorkbook, False, sht:=shtPivotSrc, _
            IsSetColNames:=True)
        .Assert tst, pvt.InitPivotTable(pvt, tblSrc, "PivotOut")
        .Assert tst, pvt.CreatePivotCacheAndTable(pvt)

        rowFields = Array("Category", "SubCategory")
        colFields = vbNullString
        analytes = Array(Array("Amount", xlSum))

        .Assert tst, pvt.ValidateFieldSpecs(pvt, rowFields, colFields)
        .Assert tst, pvt.ConfigurePivotFields(pvt, analytes, vbNullString)
        .Assert tst, pvt.FormatPivotTable(pvt)
        
        'Headers
        .Assert tst, CStr(pvt.wkshtDest.Cells(1, 1).Value2) = "Category"
        .Assert tst, CStr(pvt.wkshtDest.Cells(1, 2).Value2) = "SubCategory"
        .Assert tst, CStr(pvt.wkshtDest.Cells(1, 3).Value2) = "Amount_"
        
        'First and last data rows
        .Assert tst, CStr(pvt.wkshtDest.Cells(2, 1).Value2) = "A"
        .Assert tst, CStr(pvt.wkshtDest.Cells(2, 2).Value2) = "X"
        .Assert tst, CStr(pvt.wkshtDest.Cells(2, 3).Value2) = "15"
        
        .Assert tst, CStr(pvt.wkshtDest.Cells(5, 1).Value2) = "B"
        .Assert tst, CStr(pvt.wkshtDest.Cells(5, 2).Value2) = "Y"
        .Assert tst, CStr(pvt.wkshtDest.Cells(5, 3).Value2) = "5"
        
        'Grand Total off by default
        Set cellGrand = ExcelSteps.FindInRange(pvt.wkshtDest.UsedRange, "Grand Total")
        .Assert tst, cellGrand Is Nothing
        
        DeletePivotSht pvt
        .Update tst, procs
    End With
End Sub
'--------------------------------------------------------------------------------------
' Convert pivot layout to fixed values and persist data range pointer
' SubCategory as colField
' JDL 3/27/26; Modified 3/30/26
'
Sub test_ConvertPivotToValues(procs)
    Dim tst As New Test: tst.Init tst, "test_ConvertPivotToValues"
    Set ExcelSteps.errs = Nothing
    
    Dim tblSrc As Object: Set tblSrc = ExcelSteps.New_tbl
    Dim pvt As Object: Set pvt = ExcelSteps.New_PivotTable
    Dim rowFields As Variant, colFields As Variant, analytes As Variant

    With tst
        PopulatePivotTableSimple ThisWorkbook, shtPivotSrc
        .Assert tst, tblSrc.Provision(tblSrc, ThisWorkbook, False, sht:=shtPivotSrc, _
            IsSetColNames:=True)
        .Assert tst, pvt.InitPivotTable(pvt, tblSrc, "PivotOut")
        .Assert tst, pvt.CreatePivotCacheAndTable(pvt)

        rowFields = Array("Category")
        colFields = Array("SubCategory")
        analytes = Array(Array("Amount", xlSum))

        .Assert tst, pvt.ValidateFieldSpecs(pvt, rowFields, colFields)
        .Assert tst, pvt.ConfigurePivotFields(pvt, analytes, vbNullString)
        .Assert tst, pvt.FormatPivotTable(pvt)
        .Assert tst, pvt.ConvertPivotToValues(pvt)
        .Assert tst, Not pvt.rngDataOut Is Nothing
        DeletePivotSht pvt
        .Update tst, procs
    End With
End Sub
'--------------------------------------------------------------------------------------
' Set tblRowsCols output pointers from pivot output sheet
' SubCategory as colField
' JDL 3/27/26
'
Sub test_SetOutputRanges(procs)
    Dim tst As New Test: tst.Init tst, "test_SetOutputRanges"
    Set ExcelSteps.errs = Nothing
    
    Dim tblSrc As Object: Set tblSrc = ExcelSteps.New_tbl
    Dim pvt As Object: Set pvt = ExcelSteps.New_PivotTable
    Dim rowFields As Variant, colFields As Variant, analytes As Variant

    With tst
        PopulatePivotTableSimple ThisWorkbook, shtPivotSrc
        .Assert tst, tblSrc.Provision(tblSrc, ThisWorkbook, False, sht:=shtPivotSrc, _
            IsSetColNames:=True)
        .Assert tst, pvt.InitPivotTable(pvt, tblSrc, "PivotOut")
        .Assert tst, pvt.CreatePivotCacheAndTable(pvt)

        rowFields = Array("Category")
        colFields = Array("SubCategory")
        analytes = Array(Array("Amount", xlSum))

        .Assert tst, pvt.ValidateFieldSpecs(pvt, rowFields, colFields)
        .Assert tst, pvt.ConfigurePivotFields(pvt, analytes, vbNullString)
        .Assert tst, pvt.FormatPivotTable(pvt)
        .Assert tst, pvt.ConvertPivotToValues(pvt)
        .Assert tst, pvt.SetOutputRanges(pvt)
        .Assert tst, Not pvt.rngTableOut Is Nothing
        .Assert tst, Not pvt.tblOut Is Nothing
        DeletePivotSht pvt
        .Update tst, procs
    End With
End Sub
'--------------------------------------------------------------------------------------
' Build a simple pivot with sum analyte and verify category/subcategory totals
' SubCategory as colField
' JDL 3/27/26; Modified 3/30/26
'
Sub test_MakePivotTableProcedure1(procs)
    Dim tst As New Test: tst.Init tst, "test_MakePivotTableProcedure1"
    Set ExcelSteps.errs = Nothing
    Dim tblSrc As Object: Set tblSrc = ExcelSteps.New_tbl
    Dim pvt As Object: Set pvt = ExcelSteps.New_PivotTable
    Dim rowFields As Variant, colFields As Variant, analytes As Variant

    With tst
        PopulatePivotTableSimple ThisWorkbook, shtPivotSrc
        .Assert tst, tblSrc.Provision(tblSrc, ThisWorkbook, False, sht:=shtPivotSrc, _
            IsSetColNames:=True)

        rowFields = Array("Category")
        colFields = Array("SubCategory")
        analytes = Array(Array("Amount", xlSum))

        .Assert tst, pvt.MakePivotTableProcedure(pvt, tblSrc, rowFields, colFields, "PivotOut", _
            analytes)
        .Assert tst, Not pvt.rngTableOut Is Nothing

        'Headers
        .Assert tst, CStr(pvt.wkshtDest.Cells(1, 1).Value2) = "Amount_"
        .Assert tst, CStr(pvt.wkshtDest.Cells(1, 2).Value2) = "SubCategory"
        
        .Assert tst, CStr(pvt.wkshtDest.Cells(2, 1).Value2) = "Category"
        .Assert tst, CStr(pvt.wkshtDest.Cells(2, 2).Value2) = "X"
        .Assert tst, CStr(pvt.wkshtDest.Cells(2, 3).Value2) = "Y"
        
        'First and last data rows
        .Assert tst, CStr(pvt.wkshtDest.Cells(3, 1).Value2) = "A"
        .Assert tst, CStr(pvt.wkshtDest.Cells(3, 2).Value2) = "15"
        .Assert tst, CStr(pvt.wkshtDest.Cells(3, 3).Value2) = "20"
        
        .Assert tst, CStr(pvt.wkshtDest.Cells(4, 1).Value2) = "B"
        .Assert tst, CStr(pvt.wkshtDest.Cells(4, 2).Value2) = "7"
        .Assert tst, CStr(pvt.wkshtDest.Cells(4, 3).Value2) = "5"

        DeletePivotSht pvt
        .Update tst, procs
    End With
End Sub

'--------------------------------------------------------------------------------------
' Build a row-only pivot (no col fields) and verify grand total shape
' JDL 3/27/26
'
Sub test_MakePivotTableProcedure2(procs)
    Dim tst As New Test: tst.Init tst, "test_MakePivotTableProcedure2"
    Set ExcelSteps.errs = Nothing
    Dim tblSrc As Object: Set tblSrc = ExcelSteps.New_tbl
    Dim pvt As Object: Set pvt = ExcelSteps.New_PivotTable
    Dim rowFields As Variant, colFields As Variant, analytes As Variant
    Dim dictParams As Object: Set dictParams = ExcelSteps.New_Dictionary
    Dim cellGrand As Range

    With tst
        PopulatePivotTableSimple ThisWorkbook, shtPivotSrc
        .Assert tst, tblSrc.Provision(tblSrc, ThisWorkbook, False, sht:=shtPivotSrc, _
            IsSetColNames:=True)

        rowFields = Array("Category", "SubCategory")
        colFields = vbNullString
        analytes = Array(Array("Amount", xlSum))
        
        ' Expect just totals at bottom (isColGrand) because no data columns for isRowGrand to sum
        dictParams.Add "isRowGrand", True
        dictParams.Add "isColGrand", True

        .Assert tst, pvt.MakePivotTableProcedure(pvt, tblSrc, rowFields, colFields, "PivotOut", _
            analytes, dictParams:=dictParams)
        .Assert tst, Not pvt.rngTableOut Is Nothing

        'Headers
        .Assert tst, CStr(pvt.wkshtDest.Cells(1, 1).Value2) = "Category"
        .Assert tst, CStr(pvt.wkshtDest.Cells(1, 2).Value2) = "SubCategory"
        .Assert tst, CStr(pvt.wkshtDest.Cells(1, 3).Value2) = "Amount_"
        
        'First and last data rows
        .Assert tst, CStr(pvt.wkshtDest.Cells(2, 1).Value2) = "A"
        .Assert tst, CStr(pvt.wkshtDest.Cells(2, 2).Value2) = "X"
        .Assert tst, CStr(pvt.wkshtDest.Cells(2, 3).Value2) = "15"
        
        .Assert tst, CStr(pvt.wkshtDest.Cells(5, 1).Value2) = "B"
        .Assert tst, CStr(pvt.wkshtDest.Cells(5, 2).Value2) = "Y"
        .Assert tst, CStr(pvt.wkshtDest.Cells(5, 3).Value2) = "5"
        
        Set cellGrand = ExcelSteps.FindInRange(pvt.wkshtDest.UsedRange, "Grand Total")
        .Assert tst, Not cellGrand Is Nothing

        DeletePivotSht pvt
        .Update tst, procs
    End With
End Sub

'--------------------------------------------------------------------------------------
' Build OTB-style pivot with Store/Values/Prodtype rows, Week/Year columns, and 3 analytes
' JDL 3/30/26
'
Sub test_MakePivotTableProcedure3(procs)
    Dim tst As New Test: tst.Init tst, "test_MakePivotTableProcedure3"
    Set ExcelSteps.errs = Nothing
    Dim tblSrc As Object: Set tblSrc = ExcelSteps.New_tbl
    Dim pvt As Object: Set pvt = ExcelSteps.New_PivotTable
    Dim rowFields As Variant, colFields As Variant, analytes As Variant
    Dim dictParams As Object: Set dictParams = ExcelSteps.New_Dictionary
    Dim rowDX As Range, rowMX As Range, rowCX As Range

    With tst
        PopulatePivotTableOTBLike ThisWorkbook, shtPivotSrc
        .Assert tst, tblSrc.Provision(tblSrc, ThisWorkbook, False, sht:=shtPivotSrc, _
            IsSetColNames:=True)

        rowFields = Array("Store", "Prodtype")
        colFields = Array("Week", "Year")
        analytes = Array(Array("Discounts", xlSum), _
            Array("Markdowns", xlSum), _
            Array("COGS", xlSum))

        ' Place Values field in rows between Store and Prodtype (Store, Sigma Values, Prodtype)
        dictParams.Add "DataPivotFieldOrientation", xlRowField
        dictParams.Add "DataPivotFieldPosition", 2

        .Assert tst, pvt.MakePivotTableProcedure(pvt, tblSrc, rowFields, colFields, "PivotOut", _
            analytes, dictParams:=dictParams)
        .Assert tst, Not pvt.rngTableOut Is Nothing

        Set rowDX = FindRowByThreeVals(pvt.wkshtDest, "Store1", "Discounts", "X")
        Set rowMX = FindRowByThreeVals(pvt.wkshtDest, "Store1", "Markdowns", "X")
        Set rowCX = FindRowByThreeVals(pvt.wkshtDest, "Store1", "COGS", "X")

        .Assert tst, Not rowDX Is Nothing
        .Assert tst, Not rowMX Is Nothing
        .Assert tst, Not rowCX Is Nothing

        If Not rowDX Is Nothing Then
            .Assert tst, Not ExcelSteps.FindInRange(rowDX.EntireRow, 10) Is Nothing
            .Assert tst, Not ExcelSteps.FindInRange(rowDX.EntireRow, 11) Is Nothing
        End If
        If Not rowMX Is Nothing Then
            .Assert tst, Not ExcelSteps.FindInRange(rowMX.EntireRow, 100) Is Nothing
            .Assert tst, Not ExcelSteps.FindInRange(rowMX.EntireRow, 101) Is Nothing
        End If
        If Not rowCX Is Nothing Then
            .Assert tst, Not ExcelSteps.FindInRange(rowCX.EntireRow, 1000) Is Nothing
            .Assert tst, Not ExcelSteps.FindInRange(rowCX.EntireRow, 1001) Is Nothing
        End If

        DeletePivotSht pvt
        .Update tst, procs
    End With
End Sub

Private Function FindRowByThreeVals(wksht As Worksheet, ByVal val1 As String, _
        ByVal val2 As String, ByVal val3 As String) As Range
    Dim i As Long, lastRow As Long

    lastRow = wksht.Cells(wksht.Rows.Count, 1).End(xlUp).Row
    For i = 1 To lastRow
        If CStr(wksht.Cells(i, 1).Value2) = val1 _
                And InStr(1, CStr(wksht.Cells(i, 2).Value2), val2, vbTextCompare) > 0 _
                And CStr(wksht.Cells(i, 3).Value2) = val3 Then
            Set FindRowByThreeVals = wksht.Cells(i, 1)
            Exit Function
        End If
    Next i
End Function

' Private Function FindRowByTwoVals(wksht As Worksheet, ByVal val1 As String, _
'         ByVal val2 As String) As Range
'     Dim i As Long, lastRow As Long

'     lastRow = wksht.Cells(wksht.Rows.Count, 1).End(xlUp).Row
'     For i = 2 To lastRow
'         If CStr(wksht.Cells(i, 1).Value2) = val1 And CStr(wksht.Cells(i, 2).Value2) = val2 Then
'             Set FindRowByTwoVals = wksht.Cells(i, 1)
'             Exit Function
'         End If
'     Next i
' End Function

Private Sub DeletePivotSht(pvt)
    If pvt Is Nothing Then Exit Sub
    If pvt.wkbkDest Is Nothing Then Exit Sub
    If Len(pvt.shtDest) < 1 Then Exit Sub
    If Not SheetExists(pvt.wkbkDest, pvt.shtDest) Then Exit Sub

    Application.DisplayAlerts = False
    pvt.wkbkDest.Sheets(pvt.shtDest).Delete
    Application.DisplayAlerts = True
End Sub




