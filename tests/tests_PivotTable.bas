Attribute VB_Name = "tests_PivotTable"
Option Explicit
'Version 3/27/26
'--------------------------------------------------------------------------------------
' PivotTable Class Testing
Sub TestDriver_PivotTable()
	Dim procs As New Procedures, AllEnabled As Boolean

	With procs
		.Init procs, ThisWorkbook, "Tests_PivotTable", "Tests_PivotTable"
		SetApplEnvir False, False, xlCalculationManual

		'Enable testing of all or individual procedures
		AllEnabled = False
		.PivotTable.Enabled = True
	End With

	'Setup procedure group
	With procs.PivotTable
		If .Enabled Or AllEnabled Then
			procs.curProcedure = .Name
			test_InitPivotTable procs
			test_CreatePivotCacheAndTable procs
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
		PopulatePivotTableSimple ThisWorkbook, "PivotSrc"
		.Assert tst, tblSrc.Provision(tblSrc, ThisWorkbook, False, sht:="PivotSrc", _
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
' JDL 3/27/26
'
Sub test_CreatePivotCacheAndTable(procs)
	Dim tst As New Test: tst.Init tst, "test_CreatePivotCacheAndTable"
	Set ExcelSteps.errs = Nothing
	Dim tblSrc As Object, pvt As Object

	Set tblSrc = ExcelSteps.New_tbl
	Set pvt = ExcelSteps.New_PivotTable

	With tst
		PopulatePivotTableSimple ThisWorkbook, "PivotSrc"
		.Assert tst, tblSrc.Provision(tblSrc, ThisWorkbook, False, sht:="PivotSrc", _
			IsSetColNames:=True)
		.Assert tst, pvt.InitPivotTable(pvt, tblSrc, "PivotOut")
		.Assert tst, pvt.CreatePivotCacheAndTable(pvt)
		.Assert tst, Not pvt.wkbkDest Is Nothing
		.Assert tst, Not pvt.wkshtDest Is Nothing
		.Assert tst, Not pvt.pvtCache Is Nothing
		.Assert tst, Not pvt.pvtTable Is Nothing
		.Assert tst, pvt.wkshtDest.Name = "PivotOut"
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
	Dim tblSrc As Object, pvt As Object
	Dim rowFields As Variant, colFields As Variant, analytes As Variant

	Set tblSrc = ExcelSteps.New_tbl
	Set pvt = ExcelSteps.New_PivotTable

	With tst
		PopulatePivotTableSimple ThisWorkbook, "PivotSrc"
		.Assert tst, tblSrc.Provision(tblSrc, ThisWorkbook, False, sht:="PivotSrc", _
			IsSetColNames:=True)
		.Assert tst, pvt.InitPivotTable(pvt, tblSrc, "PivotOut")
		.Assert tst, pvt.CreatePivotCacheAndTable(pvt)

		rowFields = Array(Array("Category", True))
		colFields = Array(Array("SubCategory", False))

        'We specify (field name, aggregation, output name) for each analyte (inner array)
		analytes = Array(Array("Amount", xlSum, "Sum of Amount"))

		.Assert tst, pvt.ConfigurePivotFields(pvt, rowFields, colFields, analytes, vbNullString)
		.Assert tst, pvt.pvtTable.RowFields.Count = 1
		.Assert tst, pvt.pvtTable.ColumnFields.Count = 1
		.Assert tst, pvt.pvtTable.DataFields.Count = 1
		.Assert tst, pvt.pvtTable.DataFields(1).Name = "Sum of Amount"
		.Assert tst, pvt.pvtTable.PivotFields("Category").Subtotals(1) = True
		.Assert tst, pvt.pvtTable.PivotFields("SubCategory").Subtotals(1) = False
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
	Dim tblSrc As Object, pvt As Object
	Dim rowFields As Variant, colFields As Variant, analytes As Variant

	Set tblSrc = ExcelSteps.New_tbl
	Set pvt = ExcelSteps.New_PivotTable

	With tst
		PopulatePivotTableSimple ThisWorkbook, "PivotSrc"
		.Assert tst, tblSrc.Provision(tblSrc, ThisWorkbook, False, sht:="PivotSrc", _
			IsSetColNames:=True)
		.Assert tst, pvt.InitPivotTable(pvt, tblSrc, "PivotOut")
		.Assert tst, pvt.CreatePivotCacheAndTable(pvt)

		rowFields = Array("Category")
		colFields = Array("SubCategory")
		analytes = Array(Array("Amount", xlSum, "Sum of Amount"))

		.Assert tst, pvt.ConfigurePivotFields(pvt, rowFields, colFields, analytes, vbNullString)
		.Assert tst, pvt.ApplySortOrder(pvt, "asc")
		DeletePivotSht pvt
		.Update tst, procs
	End With
End Sub
'--------------------------------------------------------------------------------------
' Apply pivot formatting and grand total toggles
' JDL 3/27/26
'
Sub test_FormatPivotTable(procs)
	Dim tst As New Test: tst.Init tst, "test_FormatPivotTable"
	Set ExcelSteps.errs = Nothing
	Dim tblSrc As Object, pvt As Object
	Dim rowFields As Variant, colFields As Variant, analytes As Variant

	Set tblSrc = ExcelSteps.New_tbl
	Set pvt = ExcelSteps.New_PivotTable

	With tst
		PopulatePivotTableSimple ThisWorkbook, "PivotSrc"
		.Assert tst, tblSrc.Provision(tblSrc, ThisWorkbook, False, sht:="PivotSrc", _
			IsSetColNames:=True)
		.Assert tst, pvt.InitPivotTable(pvt, tblSrc, "PivotOut")
		.Assert tst, pvt.CreatePivotCacheAndTable(pvt)

		rowFields = Array("Category")
		colFields = Array("SubCategory")
		analytes = Array(Array("Amount", xlSum, "Sum of Amount"))

		.Assert tst, pvt.ConfigurePivotFields(pvt, rowFields, colFields, analytes, vbNullString)
		.Assert tst, pvt.FormatPivotTable(pvt, False, False)
		.Assert tst, pvt.pvtTable.RowGrand = False
		.Assert tst, pvt.pvtTable.ColumnGrand = False
		DeletePivotSht pvt
		.Update tst, procs
	End With
End Sub
'--------------------------------------------------------------------------------------
' Convert pivot layout to fixed values and persist data range pointer
' JDL 3/27/26
'
Sub test_ConvertPivotToValues(procs)
	Dim tst As New Test: tst.Init tst, "test_ConvertPivotToValues"
	Set ExcelSteps.errs = Nothing
	Dim tblSrc As Object, pvt As Object
	Dim rowFields As Variant, colFields As Variant, analytes As Variant

	Set tblSrc = ExcelSteps.New_tbl
	Set pvt = ExcelSteps.New_PivotTable

	With tst
		PopulatePivotTableSimple ThisWorkbook, "PivotSrc"
		.Assert tst, tblSrc.Provision(tblSrc, ThisWorkbook, False, sht:="PivotSrc", _
			IsSetColNames:=True)
		.Assert tst, pvt.InitPivotTable(pvt, tblSrc, "PivotOut")
		.Assert tst, pvt.CreatePivotCacheAndTable(pvt)

		rowFields = Array("Category")
		colFields = Array("SubCategory")
		analytes = Array(Array("Amount", xlSum, "Sum of Amount"))

		.Assert tst, pvt.ConfigurePivotFields(pvt, rowFields, colFields, analytes, vbNullString)
		.Assert tst, pvt.ConvertPivotToValues(pvt)
		.Assert tst, Not pvt.rngDataOut Is Nothing
		DeletePivotSht pvt
		.Update tst, procs
	End With
End Sub
'--------------------------------------------------------------------------------------
' Set tblRowsCols output pointers from pivot output sheet
' JDL 3/27/26
'
Sub test_SetOutputRanges(procs)
	Dim tst As New Test: tst.Init tst, "test_SetOutputRanges"
	Set ExcelSteps.errs = Nothing
	Dim tblSrc As Object, pvt As Object
	Dim rowFields As Variant, colFields As Variant, analytes As Variant

	Set tblSrc = ExcelSteps.New_tbl
	Set pvt = ExcelSteps.New_PivotTable

	With tst
		PopulatePivotTableSimple ThisWorkbook, "PivotSrc"
		.Assert tst, tblSrc.Provision(tblSrc, ThisWorkbook, False, sht:="PivotSrc", _
			IsSetColNames:=True)
		.Assert tst, pvt.InitPivotTable(pvt, tblSrc, "PivotOut")
		.Assert tst, pvt.CreatePivotCacheAndTable(pvt)

		rowFields = Array("Category")
		colFields = Array("SubCategory")
		analytes = Array(Array("Amount", xlSum, "Sum of Amount"))

		.Assert tst, pvt.ConfigurePivotFields(pvt, rowFields, colFields, analytes, vbNullString)
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
' JDL 3/27/26
'
Sub test_MakePivotTableProcedure1(procs)
	Dim tst As New Test: tst.Init tst, "test_MakePivotTableProcedure1"
	Set ExcelSteps.errs = Nothing
	Dim tblSrc As Object, pvt As Object
	Dim rowFields As Variant, colFields As Variant, analytes As Variant
	Dim rowA As Range, rowB As Range, colX As Range, colY As Range

	Set tblSrc = ExcelSteps.New_tbl
	Set pvt = ExcelSteps.New_PivotTable

	With tst
		PopulatePivotTableSimple ThisWorkbook, "PivotSrc"
		.Assert tst, tblSrc.Provision(tblSrc, ThisWorkbook, False, sht:="PivotSrc", _
			IsSetColNames:=True)

		rowFields = Array("Category")
		colFields = Array("SubCategory")
		analytes = Array(Array("Amount", xlSum, "Sum of Amount"))

		.Assert tst, pvt.MakePivotTableProcedure(pvt, tblSrc, rowFields, colFields, "PivotOut", _
			analytes, isRowGrand:=False, isColGrand:=False)
		.Assert tst, Not pvt.rngTableOut Is Nothing

		Set colX = ExcelSteps.FindInRange(pvt.wkshtDest.Rows(2), "X")
		Set colY = ExcelSteps.FindInRange(pvt.wkshtDest.Rows(2), "Y")
		Set rowA = ExcelSteps.FindInRange(pvt.wkshtDest.Columns(1), "A")
		Set rowB = ExcelSteps.FindInRange(pvt.wkshtDest.Columns(1), "B")

		.Assert tst, Not colX Is Nothing
		.Assert tst, Not colY Is Nothing
		.Assert tst, Not rowA Is Nothing
		.Assert tst, Not rowB Is Nothing

		.Assert tst, CDbl(Intersect(rowA.EntireRow, colX.EntireColumn).Value2) = 15
		.Assert tst, CDbl(Intersect(rowA.EntireRow, colY.EntireColumn).Value2) = 20
		.Assert tst, CDbl(Intersect(rowB.EntireRow, colX.EntireColumn).Value2) = 7
		.Assert tst, CDbl(Intersect(rowB.EntireRow, colY.EntireColumn).Value2) = 5

		DeletePivotSht pvt
		.Update tst, procs
	End With
End Sub

'--------------------------------------------------------------------------------------
' Build a row-only pivot (no col fields) and verify subtotal/grand total shape
' JDL 3/27/26
'
Sub test_MakePivotTableProcedure2(procs)
	Dim tst As New Test: tst.Init tst, "test_MakePivotTableProcedure2"
	Set ExcelSteps.errs = Nothing
	Dim tblSrc As Object, pvt As Object
	Dim rowFields As Variant, colFields As Variant, analytes As Variant
	Dim rowA As Range, rowB As Range, rowGrand As Range, colAmount As Range

	Set tblSrc = ExcelSteps.New_tbl
	Set pvt = ExcelSteps.New_PivotTable

	With tst
		PopulatePivotTableSimple ThisWorkbook, "PivotSrc"
		.Assert tst, tblSrc.Provision(tblSrc, ThisWorkbook, False, sht:="PivotSrc", _
			IsSetColNames:=True)

		rowFields = Array("Category", "SubCategory")
		colFields = vbNullString
		analytes = Array(Array("Amount", xlSum, "Sum of Amount"))

		.Assert tst, pvt.MakePivotTableProcedure(pvt, tblSrc, rowFields, colFields, "PivotOut", _
			analytes)
		.Assert tst, Not pvt.rngTableOut Is Nothing
		'.Assert tst, pvt.pvtTable.ColumnFields.Count = 0

		Set colAmount = ExcelSteps.FindInRange(pvt.wkshtDest.Rows(1), "Sum of Amount")
		Set rowA = ExcelSteps.FindInRange(pvt.wkshtDest.Columns(1), "A")
		Set rowB = ExcelSteps.FindInRange(pvt.wkshtDest.Columns(1), "B")
		Set rowGrand = ExcelSteps.FindInRange(pvt.wkshtDest.Columns(1), "Grand Total")

		.Assert tst, Not colAmount Is Nothing
		.Assert tst, Not rowA Is Nothing
		.Assert tst, Not rowB Is Nothing
		.Assert tst, Not rowGrand Is Nothing

		.Assert tst, CDbl(Intersect(rowA.EntireRow, colAmount.EntireColumn).Value2) = 35
		.Assert tst, CDbl(Intersect(rowB.EntireRow, colAmount.EntireColumn).Value2) = 12
		.Assert tst, CDbl(Intersect(rowGrand.EntireRow, colAmount.EntireColumn).Value2) = 47

		DeletePivotSht pvt
		.Update tst, procs
	End With
End Sub

'--------------------------------------------------------------------------------------
' Build a nested row pivot and verify field-level subtotals toggles
' JDL 3/27/26
'
Sub test_MakePivotTableProcedure3(procs)
	Dim tst As New Test: tst.Init tst, "test_MakePivotTableProcedure3"
	Set ExcelSteps.errs = Nothing
	Dim tblSrc As Object, pvt As Object
	Dim rowFields As Variant, colFields As Variant, analytes As Variant

	Set tblSrc = ExcelSteps.New_tbl
	Set pvt = ExcelSteps.New_PivotTable

	With tst
		PopulatePivotTableSimple ThisWorkbook, "PivotSrc"
		.Assert tst, tblSrc.Provision(tblSrc, ThisWorkbook, False, sht:="PivotSrc", _
			IsSetColNames:=True)

		rowFields = Array(Array("Category", False), Array("SubCategory", True))
		colFields = vbNullString
		analytes = Array(Array("Amount", xlSum, "Sum of Amount"))

		.Assert tst, pvt.MakePivotTableProcedure(pvt, tblSrc, rowFields, colFields, "PivotOut", _
			analytes)
		.Assert tst, Not pvt.rngTableOut Is Nothing
		'.Assert tst, pvt.pvtTable.PivotFields("Category").Subtotals(1) = False
		'.Assert tst, pvt.pvtTable.PivotFields("SubCategory").Subtotals(1) = True

		DeletePivotSht pvt
		.Update tst, procs
	End With
End Sub

Private Sub DeletePivotSht(pvt)
	If pvt Is Nothing Then Exit Sub
	If pvt.wkbkDest Is Nothing Then Exit Sub
	If Len(pvt.shtDest) < 1 Then Exit Sub
	If Not SheetExists(pvt.wkbkDest, pvt.shtDest) Then Exit Sub

	Application.DisplayAlerts = False
	pvt.wkbkDest.Sheets(pvt.shtDest).Delete
	Application.DisplayAlerts = True
End Sub
