Attribute VB_Name = "Populate_PivotTable"
Option Explicit
'-----------------------------------------------------------------------------------------------
' Populate simple rows/columns data for PivotTable tests
' JDL 3/27/26
'
Sub PopulatePivotTableSimple(wkbk, sht)
	Dim wksht As Worksheet

	If Not SheetExists(wkbk, sht) Then AddSheet wkbk, sht, wkbk.Sheets(wkbk.Sheets.Count).Name
	Set wksht = wkbk.Sheets(sht)

	With wksht
		ClearTestSheetAndNames wksht

		.Range("A1:C1").Value2 = Split("Category,SubCategory,Amount", ",")

		.Range("A2:A7").Value2 = WorksheetFunction.Transpose(Split("A,A,A,B,B,B", ","))
		.Range("B2:B7").Value2 = WorksheetFunction.Transpose(Split("X,Y,X,X,Y,Y", ","))
		.Range("C2:C7").Value2 = WorksheetFunction.Transpose(Split("10,20,5,7,3,2", ","))
	End With
End Sub
