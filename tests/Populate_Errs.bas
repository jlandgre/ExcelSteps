Attribute VB_Name = "Populate_Errs"
Option Explicit
'Version 3/11/26

'--------------------------------------------------------------------------------------
' Populate mock Errors_ table for ErrorHandling tests
' JDL 3/11/26
'
Sub Populate_Errs_Default()
	Dim wksht As Worksheet

	Set wksht = ExcelSteps.errs.wkbkE.Sheets(ExcelSteps.shtErrors)

	With wksht
		.Cells.Clear

		'Set canonical mock Errors_ headers from project constant
		.Range("A1:G1").Value = Split(ExcelSteps.sErrorsHeadings, ",")

		'Base rows are always developer-facing (iMsgDevUser=False).
		'Base row is used when code resolves to base (e.g., unknown VBA error fallback).
		'For normal flow, iCodeReport is Base + iCodeLocal where iCodeLocal comes from IsFail.

		'Base code row for TestProc (developer-facing)
		.Cells(2, 1).Value = 2000
		.Cells(2, 3).Value = "TestProc"
		.Cells(2, 4).Value = ExcelSteps.sErrBase
		.Cells(2, 6).Value = False

		'User-facing row for TestProc (Base 2000 + iCodeLocal 1 = 2001)
		.Cells(3, 1).Value = 2001
		.Cells(3, 3).Value = "TestProc"
		.Cells(3, 4).Value = "User visible: "
		.Cells(3, 6).Value = True

		'Developer-facing non-base row for TestProc (Base 2000 + iCodeLocal 2 = 2002)
		.Cells(4, 1).Value = 2002
		.Cells(4, 3).Value = "TestProc"
		.Cells(4, 4).Value = "Developer detail: "
		.Cells(4, 6).Value = False

		'Base code row for malformed case (developer-facing)
		.Cells(5, 1).Value = 3000
		.Cells(5, 3).Value = "BadProc"
		.Cells(5, 4).Value = ExcelSteps.sErrBase
		.Cells(5, 6).Value = False

		'Malformed row: blank routine and message plus non-Boolean user flag
		.Cells(6, 1).Value = 3001
		.Cells(6, 3).Value = ""
		.Cells(6, 4).Value = ""
		.Cells(6, 6).Value = "maybe"

		'Base code row for UserProc (developer-facing)
		.Cells(7, 1).Value = 4000
		.Cells(7, 3).Value = "UserProc"
		.Cells(7, 4).Value = ExcelSteps.sErrBase
		.Cells(7, 6).Value = False

		'User-facing row
		.Cells(8, 1).Value = 4001
		.Cells(8, 3).Value = "UserProc"
		.Cells(8, 4).Value = "User visible: "
		.Cells(8, 6).Value = True
	End With
End Sub
