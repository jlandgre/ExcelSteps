Attribute VB_Name = "tests_ErrorHandling"
Option Explicit
'Version 3/11/26

'--------------------------------------------------------------------------------------
' ErrorHandling Class Testing
Sub TestDriver_ErrorHandling()
	Dim procs As New Procedures, AllEnabled As Boolean

	With procs
		.Init procs, ThisWorkbook, "Tests_ErrorHandling", "Tests_ErrorHandling"
		SetApplEnvir False, False, xlCalculationManual

		AllEnabled = True
		.ErrorHandling.Enabled = True
	End With

	With procs.ErrorHandling
		If .Enabled Or AllEnabled Then
			procs.curProcedure = .Name
			test_ErrorMeta_LoadFromLookup_Found procs
			test_ErrorMeta_LoadFromLookup_NotFound procs
			test_ErrorMeta_Validate_Malformed procs
			test_ErrorMeta_MessageBuilders procs
			test_AppendErrMsg_RootPaths procs
			test_AppendErrMsg_NestedTrace procs
		End If
	End With

	procs.EvalOverall procs
	SetApplEnvir True, True, xlCalculationAutomatic
End Sub
'--------------------------------------------------------------------------------------
'--------------------------------------------------------------------------------------
' procs.ErrorHandling
'--------------------------------------------------------------------------------------
'--------------------------------------------------------------------------------------
' Verify lookup maps typed fields for a valid row
' JDL 3/11/26
'
Sub test_ErrorMeta_LoadFromLookup_Found(procs)
	Dim tst As New Test: tst.Init tst, "test_ErrorMeta_LoadFromLookup_Found"
	Dim meta As Object

	SetupErrorsFixture tst
	Set meta = ExcelSteps.New_ErrorMeta

	With tst
		ExcelSteps.errs.Locn = "TestProc"
		ExcelSteps.errs.iCodeBase = 2000
		ExcelSteps.errs.iCodeLocal = 1

		.Assert tst, meta.LoadFromLookup(ExcelSteps.errs)
		.Assert tst, meta.IsFound
		.Assert tst, meta.Code = 2001
		.Assert tst, meta.Routine = "TestProc"
		.Assert tst, meta.Message = "User visible: "
		.Assert tst, meta.IsUserFacing = True
		.Update tst, procs
	End With
End Sub
'--------------------------------------------------------------------------------------
' Verify explicit not-found state when lookup row is missing
' JDL 3/11/26
'
Sub test_ErrorMeta_LoadFromLookup_NotFound(procs)
	Dim tst As New Test: tst.Init tst, "test_ErrorMeta_LoadFromLookup_NotFound"
	Dim meta As Object

	SetupErrorsFixture tst
	Set meta = ExcelSteps.New_ErrorMeta

	With tst
		ExcelSteps.errs.Locn = "TestProc"
		ExcelSteps.errs.iCodeBase = 2000
		ExcelSteps.errs.iCodeLocal = 99

		.Assert tst, meta.LoadFromLookup(ExcelSteps.errs)
		.Assert tst, Not meta.IsFound
		.Assert tst, meta.Message = "Msg Not Found"
		.Update tst, procs
	End With
End Sub
'--------------------------------------------------------------------------------------
' Verify malformed metadata row is normalized to required message
' JDL 3/11/26
'
Sub test_ErrorMeta_Validate_Malformed(procs)
	Dim tst As New Test: tst.Init tst, "test_ErrorMeta_Validate_Malformed"
	Dim meta As Object

	SetupErrorsFixture tst
	Set meta = ExcelSteps.New_ErrorMeta

	With tst
		ExcelSteps.errs.Locn = "BadProc"
		ExcelSteps.errs.iCodeBase = 3000
		ExcelSteps.errs.iCodeLocal = 1

		.Assert tst, meta.LoadFromLookup(ExcelSteps.errs)
		.Assert tst, meta.Validate("BadProc")
		.Assert tst, meta.Message = "Malformed Errors_ Row for BadProc"
		.Assert tst, meta.IsUserFacing = False
		.Update tst, procs
	End With
End Sub
'--------------------------------------------------------------------------------------
' Verify user and developer formatter methods
' JDL 3/11/26
'
Sub test_ErrorMeta_MessageBuilders(procs)
	Dim tst As New Test: tst.Init tst, "test_ErrorMeta_MessageBuilders"
	Dim meta As Object, sUser As String, sDev As String

	SetupErrorsFixture tst
	Set meta = ExcelSteps.New_ErrorMeta

	With tst
		ExcelSteps.errs.Locn = "UserProc"
		ExcelSteps.errs.iCodeBase = 4000
		ExcelSteps.errs.iCodeLocal = 1

		.Assert tst, meta.LoadFromLookup(ExcelSteps.errs)
		.Assert tst, meta.Validate("UserProc")

		sUser = meta.ToUserMessage("X")
		.Assert tst, sUser = "User visible: X"

		ExcelSteps.errs.Locn = "TestProc"
		ExcelSteps.errs.iCodeBase = 2000
		ExcelSteps.errs.iCodeLocal = 2
		.Assert tst, meta.LoadFromLookup(ExcelSteps.errs)
		sDev = meta.ToDeveloperMessage(2002, "Y")
		.Assert tst, InStr(1, sDev, "Error 2002; in sub or function, ") > 0
		.Assert tst, InStr(1, sDev, "TestProc") > 0
		.Assert tst, InStr(1, sDev, "Developer detail: Y") > 0
		.Update tst, procs
	End With
End Sub
'--------------------------------------------------------------------------------------
' Verify root error message behavior for developer and user-facing branches
' JDL 3/11/26
'
Sub test_AppendErrMsg_RootPaths(procs)
	Dim tst As New Test: tst.Init tst, "test_AppendErrMsg_RootPaths"

	SetupErrorsFixture tst

	With tst
		ExcelSteps.errs.Locn = "TestProc"
		ExcelSteps.errs.iCodeBase = 2000
		ExcelSteps.errs.iCodeLocal = 2
		ExcelSteps.errs.iCodeReport = 2002
		ExcelSteps.errs.ErrParam = "ABC"
		ExcelSteps.errs.ErrMsg = ""
		ExcelSteps.errs.IsNewErr = True
		ExcelSteps.errs.AppendErrMsg ""
		.Assert tst, Not ExcelSteps.errs.IsUserFacing
		.Assert tst, InStr(1, ExcelSteps.errs.ErrMsg, "Error 2002; in sub or function,") > 0
		.Assert tst, InStr(1, ExcelSteps.errs.ErrMsg, "TestProc") > 0
		.Assert tst, InStr(1, ExcelSteps.errs.ErrMsg, "Developer detail: ABC") > 0

		ExcelSteps.errs.Locn = "UserProc"
		ExcelSteps.errs.iCodeBase = 4000
		ExcelSteps.errs.iCodeLocal = 1
		ExcelSteps.errs.iCodeReport = 4001
		ExcelSteps.errs.ErrParam = "XYZ"
		ExcelSteps.errs.ErrMsg = ""
		ExcelSteps.errs.IsNewErr = True
		ExcelSteps.errs.AppendErrMsg ""
		.Assert tst, ExcelSteps.errs.IsUserFacing
		.Assert tst, ExcelSteps.errs.ErrMsg = "User visible: XYZ"

		.Update tst, procs
	End With
End Sub
'--------------------------------------------------------------------------------------
' Verify nested stack trace suffix only appends for non user-facing messages
' JDL 3/11/26
'
Sub test_AppendErrMsg_NestedTrace(procs)
	Dim tst As New Test: tst.Init tst, "test_AppendErrMsg_NestedTrace"

	SetupErrorsFixture tst

	With tst
		ExcelSteps.errs.ErrMsg = "Error 2002; in sub or function," & vbCrLf
		ExcelSteps.errs.IsNewErr = False
		ExcelSteps.errs.IsUserFacing = False
		ExcelSteps.errs.AppendErrMsg "Called by CallerProc"
		.Assert tst, InStr(1, ExcelSteps.errs.ErrMsg, "Called by CallerProc") > 0

		ExcelSteps.errs.ErrMsg = "User visible"
		ExcelSteps.errs.IsNewErr = False
		ExcelSteps.errs.IsUserFacing = True
		ExcelSteps.errs.AppendErrMsg "Called by CallerProc"
		.Assert tst, InStr(1, ExcelSteps.errs.ErrMsg, "Called by CallerProc") = 0

		.Update tst, procs
	End With
End Sub
'--------------------------------------------------------------------------------------
' Initialize errs and populate Errors_ fixture rows used by ErrorHandling tests
' JDL 3/11/26
'
Sub SetupErrorsFixture(tst)
	Dim IsDummy As Boolean

	IsDummy = False
	ExcelSteps.SetErrs IsDummy, tst.wkbkTest
	ExcelSteps.errs.IsShowMsgs = False

	Populate_Errs_Default
End Sub
