Attribute VB_Name = "tests_ErrorHandling"
Option Explicit
'Version 3/13/26
'--------------------------------------------------------------------------------------
' ErrorHandling and ErrorsMeta Testing
'--------------------------------------------------------------------------------------
Sub TestDriver_ErrorHandling()
    Dim procs As New Procedures, AllEnabled As Boolean

    With procs
        .Init procs, ThisWorkbook, "Tests_ErrorHandling", "Tests_ErrorHandling"
        SetApplEnvir False, False, xlCalculationManual

        AllEnabled = True
        .ErrorHandling.Enabled = True
        .ErrorParams.Enabled = True
    End With

    With procs.ErrorParams
        If .Enabled Or AllEnabled Then
            procs.curProcedure = .Name
            test_SetErrs_Defaults procs
        End If
    End With

    With procs.ErrorHandling
        If .Enabled Or AllEnabled Then
            procs.curProcedure = .Name
            test_ErrorsMeta_ResolveAndLoad procs
            test_RecordErr_UserFacingCurrent procs
            test_RecordErr_DeveloperNestedTrace procs
            test_RecordErr_BaseNotFound procs
        End If
    End With

    procs.EvalOverall procs
    SetApplEnvir True, True, xlCalculationAutomatic
End Sub
'--------------------------------------------------------------------------------------
' procs.ErrorParams
'--------------------------------------------------------------------------------------
' Verify SetErrs defaults for driver and Boolean call contexts
' JDL 3/13/26
'
Sub test_SetErrs_Defaults(procs)
    Set ExcelSteps.errs = Nothing
    Dim tst As New Test: tst.Init tst, "test_SetErrs_Defaults"
    Dim bRet As Boolean

    With tst
        ExcelSteps.SetErrs "driver", tst.wkbkTest
        .Assert tst, Not ExcelSteps.errs Is Nothing
        .Assert tst, ExcelSteps.errs.IsTesting = False
        .Assert tst, ExcelSteps.errs.IsShowMsgs = True

        Set ExcelSteps.errs = Nothing
        ExcelSteps.SetErrs bRet, tst.wkbkTest
        .Assert tst, bRet = True
        .Assert tst, ExcelSteps.errs.IsTesting = True
        .Assert tst, ExcelSteps.errs.IsShowMsgs = False

        .Update tst, procs
    End With
End Sub
'--------------------------------------------------------------------------------------
' procs.ErrorHandling
'--------------------------------------------------------------------------------------
' Verify metadata pipeline resolves expected code and message
' JDL 3/13/26
'
Sub test_ErrorsMeta_ResolveAndLoad(procs)
    Dim tst As New Test: tst.Init tst, "test_ErrorsMeta_ResolveAndLoad"
    Dim meta As Object

    SetupErrorsFixture tst
    Set meta = ExcelSteps.New_ErrorsMeta

    With tst
        ExcelSteps.errs.Locn = "TestProc"
        ExcelSteps.errs.iCodeLocal = 1

        .Assert tst, meta.Init(meta, ExcelSteps.errs)
        .Assert tst, meta.ResolveCodesFromLocn(meta, ExcelSteps.errs)
        .Assert tst, meta.LoadFromLookup(meta, ExcelSteps.errs)
        .Assert tst, meta.Validate(meta, ExcelSteps.errs.Locn)

        .Assert tst, ExcelSteps.errs.iCodeReport = 101
        .Assert tst, Not meta.IsBaseNotFound
        .Assert tst, Not meta.IsCodeNotFound
        .Assert tst, Not meta.IsMalformed
        .Assert tst, meta.IsUserFacing = True
        .Assert tst, InStr(1, meta.Message, "User visible", vbTextCompare) > 0

        .Update tst, procs
    End With
End Sub
'--------------------------------------------------------------------------------------
' Verify RecordErr user-facing current message behavior
' JDL 3/13/26
'
Sub test_RecordErr_UserFacingCurrent(procs)
    Dim tst As New Test: tst.Init tst, "test_RecordErr_UserFacingCurrent"
    Dim fRet As Boolean

    SetupErrorsFixture tst
    ExcelSteps.errs.iCodeLocal = 1
    ExcelSteps.errs.ErrParam = "X"
    fRet = True

    ExcelSteps.errs.RecordErr "TestProc", fRet

    With tst
        .Assert tst, fRet = False
        .Assert tst, ExcelSteps.errs.IsUserFacing = True
        .Assert tst, InStr(1, ExcelSteps.errs.ErrMsg, "User visible", vbTextCompare) > 0
        .Assert tst, InStr(1, ExcelSteps.errs.ErrMsg, "Called by", vbTextCompare) = 0
        .Update tst, procs
    End With
End Sub
'--------------------------------------------------------------------------------------
' Verify developer-facing nested trace append behavior
' JDL 3/13/26
'
Sub test_RecordErr_DeveloperNestedTrace(procs)
    Dim tst As New Test: tst.Init tst, "test_RecordErr_DeveloperNestedTrace"
    Dim fRet As Boolean

    SetupErrorsFixture tst
    ExcelSteps.errs.iCodeLocal = 2
    fRet = True

    ExcelSteps.errs.RecordErr "TestProc", fRet
    ExcelSteps.errs.RecordErr "CallerProc", fRet

    With tst
        .Assert tst, ExcelSteps.errs.IsUserFacing = False
        .Assert tst, InStr(1, ExcelSteps.errs.ErrMsg, "Developer detail", vbTextCompare) > 0
        .Assert tst, InStr(1, ExcelSteps.errs.ErrMsg, "Called by CallerProc", vbTextCompare) > 0
        .Update tst, procs
    End With
End Sub
'--------------------------------------------------------------------------------------
' Verify base-not-found fallback messaging
' JDL 3/13/26
'
Sub test_RecordErr_BaseNotFound(procs)
    Dim tst As New Test: tst.Init tst, "test_RecordErr_BaseNotFound"
    Dim fRet As Boolean

    SetupErrorsFixture tst
    ExcelSteps.errs.iCodeLocal = 1
    fRet = True

    ExcelSteps.errs.RecordErr "MissingProc", fRet

    With tst
        .Assert tst, InStr(1, ExcelSteps.errs.ErrMsg, "Base error code not found", vbTextCompare) > 0
        .Update tst, procs
    End With
End Sub
'--------------------------------------------------------------------------------------
' Initialize ExcelSteps.errs and populate Errors_ fixture rows used by ErrorHandling tests
' JDL 3/13/26
'
Sub SetupErrorsFixture(tst)
    Set ExcelSteps.errs = Nothing
    ExcelSteps.SetErrs False, tst.wkbkTest

    'Set as False anyway by SetErrs but provides easy toggle
    ExcelSteps.errs.IsShowMsgs = False

    Populate_Errs_Default
End Sub
