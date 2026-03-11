Attribute VB_Name = "ErrorHandleUtil"
'Version 1/29/26
'Import this module and ErrorHandling Class Module to install error handling in a project

Option Explicit
Public errs As Object

'Error Handling Constants
Public Const shtErrors As String = "Errors_"
Public Const sErrBase As String = "Base"
Public Const sVBAErr As String = "Unknown VBA Error"
Public Const iErrNotFound As Integer = 10000
Public Const sFileErrs As String = "Warnings_and_Errors.txt"
Public Const sSettingErrs As String = "Warnings_Errors"
'-----------------------------------------------------------------------------------------------------
'Initialize errs Class and Boolean calling function return value for project error handling
'JDL 3/8/23; Modified 3/11/26
' * CallingFunction is either a Boolean function return variable or String mode flag.
' * CallingFunction = "driver" for driver subs and "non-bool" for non-Boolean procedures.
' * Boolean function return variables are initialized to True; "driver" and "non-bool" are not.
' * errs is initialized when missing, and always reinitialized for a driver call.
' * Defaults by call context after init:
'      - Driver call: IsTesting=False, IsShowMsgs=True
'      - Direct Boolean or "non-bool" call: IsTesting=True, IsShowMsgs=False
' * In tests/demos, you can pre-initialize errs and explicitly override IsShowMsgs as needed.
'
Sub SetErrs(CallingFunction, Optional wkbkE As Workbook = Nothing)
    Dim IsDriverCall As Boolean, IsNonBoolCall As Boolean, IsFunctionCall As Boolean

    IsDriverCall = (CallingFunction = "driver")
    IsNonBoolCall = (CallingFunction = "non-bool")
    IsFunctionCall = (Not IsDriverCall) And (Not IsNonBoolCall)

    If IsFunctionCall Then CallingFunction = True

    If (errs Is Nothing) Or IsDriverCall Then
        Set errs = New ErrorHandling
        If wkbkE Is Nothing Then Set wkbkE = ThisWorkbook
        errs.Init wkbkE, IsHandle:=True

        If IsDriverCall Then
            errs.IsTesting = False
            errs.IsShowMsgs = True
        Else
            errs.IsTesting = True
            errs.IsShowMsgs = False
        End If
    End If
End Sub