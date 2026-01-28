Attribute VB_Name = "ErrorHandleUtil"
'ExcelSteps_ErrorHandlUtil.vb
'Version 9/19/25
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
'Initialize local error handling and errs Class
'
''JDL 3/8/23; Modified 10/16/25
'
Sub SetErrs(CallingFunction, Optional wkbkE As Workbook = Nothing)
    Dim Msgs_accum As String

    'Initialize errs (In case CallingFunction called directly instead of by local driver sub)
    If (errs Is Nothing) Or (CallingFunction = "driver") Then
        Set errs = New ErrorHandling
        
        ' If errs not instanced by driver, assume CallingFunction is being tested (for errs.ReportMsg)
        errs.IsTesting = True
        If CallingFunction = "driver" Then errs.IsTesting = False
        
        'Default Errors_ sheet location
        If wkbkE Is Nothing Then Set wkbkE = ThisWorkbook
        
        'True/False = Master switch for enabling error handling in project
        errs.Init wkbkE, IsHandle:=False
    End If
    
    'Initialize Boolean calling function
    If (CallingFunction <> "driver") And (CallingFunction <> "non-bool") Then CallingFunction = True
End Sub

