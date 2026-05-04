Attribute VB_Name = "ErrorHandleUtil"
'Version 3/13/26
'Import this module and ErrorHandling Class Module to install error handling in a project

Option Explicit
Public errs As Object

'Error Handling Constants
Public Const shtErrors As String = "Errors_"
Public Const sErrBase As String = "Base"
Public Const sVBAErr As String = "Unknown VBA Error"
Public Const sFileErrs As String = "Warnings_and_Errors.txt"
Public Const sSettingErrs As String = "Warnings_Errors"
'-----------------------------------------------------------------------------------------------------
'Initialize local error handling and errs Class
'
''JDL 3/8/23; Modified 10/16/25
'
Sub SetErrs(CallingFunction, Optional wkbkE As Workbook = Nothing)
    Dim IsDriver As Boolean, IsNonBool As Boolean

    IsDriver = (VarType(CallingFunction) = vbString And LCase(CallingFunction) = "driver")
    IsNonBool = (VarType(CallingFunction) = vbString And LCase(CallingFunction) = "non-bool")

    'Initialize errs (In case CallingFunction called directly instead of by local driver sub)
    If (errs Is Nothing) Or IsDriver Then
        Set errs = New ErrorHandling
        
        'Default Errors_ sheet location
        If wkbkE Is Nothing Then Set wkbkE = ThisWorkbook
        
        'True/False = Master switch for enabling error handling in project
        errs.Init wkbkE, IsHandle:=False

        'Set defaults by calling mode
        If IsDriver Then
            errs.IsTesting = False
            errs.IsShowMsgs = True
        Else
            errs.IsTesting = True
            errs.IsShowMsgs = False
        End If

    ElseIf IsDriver Then
        'Driver call refreshes mode flags even if errs already exists
        errs.IsTesting = False
        errs.IsShowMsgs = True
    End If
    
    'Initialize Boolean calling function
    If Not IsDriver And Not IsNonBool Then CallingFunction = True
End Sub



