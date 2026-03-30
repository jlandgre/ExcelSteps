VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ErrorHandling"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Version 3/13/26
'Class for error handling in cascaded function architecture
'Version 6/7/23 - Add ShowMessage and FlagError
'Version 7/10/23 - Fix bug with ShowMessage args
'Version 8/29/23 - Refactor to eliminate errs as argument; Modify LookupCommentMsg args
'Version 11/21/23 - Mods to .RecordErr(), .AppendErrMsg() to clean up user-facing msgs
'Version 5/22/24 - Add Msgs_accum and IsShowMsgs attributes to enable passing messages to
'                  unit tests as alternative to Msgbox display during automated testing
'Version 9/19/25 Modify SetErrs and .Init method for wkbkE different from ThisWorkbook
'                add IsTesting attribute and logic to .ReportMsg
'Version 10/16/25 - update ShowMessage
'Version 10/30/25 - update .Init arg LookupCommentMsg
'Version 11/24/25 - update to .setMsgTitleAndText
'Version 12/3/25 - updates to .RecordMsg and .ReportWarningMsg

Option Explicit

Public IsHandle As Boolean 'Toggle error handling on/off
Public IsDriver As Boolean 'Trigger stack tracing messaging for nested functions
Public IsNewErr As Boolean 'Trigger handling of new error versus stack tracing
Public IsUserFacing As Boolean 'True for user-facing error messaging (versus VBA stack
                               'tracing reporting)
Public IsTesting As Boolean 'Flag to report ErrMsg if testing individual function etc.
Public Locn As String 'Current sub or function Error locn
Public iCodeLocal As Integer 'Raw error code - integer set in routine where error occurs
Public iCodeBase As Integer
Public iCodeReport As Integer 'Lookup table error code (composite sub/function base code + iCode)
Public ErrParam As String 'Optional param value to report in messaging; typically assign in code
                          'where error occurs
Public ErrMsg As String 'Error message string --constructed by ProcessCurrentErrMsg()
Public Msgs_accum As String 'Accumulated error and warning messages
Public IsShowMsgs As Boolean 'Flag to turn off showing MsgBox messages (for unit testing etc.)

Public wkbk As Workbook 'Application workbook; same as wkbkE if contains shtErrors lookup table
Public wkbkE As Workbook 'Workbook containing shtErrors lookup table (optionally can be add-in)
Public wkshtActive As Worksheet 'Initial active sheet - to reset at end of execution
'------------------------------------------------------------------------------------------------------
'Initialize errs Class
' Declare errs as global prior to first, user-initiated procedure sub (Public errs as Object)
' Call errs.Init from user-initiated sub after instancing errs (dim errs as New ErrorHandling)
'
'JDL 12/20/22   Modified 9/19/25 add wkbkE arg and logic for setting .wkbkE
'
Public Sub Init(wkbkE, Optional IsHandle)
    With errs
        If IsMissing(IsHandle) Then
            .IsHandle = True
        Else
            .IsHandle = IsHandle
        End If
        
        'Not used anywhere? JDL 9/19/25 - delete attr if no issues found
        'Set .wkbk = ThisWorkbook
        
        'If shtErrors resides in a different wkbk (e.g. in project wkbk) than ErrorHandling class
        Set .wkbkE = ThisWorkbook
        If Not IsMissing(wkbkE) Then Set .wkbkE = wkbkE
        
        .iCodeLocal = 0
        .iCodeBase = 0
        .iCodeReport = 0
        .ErrParam = ""
        .ErrMsg = ""
        .IsUserFacing = False
        .IsNewErr = True
        .IsDriver = False
        .IsShowMsgs = True
    End With
End Sub
'-----------------------------------------------------------------------------------------------------
' Record new error or append nested trace message
'
' Inputs:   Locn [String] sub or function where error occurred
'         IsDriver [Boolean] True if Locn is a driver sub
'
' Created:  2/17/21 JDL      Modified: 9/19/25 add .IsTesting attribute
'                                     12/2/25 set .Locn with each call/outside of If .IsNewErr
'
Sub RecordErr(Locn, Optional ByRef CallingFunction)
    With errs
        If Not .SetCallContext(CallingFunction) Then Exit Sub
        .Locn = Locn

        'Nested branch: append trace only for developer-facing flow
        If Not .IsNewErr Then
            If Not .IsUserFacing Then .ErrMsg = .ErrMsg & "Called by " & Locn & vbCrLf
            .ReportIfEnabled
            Exit Sub
        End If

        'Current branch: build message segment from metadata pipeline
        If Not .ProcessCurrentErrMsg() Then Exit Sub
        .ReportIfEnabled
    End With
End Sub
'-----------------------------------------------------------------------------------------------------
' Set call-context flags and optional calling-function return value
'
Public Function SetCallContext(Optional ByRef CallingFunction) As Boolean
    With errs
        If IsMissing(CallingFunction) Then
            .IsDriver = True
        Else
            CallingFunction = False
        End If
    End With
    SetCallContext = True
End Function
'-----------------------------------------------------------------------------------------------------
' Build and append current error message segment
'
Public Function ProcessCurrentErrMsg() As Boolean
    Dim meta As Object, sMsgNew As String

    Set meta = New ErrorsMeta

    With errs
        If Not meta.Init(meta, errs) Then Exit Function
        If Not meta.ResolveCodesFromLocn(meta, errs) Then Exit Function
        If Not meta.LoadFromLookup(meta, errs) Then Exit Function
        If Not meta.Validate(meta, .Locn) Then Exit Function

        If Len(.ErrMsg) > 0 Then .ErrMsg = .ErrMsg & vbCrLf

        If meta.IsBaseNotFound Then
            sMsgNew = "Base error code not found for routine: " & .Locn
            .IsUserFacing = False
        ElseIf meta.IsCodeNotFound Then
            sMsgNew = .Locn & " Error code not found: " & .iCodeReport
            .IsUserFacing = False
        ElseIf meta.IsMalformed Then
            sMsgNew = meta.Message
            .IsUserFacing = False
        Else
            .IsUserFacing = meta.IsUserFacing
            If .IsUserFacing Then
                sMsgNew = meta.ToUserMessage(meta, errs)
            Else
                sMsgNew = meta.ToDeveloperMessage(meta, errs)
            End If
        End If

        .ErrMsg = .ErrMsg & sMsgNew
        .IsNewErr = False
    End With

    ProcessCurrentErrMsg = True
End Function
'-----------------------------------------------------------------------------------------------------
' Apply reporting gate and trigger ReportMsg
'
Public Function ReportIfEnabled() As Boolean
    With errs
        If (.IsDriver Or .IsTesting) Then .ReportMsg
    End With
    ReportIfEnabled = True
End Function
'------------------------------------------------------------------------------------------------------
' Report Errors to user (Call from ErrorExit at end of first, user-initiated procedural routine)
'
' Modified JDL 5/22/24 add .Msgs_accum and .IsShowMsgs
'
Public Function ReportMsg() As Long
    Dim sMsg As String, sTitle As String, i As Integer
    With errs
        If Not errs.IsHandle Then Exit Function
        
        'Add the message to accumulated error/warning string
        .UpdateMsgsAccum errs.ErrMsg
        
        'Show either default or error-specific message
        If .IsShowMsgs Then
            sTitle = "Execution Error"
            i = MsgBox(errs.ErrMsg, vbOKOnly + vbCritical, sTitle)
        End If
        
        'Reset (global) IsDriver for next usage
        'Set errs = Nothing
    End With
End Function

'-----------------------------------------------------------------------------------------------------
' Configure and report warnings
'
'
' Created:  6/8/21 JDL      Modified: 6/8/23 Add & errs.ErrParam to sMsg
'                                    5/22/24 add .Msgs_accum and .IsShowMsgs
'                                    12/3/25 Fix bug; add .wkbkE arg to .Init call at end
'                                    3/13/26 add optional vbType arg
'
Sub ReportWarningMsg(iCode, Locn, Optional param, Optional ByVal vbType As VbMsgBoxStyle = vbInformation)
    Dim sMsg As String, sTitle As String, i As Integer, meta As Object
    Set meta = New ErrorsMeta

    With errs
        .Locn = Locn
        .iCodeLocal = iCode

        .ErrParam = ""
        If Not IsMissing(param) Then .ErrParam = param

        If Not meta.Init(meta, errs) Then Exit Sub
        If Not meta.ResolveCodesFromLocn(meta, errs) Then Exit Sub
        If Not meta.LoadFromLookup(meta, errs) Then Exit Sub
        If Not meta.Validate(meta, Locn) Then Exit Sub

        If meta.IsBaseNotFound Then
            sMsg = "Base error code not found for routine: " & Locn
        ElseIf meta.IsCodeNotFound Then
            sMsg = Locn & " Error code not found: " & .iCodeReport
        ElseIf meta.IsMalformed Then
            sMsg = meta.Message
        Else
            'Warning path is always user-facing
            sMsg = meta.ToUserMessage(meta, errs)
        End If
        
        'Add the message to accumulated error/warning string
        .UpdateMsgsAccum sMsg
        
        If .IsShowMsgs Then
            sTitle = "Warning/Information"
            i = MsgBox(sMsg, vbOKOnly + vbType, sTitle)
        End If
        
        'Re-initialize to be ready for a new error
        .Init errs.wkbkE, IsHandle:=.IsHandle
    End With
End Sub
'------------------------------------------------------------------------------------------------------
' Append new message to .Msgs_accum (with vbCrLf optionally)
' JDL 5/22/24
'
Public Sub UpdateMsgsAccum(msg)
    With errs
        .Msgs_accum = .Msgs_accum & msg & vbCrLf
    End With
End Sub
'-----------------------------------------------------------------------------------------------------
' Create and return a comment message for specified cell range
' IsReinitialize=False retains .iCodeLocal to comment multiple cells
'
' Created:   7/30/21 JDL     Modified: 12/20/22 for ErrorHandling Class; 8/29/23 elim errs arg
'                                       and add IsReinitialize arg; 10/30/25 add wkbkE Init arg
'
Sub LookupCommentMsg(rng, Locn, Optional IsReinitialize = True)
    Dim sMsg As String, meta As Object
    Set meta = New ErrorsMeta

    With errs
        .Locn = Locn

        If Not meta.Init(meta, errs) Then Exit Sub
        If Not meta.ResolveCodesFromLocn(meta, errs) Then Exit Sub
        If Not meta.LoadFromLookup(meta, errs) Then Exit Sub
        If Not meta.Validate(meta, Locn) Then Exit Sub

        If meta.IsBaseNotFound Then
            sMsg = "Base error code not found for routine: " & Locn
        ElseIf meta.IsCodeNotFound Then
            sMsg = Locn & " Error code not found: " & .iCodeReport
        ElseIf meta.IsMalformed Then
            sMsg = meta.Message
        Else
            'Comment path is always user-facing
            sMsg = meta.ToUserMessage(meta, errs)
        End If

        If Len(sMsg) > 0 Then AddComment rng, sMsg
        
        'Re-initialize to be ready for a new error
        If IsReinitialize Then .Init wkbkE:=errs.wkbkE, IsHandle:=.IsHandle
    End With
End Sub
'-----------------------------------------------------------------------------------------------------
'Assign error code based on Boolean argument and set ErrorHandling attributes
'
' Inputs: IsError [Boolean] Boolean expression; evaluates to True if error
'         iCode [Integer] local error code in case IsError = True
'                         "Local" refers to local to calling function or sub
'         ErrParam [Variant] optional parameter to report with message
'
' Created:  6/8/21 JDL   Modified 11/21/23 Clean up logic and Add Optional ErrParam argument
'
Function IsFail(IsError, iCode, Optional ErrParam) As Boolean
    IsFail = False
    If Not IsError Then Exit Function
    
    IsFail = True
    errs.iCodeLocal = iCode
    If Not IsMissing(ErrParam) Then errs.ErrParam = ErrParam
End Function
'-----------------------------------------------------------------------------------------------------
' Split a Setting or Error message into Title and Text parts
' JDL Modified 11/24/25 to default title to empty string
'
Sub setMsgTitleAndText(ByRef sMsg As String, ByRef sTitle As String)
    Dim aryMsg As Variant
    
    aryMsg = Split(sMsg, "|")
    If UBound(aryMsg) = 0 Then
        sTitle = ""
        Exit Sub
    End If
    
    'Set parts if array has two parts
    sTitle = aryMsg(0)
    sMsg = aryMsg(1)
End Sub
'-----------------------------------------------------------------------------------------------------
' Delete previous warnings/errors file and Reset errors/warnings setting
'
' Created:   7/7/21 JDL      Modified: 6/12/23 migrate to ErrorHandling class
'
Sub ResetWarningsAndErrors(wkbk, sSelf, Optional ByVal IsVal)
    Dim sPath As String, sString
    If IsMissing(IsVal) Then IsVal = False
    
    'Delete previous IsAuto WarningsAndErrors.txt file if any
    'sPath = wkbk.Path & "\" & sFileErrs        ' onedrive path mitigation
    sPath = ReadSetting(wkbk, "LocalPath")
    sPath = BuildPath(sPath, sFileErrs)
    If Len(Dir(sPath)) > 0 Then Kill sPath
    
    'Initialize a Setting for new warnings and errors
    sString = ParseNow & " " & sSelf
    If IsVal Then sString = sSelf & " Validation"
    UpdateSetting wkbk, sSettingErrs, sString & vbCrLf
End Sub
'-----------------------------------------------------------------------------------------------------
'Purpose:   Parse current timestamp into a fixed-length string
'
'Created:   2/15/22 JDL
'
'Output format: YYYYMMDD_HHMM_SS (length = 16 characters)
'
Function ParseNow() As String
    Dim dNow As Date, ary As Variant, s As Variant
    dNow = Now
    ary = Array(Year(dNow), Month(dNow), Day(dNow))
    For Each s In ary
        If Len(s) = 1 Then s = "0" & s
        ParseNow = ParseNow & s
    Next s
    ParseNow = ParseNow & "_"
    ary = Array(Hour(dNow), Minute(dNow))
    For Each s In ary
        If Len(s) = 1 Then s = "0" & s
        ParseNow = ParseNow & s
    Next s
    
    'Add Seconds
    s = Second(dNow)
    If Len(s) = 1 Then s = "0" & s
    ParseNow = ParseNow & "_" & s
End Function
'-----------------------------------------------------------------------------------------------------
' Write accumulated warnings and errors to file
'
'Inputs: wkbk [Workbook object]
'        IsVal [Boolean] True if validation run (used to create stable/repeatable output)
'
'Created:   7/7/21 JDL  Modified: 2/24/22 JDL IsVal = False by default if missing
'                                 4/5/22 ThiswWorkbook.Path as default; clarify sNow formula
'                                 6/12/23 move to ErrorHandling class
'
Sub WriteErrorsToFile(wkbk, IsVal)
    Dim sPathFile As String, sErrs As String, sNow As String
    
    If IsMissing(IsVal) Then IsVal = False
    sErrs = ReadSetting(wkbk, sSettingErrs)
    sPathFile = ReadSetting(wkbk, "LocalPath")
    If Len(sPathFile) < 1 Then sPathFile = ThisWorkbook.Path
        
    'Set suffix for filename: either datetime or "Val"
    If IsVal Then
        sNow = "_Val"
    ElseIf Len(sErrs) > 0 Then
        sNow = "_" & Left(sErrs, 16)
    End If
    sPathFile = sPathFile & "\" & sFileErrs & sNow & ".txt"
            
    Call WriteFile(sPathFile, sErrs)
    DeleteSetting wkbk, sSettingErrs
End Sub
