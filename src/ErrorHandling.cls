VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ErrorHandling"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'Version 3/12/26
'Class for error handling in cascaded function architecture
Public IsHandle As Boolean 'Toggle error handling on/off
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
Public ErrMsg As String 'Error message string --constructed by AppendErrMsg() class method
Public Msgs_accum As String 'Accumulated error and warning messages
Public IsShowMsgs As Boolean 'Flag to turn off showing MsgBox messages (for unit testing etc.)

Public wkbkE As Workbook 'Workbook containing shtErrors lookup table (optionally can be add-in)
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
        .IsShowMsgs = True
    End With
End Sub
'-----------------------------------------------------------------------------------------------------
' Record new error or append stack track message
' * Call with Boolean CallingFunction argument to set that to False in case of error
' * Call without CallingFunction arg if driver routine; triggers reporting of error that 
'   occurred either in driver or nested functions (i.e. stack tracing) 
' * error message will be reported if .IsShowMsgs = True (SetErrs sets True by default when errs
'   initialized by driver sub; sets to False by default when errs initialized by calling SetErrs
'   directly from non-driver Boolean function as in call from unit test)
'
' Created:  2/17/21 JDL      Modified: 3/12/26
'
Sub RecordErr(Locn, Optional ByRef CallingFunction)
    Dim IsDriverCall As Boolean, sMsgSuffix As String

    IsDriverCall = IsMissing(CallingFunction)
    If Not IsDriverCall Then CallingFunction = False
    
    With errs
        .Locn = Locn

        'Add trace for calling function's location for previous error in nested function
        If Not .IsNewErr And Not .IsUserFacing Then sMsgSuffix = "Called by " & Locn

        If Not AppendErrMsg(sMsgSuffix) Then Exit Sub
        
        'IsShowMsgs gates reporting for both driver and testing flows.
        If .IsShowMsgs And (IsDriverCall Or .IsTesting) Then .ReportMsg
    End With
End Sub
'-----------------------------------------------------------------------------------------------------
' Append error message updates onto ErrMsg'
' Created:   2/17/21 JDL      Modified: 3/12/26
'
Function AppendErrMsg(sMsgSuffix) As Boolean
    With errs
        If .IsNewErr Then
            AppendErrMsg = ProcessRootErrMsg()
            Exit Function
        ElseIf Not .IsUserFacing Then
            .ErrMsg = .ErrMsg & sMsgSuffix & vbCrLf
        End If
    End With
    AppendErrMsg = True
End Function

'-----------------------------------------------------------------------------------------------------
' Build and append root error message using ErrorMeta lookup state
' JDL 3/12/26
'
Private Function ProcessRootErrMsg() As Boolean
    Dim sMsgNew As String, meta As ErrorMeta

    ProcessRootErrMsg = False

    With errs
        'Set base/report code for root lookup
        Set meta = New ErrorMeta
        If Not meta.GetBaseErrCode(meta, errs, .Locn) Then .iCodeBase = iErrNotFound
        .iCodeReport = .iCodeBase + .iCodeLocal

        If Not meta.LoadFromLookup(meta, errs) Then Exit Function
        If Not meta.Validate(meta, .Locn) Then Exit Function

        'Add additional line feed if there are already accumulated warning(s)
        If Len(.ErrMsg) > 0 Then .ErrMsg = .ErrMsg & vbCrLf

        If Not meta.IsFound Then
            .IsUserFacing = False
            If .iCodeBase <> iErrNotFound Then
                sMsgNew = .Locn & " Error code not found: " & .iCodeBase
            Else
                sMsgNew = "Base error code not found for routine: " & .Locn
            End If
        Else
            .IsUserFacing = meta.IsUserFacing
            If .IsUserFacing Then
                sMsgNew = meta.ToUserMessage(meta, .ErrParam)
            Else
                sMsgNew = meta.ToDeveloperMessage(meta, .iCodeReport, .ErrParam)
            End If
        End If

        .ErrMsg = .ErrMsg & sMsgNew
        .IsNewErr = False
    End With

    ProcessRootErrMsg = True
End Function
'-----------------------------------------------------------------------------------------------------
' Lookup metadata from Errors table
'
' Created:  2/17/21 JDL      Modified: 12/20/22 for ErrorHandling Class
'
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
    End With
End Function

'-----------------------------------------------------------------------------------------------------
' Configure and report warnings
'
' Created:  6/8/21 JDL      Modified: 6/8/23 Add & errs.ErrParam to sMsg
'                                     4/1/24 fix bug with call .GetBaseErrCode should be Boolean
'                                    5/22/24 add .Msgs_accum and .IsShowMsgs
'                                    12/3/25 Fix bug; add .wkbkE arg to .Init call at end
'
Sub ReportWarningMsg(iCode, Locn, Optional param)
    Dim sMsg As String, sTitle As String, i As Integer, meta As ErrorMeta
    With errs
        .Locn = Locn
        .iCodeLocal = iCode
        Set meta = New ErrorMeta

        If Not meta.GetBaseErrCode(meta, errs, Locn) Then
            sMsg = "Warning msg lookup failed; base row not found for routine: " & Locn
            GoTo ReportAndReset
        End If

        If Not meta.LoadFromLookup(meta, errs) Then Exit Sub
        If Not meta.Validate(meta, Locn) Then Exit Sub

        If Not meta.IsFound Then
            sMsg = "Warning message not found for code " & .iCodeReport & " in routine: " & Locn
        Else
            sMsg = meta.Message
        End If

        sMsg = sMsg & errs.ErrParam
        If Not IsMissing(param) Then sMsg = sMsg & param

ReportAndReset:
        'Add the message to accumulated error/warning string
        .UpdateMsgsAccum sMsg
        
        If .IsShowMsgs Then
            sTitle = "Warning/Information"
            i = MsgBox(sMsg, vbOKOnly + vbInformation, sTitle)
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
    Dim sMsg As String, meta As ErrorMeta
    With errs
        .Locn = Locn
        Set meta = New ErrorMeta
        sMsg = "Comment msg lookup failed for routine: " & Locn
        
        'Get base and message metadata for comment text
        If meta.GetBaseErrCode(meta, errs, Locn) Then
            If Not meta.LoadFromLookup(meta, errs) Then
                Exit Sub
            End If
            If Not meta.Validate(meta, Locn) Then
                sMsg = "Comment msg validation failed for routine: " & Locn
            ElseIf meta.IsFound Then
                sMsg = meta.Message
            End If
        End If
        AddComment rng, sMsg

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
' Create and return user response from msgbox prompt
'
' Created:   7/20/21 JDL; 6/7/23 adapt to ErrorHandling Class
'                         7/10/23 fix bug by eliminate errs as argument
'                        10/16/25 add optional param argument and ConvertVbCrLfToConcat
'
' Inputs:   Locn [String] name of calling routine (for Errors_ lookup)
'         iCode [Integer] local error code ( iCode + errs.iCodeBase = errs.iCodeReport)
'         vbType [Integer] VBA MsgBox function buttons argument (e.g. vbCritical + vbOK = 17)
'
Function ShowMessage(Locn As String, iCode As Integer, vbType As Integer, Optional param) As Long
    Dim iMsg As Integer, sMsg As String, sTitle As String, meta As ErrorMeta
    
    'Default message and title
    sMsg = "Base Error Code Not Found: " & Locn
    sTitle = "Missing Base Code"
    
    'Look up message based on Locn and iCode
    With errs
        .Locn = Locn
        .iCodeLocal = iCode
        Set meta = New ErrorMeta

        If meta.GetBaseErrCode(meta, errs, Locn) Then
            If Not meta.LoadFromLookup(meta, errs) Then GoTo ShowExit
            If Not meta.Validate(meta, Locn) Then GoTo ShowExit

            If meta.IsFound Then
                sMsg = meta.Message
            Else
                sMsg = "Message not found for code " & .iCodeReport & " in routine: " & Locn
            End If
            
            'Set Title and Msg based by splitting
            .setMsgTitleAndText sMsg, sTitle
        
            'Append param if any
            If Not IsMissing(param) Then sMsg = sMsg & param
            
            'Check for line breaks and convert
            sMsg = ConvertVbCrLfToConcat(sMsg)
        End If
    End With
ShowExit:
    ShowMessage = MsgBox(sMsg, vbType, sTitle)
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



