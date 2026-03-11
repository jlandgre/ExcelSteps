VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ErrorMeta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'Version 3/11/26

Private pCode As Long
Private pRoutine As String
Private pMessage As String
Private pIsUserFacing As Boolean
Private pIsFound As Boolean
Private pIsBaseMessage As Boolean
Private pIsMalformed As Boolean

'--------------------------------------------------------------------------------------
' Initialize ErrorMeta defaults
' JDL 3/11/26
'
Private Sub Class_Initialize()
  pCode = iErrNotFound
  pRoutine = ""
  pMessage = ""
  pIsUserFacing = False
  pIsFound = False
  pIsBaseMessage = False
  pIsMalformed = False
End Sub
'--------------------------------------------------------------------------------------
' Load metadata for current error context
' JDL 3/11/26
'
Public Function LoadFromLookup(errsObj) As Boolean
  On Error GoTo ErrorExit
  Dim aryMetaData As Variant, sBool As String

  LoadFromLookup = True
  If errsObj Is Nothing Then GoTo ErrorExit

  aryMetaData = errsObj.aryErrLookup()
  If aryMetaData(0) = iErrNotFound Then
    pCode = iErrNotFound
    pRoutine = ""
    pMessage = "Msg Not Found"
    pIsUserFacing = False
    pIsFound = False
    pIsBaseMessage = False
    pIsMalformed = False
    Exit Function
  End If

  pCode = aryMetaData(0)
  pRoutine = aryMetaData(1)
  pMessage = aryMetaData(2)
  pIsFound = True
  pIsBaseMessage = (pMessage = sErrBase)
  pIsMalformed = False

  sBool = LCase(Trim(CStr(aryMetaData(3))))
  If Len(sBool) = 0 Then
    pIsUserFacing = False
  ElseIf sBool = "true" Or sBool = "false" Then
    pIsUserFacing = CBool(sBool)
  Else
    pIsUserFacing = False
    pIsMalformed = True
  End If

  Exit Function

ErrorExit:
  LoadFromLookup = False
End Function
'--------------------------------------------------------------------------------------
' Validate loaded metadata and normalize malformed message
' JDL 3/11/26
'
Public Function Validate(Optional Locn As String = "") As Boolean
  On Error GoTo ErrorExit

  Validate = True

  If Not pIsFound Then Exit Function

  If pCode = iErrNotFound Or Len(Trim(pRoutine)) = 0 Or Len(Trim(pMessage)) = 0 Then
    pIsMalformed = True
  End If

  If pIsMalformed Then
    pMessage = "Malformed Errors_ Row for " & Locn
    pIsUserFacing = False
  End If

  Exit Function

ErrorExit:
  Validate = False
End Function
'--------------------------------------------------------------------------------------
' Build user-facing message text
' JDL 3/11/26
'
Public Function ToUserMessage(Optional errParam As String = "") As String
  Dim sMsg As String

  sMsg = pMessage
  If pIsBaseMessage Then sMsg = sVBAErr
  If Len(errParam) > 0 Then sMsg = sMsg & errParam
  ToUserMessage = sMsg
End Function
'--------------------------------------------------------------------------------------
' Build developer-facing message text
' JDL 3/11/26
'
Public Function ToDeveloperMessage(errCodeReport As Long, Optional errParam As String = "") As String
  Dim sMsg As String

  sMsg = "Error " & errCodeReport & "; in sub or function, "
  If Len(pRoutine) > 0 Then sMsg = sMsg & pRoutine & vbCrLf

  If pIsBaseMessage Then
    sMsg = sMsg & sVBAErr & errParam & vbCrLf
  ElseIf Len(pMessage) > 0 Then
    sMsg = sMsg & pMessage & errParam & vbCrLf
  End If

  ToDeveloperMessage = sMsg
End Function
'--------------------------------------------------------------------------------------
Public Property Get Code() As Long
  Code = pCode
End Property

Public Property Let Code(ByVal val As Long)
  pCode = val
End Property

Public Property Get Routine() As String
  Routine = pRoutine
End Property

Public Property Let Routine(ByVal val As String)
  pRoutine = val
End Property

Public Property Get Message() As String
  Message = pMessage
End Property

Public Property Let Message(ByVal val As String)
  pMessage = val
  pIsBaseMessage = (val = sErrBase)
End Property

Public Property Get IsUserFacing() As Boolean
  IsUserFacing = pIsUserFacing
End Property

Public Property Let IsUserFacing(ByVal val As Boolean)
  pIsUserFacing = val
End Property

Public Property Get IsFound() As Boolean
  IsFound = pIsFound
End Property

Public Property Let IsFound(ByVal val As Boolean)
  pIsFound = val
End Property

Public Property Get IsBaseMessage() As Boolean
  IsBaseMessage = pIsBaseMessage
End Property

Public Property Let IsBaseMessage(ByVal val As Boolean)
  pIsBaseMessage = val
End Property

