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

Public Code As Long
Public Routine As String
Public Message As String
Public IsUserFacing As Boolean
Public IsFound As Boolean
Public IsBaseMessage As Boolean
Public IsMalformed As Boolean

'--------------------------------------------------------------------------------------
' Initialize ErrorMeta defaults
' JDL 3/11/26
'
Private Sub Class_Initialize()
  Code = iErrNotFound
  Routine = ""
  Message = ""
  IsUserFacing = False
  IsFound = False
  IsBaseMessage = False
  IsMalformed = False
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
    Code = iErrNotFound
    Routine = ""
    Message = "Msg Not Found"
    IsUserFacing = False
    IsFound = False
    IsBaseMessage = False
    IsMalformed = False
    Exit Function
  End If

  Code = aryMetaData(0)
  Routine = aryMetaData(1)
  Message = aryMetaData(2)
  IsFound = True
  IsBaseMessage = (Message = sErrBase)
  IsMalformed = False

  sBool = LCase(Trim(CStr(aryMetaData(3))))
  If Len(sBool) = 0 Then
    IsUserFacing = False
  ElseIf sBool = "true" Or sBool = "false" Then
    IsUserFacing = CBool(sBool)
  Else
    IsUserFacing = False
    IsMalformed = True
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

  If Not IsFound Then Exit Function

  If Code = iErrNotFound Or Len(Trim(Routine)) = 0 Or Len(Trim(Message)) = 0 Then
    IsMalformed = True
  End If

  If IsMalformed Then
    Message = "Malformed Errors_ Row for " & Locn
    IsUserFacing = False
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

  sMsg = Message
  If IsBaseMessage Then sMsg = sVBAErr
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
  If Len(Routine) > 0 Then sMsg = sMsg & Routine & vbCrLf

  If IsBaseMessage Then
    sMsg = sMsg & sVBAErr & errParam & vbCrLf
  ElseIf Len(Message) > 0 Then
    sMsg = sMsg & Message & errParam & vbCrLf
  End If

  ToDeveloperMessage = sMsg
End Function

