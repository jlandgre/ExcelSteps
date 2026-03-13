VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ErrorsMeta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'Version 3/13/26

Public Code As Long
Public Routine As String
Public Message As String
Public IsUserFacing As Boolean
Public IsBaseNotFound As Boolean
Public IsCodeNotFound As Boolean
Public IsMalformed As Boolean
Public tblErrs As Object
'--------------------------------------------------------------------------------------
' Initialize metadata state and provision Errors_ table wayfinding
' JDL 3/13/26
'
Public Function Init(meta, errsObj) As Boolean
  On Error GoTo ErrorExit

  Set meta.tblErrs = New tblRowsCols
  If Not meta.tblErrs.Provision(meta.tblErrs, errsObj.wkbkE, False, _
    sht:=shtErrors, IsSetColRngs:=True) Then GoTo ErrorExit

  meta.Code = 0
  meta.Routine = ""
  meta.Message = ""
  meta.IsUserFacing = False
  meta.IsBaseNotFound = False
  meta.IsCodeNotFound = False
  meta.IsMalformed = False
  Init = True
  Exit Function

ErrorExit:
  Init = False
End Function
'--------------------------------------------------------------------------------------
' Resolve base code and computed report code from Locn
' JDL 3/13/26
'
Public Function ResolveCodesFromLocn(meta, errsObj) As Boolean
  On Error GoTo ErrorExit
  Dim rngBase As Range, valCode As Variant

  Set rngBase = rngMultiKeyBasic(meta.tblErrs.rngRows, _
    Array(meta.tblErrs.colrngRoutine, meta.tblErrs.colrngMsg), _
    Array(errsObj.Locn, sErrBase))

  If rngBase Is Nothing Then
    meta.IsBaseNotFound = True
    errsObj.iCodeBase = 0
    errsObj.iCodeReport = 0
    ResolveCodesFromLocn = True
    Exit Function
  End If

  valCode = Intersect(rngBase, meta.tblErrs.colrngCode).Value
  If Not IsNumeric(valCode) Then GoTo ErrorExit

  errsObj.iCodeBase = CLng(valCode)
  errsObj.iCodeReport = errsObj.iCodeBase + errsObj.iCodeLocal
  ResolveCodesFromLocn = True
  Exit Function

ErrorExit:
  ResolveCodesFromLocn = False
End Function
'--------------------------------------------------------------------------------------
' Load row metadata by computed report code
' JDL 3/13/26
'
Public Function LoadFromLookup(meta, errsObj) As Boolean
  On Error GoTo ErrorExit
  Dim rngCode As Range, valUser As Variant, sUser As String

  If meta.IsBaseNotFound Then
    LoadFromLookup = True
    Exit Function
  End If

  Set rngCode = rngMultiKeyBasic(meta.tblErrs.rngRows, _
    Array(meta.tblErrs.colrngCode), Array(errsObj.iCodeReport))

  If rngCode Is Nothing Then
    meta.IsCodeNotFound = True
    LoadFromLookup = True
    Exit Function
  End If

  meta.Code = CLng(Intersect(rngCode, meta.tblErrs.colrngCode).Value)
  meta.Routine = CStr(Intersect(rngCode, meta.tblErrs.colrngRoutine).Value)
  meta.Message = CStr(Intersect(rngCode, meta.tblErrs.colrngMsg).Value)

  valUser = Intersect(rngCode, meta.tblErrs.colrngUserFace).Value
  If VarType(valUser) = vbBoolean Then
    meta.IsUserFacing = CBool(valUser)
  Else
    sUser = LCase(Trim(CStr(valUser)))
    If sUser = "true" Then
      meta.IsUserFacing = True
    ElseIf sUser = "false" Then
      meta.IsUserFacing = False
    ElseIf IsNumeric(sUser) Then
      meta.IsUserFacing = (CLng(sUser) <> 0)
    Else
      meta.IsMalformed = True
      meta.IsUserFacing = False
    End If
  End If

  LoadFromLookup = True
  Exit Function

ErrorExit:
  LoadFromLookup = False
End Function
'--------------------------------------------------------------------------------------
' Validate metadata completeness and normalize malformed state
' JDL 3/13/26
'
Public Function Validate(meta, ByVal Locn As String) As Boolean
  On Error GoTo ErrorExit

  If meta.IsBaseNotFound Or meta.IsCodeNotFound Then
    Validate = True
    Exit Function
  End If

  If meta.Code <= 0 Then meta.IsMalformed = True
  If Len(Trim(meta.Routine)) < 1 Then meta.IsMalformed = True
  If Len(Trim(meta.Message)) < 1 Then meta.IsMalformed = True

  If meta.IsMalformed Then
    meta.Message = "Malformed Errors_ row for routine: " & Locn
    meta.IsUserFacing = False
  End If

  Validate = True
  Exit Function

ErrorExit:
  Validate = False
End Function
'--------------------------------------------------------------------------------------
' Return user-facing message segment
' JDL 3/13/26
'
Public Function ToUserMessage(meta, errsObj) As String
  If meta.IsBaseNotFound Then
    ToUserMessage = "Base error code not found for routine: " & errsObj.Locn
    Exit Function
  End If
  If meta.IsCodeNotFound Then
    ToUserMessage = errsObj.Locn & " Error code not found: " & errsObj.iCodeReport
    Exit Function
  End If
  If meta.IsMalformed Then
    ToUserMessage = meta.Message
    Exit Function
  End If

  ToUserMessage = meta.Message
  If Len(errsObj.ErrParam) > 0 Then ToUserMessage = ToUserMessage & errsObj.ErrParam
End Function
'--------------------------------------------------------------------------------------
' Return developer-facing message segment
' JDL 3/13/26
'
Public Function ToDeveloperMessage(meta, errsObj) As String
  Dim sBody As String

  If meta.IsBaseNotFound Then
    ToDeveloperMessage = "Base error code not found for routine: " & errsObj.Locn
    Exit Function
  End If
  If meta.IsCodeNotFound Then
    ToDeveloperMessage = errsObj.Locn & " Error code not found: " & errsObj.iCodeReport
    Exit Function
  End If
  If meta.IsMalformed Then
    ToDeveloperMessage = meta.Message
    Exit Function
  End If

  sBody = meta.Message
  If sBody = sErrBase Then sBody = sVBAErr

  ToDeveloperMessage = "Error " & errsObj.iCodeReport & "; in sub or function, "
  If Len(meta.Routine) > 0 Then ToDeveloperMessage = ToDeveloperMessage & meta.Routine & vbCrLf
  If Len(sBody) > 0 Then ToDeveloperMessage = ToDeveloperMessage & sBody & errsObj.ErrParam & vbCrLf
End Function
