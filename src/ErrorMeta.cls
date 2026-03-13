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
'Version 3/12/26

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
  With Me
    .Code = iErrNotFound
    .Routine = ""
    .Message = ""
    .IsUserFacing = False
    .IsFound = False
    .IsBaseMessage = False
    .IsMalformed = False
  End With
End Sub
'--------------------------------------------------------------------------------------
' Set Errors_ table column ranges and data row extent
' JDL 3/12/26
'
Public Function SetTblELocations(meta, errs, colrngFunc, colrngMsg, colrngCode, colrngIMsg, rngRows) As Boolean
  On Error GoTo ErrorExit
  Dim c As Range

  SetTblELocations = False
  If errs Is Nothing Then GoTo ErrorExit

  With errs.wkbkE.Sheets(shtErrors)
    Set colrngCode = .Columns(1)
    Set colrngFunc = .Columns(3)
    Set colrngMsg = .Columns(4)
    Set colrngIMsg = .Columns(6)
  End With

  With errs.wkbkE.Sheets(shtErrors).Columns(1)
    Set c = .Cells(.Cells.Count).End(xlUp)
  End With
  Set rngRows = errs.wkbkE.Sheets(shtErrors).Range(errs.wkbkE.Sheets(shtErrors).Rows(2), c.EntireRow)

  SetTblELocations = True
  Exit Function

ErrorExit:
  SetTblELocations = False
End Function
'--------------------------------------------------------------------------------------
' Resolve and set base error code for a routine
' JDL 3/12/26
'
Public Function GetBaseErrCode(meta, errs, ByVal Locn As String) As Boolean
  On Error GoTo ErrorExit
  Dim r As Range, rngRows As Range
  Dim colrngCode As Range, colrngFunc As Range, colrngMsg As Range, colrngIMsg As Range

  GetBaseErrCode = False
  If errs Is Nothing Then GoTo ErrorExit

  errs.Locn = Locn
  If Not SetTblELocations(meta, errs, colrngFunc, colrngMsg, colrngCode, colrngIMsg, rngRows) Then GoTo ErrorExit

  Set r = rngMultiKeyBasic(rngRows, Array(colrngFunc, colrngMsg), Array(Locn, sErrBase))
  errs.iCodeBase = iErrNotFound
  If Not r Is Nothing Then
    errs.iCodeBase = CLng(Intersect(r, colrngCode).Value)
    GetBaseErrCode = True
  End If

  Exit Function

ErrorExit:
  GetBaseErrCode = False
End Function
'--------------------------------------------------------------------------------------
' Set explicit not-found state
' JDL 3/13/26
'
Private Sub SetNotFound(meta, Optional ByVal Routine As String = "")
  With meta
    .Code = 0
    .Routine = Routine
    .Message = "Msg Not Found"
    .IsUserFacing = False
    .IsFound = False
    .IsBaseMessage = False
    .IsMalformed = False
  End With
End Sub
'--------------------------------------------------------------------------------------
' Parse user-facing Boolean-like field and malformed status
' JDL 3/12/26
'
Private Function ParseIsUserFacing(ByVal valRaw As Variant, ByRef IsMalformed As Boolean) As Boolean
  Dim sBool As String

  IsMalformed = False
  sBool = LCase(Trim(CStr(valRaw)))
  If Len(sBool) = 0 Then Exit Function

  If sBool = "true" Or sBool = "false" Then
    ParseIsUserFacing = CBool(sBool)
  Else
    ParseIsUserFacing = False
    IsMalformed = True
  End If
End Function
'--------------------------------------------------------------------------------------
' Load metadata for current error context
' JDL 3/11/26
'
Public Function LoadFromLookup(meta, errs) As Boolean
  On Error GoTo ErrorExit
  Dim rng As Range, rngRows As Range
  Dim colrngCode As Range, colrngFunc As Range, colrngMsg As Range, colrngIMsg As Range

  LoadFromLookup = True
  If errs Is Nothing Then GoTo ErrorExit

  If Not SetTblELocations(meta, errs, colrngFunc, colrngMsg, colrngCode, colrngIMsg, rngRows) Then GoTo ErrorExit

  errs.iCodeReport = errs.iCodeBase + errs.iCodeLocal
  Set rng = rngMultiKeyBasic(rngRows, Array(colrngCode), Array(errs.iCodeReport))

  If rng Is Nothing Then
    SetNotFound meta
    Exit Function
  End If

  With meta
    .Code = Intersect(rng, colrngCode).Value
    .Routine = Intersect(rng, colrngFunc).Value
    .Message = Intersect(rng, colrngMsg).Value
    .IsFound = True
    .IsBaseMessage = (.Message = sErrBase)
    .IsUserFacing = ParseIsUserFacing(Intersect(rng, colrngIMsg).Value, .IsMalformed)
  End With

  Exit Function

ErrorExit:
  LoadFromLookup = False
End Function
'--------------------------------------------------------------------------------------
' Validate loaded metadata and normalize malformed message
' JDL 3/11/26
'
Public Function Validate(meta, Optional ByVal Locn As String = "") As Boolean
  On Error GoTo ErrorExit

  Validate = True

  With meta
    If Not .IsFound Then Exit Function

    If .Code = iErrNotFound Or Len(Trim(.Routine)) = 0 Or Len(Trim(.Message)) = 0 Then
      .IsMalformed = True
    End If

    If .IsMalformed Then
      .Message = "Malformed Errors_ Row for " & Locn
      .IsUserFacing = False
    End If
  End With

  Exit Function

ErrorExit:
  Validate = False
End Function
'--------------------------------------------------------------------------------------
' Build user-facing message text
' JDL 3/11/26
'
Public Function ToUserMessage(meta, Optional errParam As String = "") As String
  Dim sMsg As String

  sMsg = meta.Message
  If meta.IsBaseMessage Then sMsg = sVBAErr
  If Len(errParam) > 0 Then sMsg = sMsg & errParam
  ToUserMessage = sMsg
End Function
'--------------------------------------------------------------------------------------
' Build developer-facing message text
' JDL 3/11/26
'
Public Function ToDeveloperMessage(meta, errCodeReport As Long, Optional errParam As String = "") As String
  Dim sMsg As String

  sMsg = "Error " & errCodeReport & "; in sub or function, "
  If Len(meta.Routine) > 0 Then sMsg = sMsg & meta.Routine & vbCrLf

  With meta
    If .IsBaseMessage Then
      sMsg = sMsg & sVBAErr & errParam & vbCrLf
    ElseIf Len(.Message) > 0 Then
      sMsg = sMsg & .Message & errParam & vbCrLf
    End If
  End With

  ToDeveloperMessage = sMsg
End Function

