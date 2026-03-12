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
' Load metadata for current error context
' JDL 3/11/26
'
Public Function LoadFromLookup(meta, errs) As Boolean
  On Error GoTo ErrorExit
  Dim sBool As String, rng As Range, rngRows As Range, c As Range
  Dim colrngCode As Range, colrngFunc As Range, colrngMsg As Range, colrngIMsg As Range
  Dim wksht As Worksheet

  LoadFromLookup = True
  If errs Is Nothing Then GoTo ErrorExit

  Set wksht = errs.wkbkE.Sheets(shtErrors)
  Set colrngCode = wksht.Columns(1)
  Set colrngFunc = wksht.Columns(3)
  Set colrngMsg = wksht.Columns(4)
  Set colrngIMsg = wksht.Columns(6)

  With wksht.Columns(1)
    Set c = .Cells(.Cells.Count).End(xlUp)
  End With
  Set rngRows = wksht.Range(wksht.Rows(2), c.EntireRow)

  errs.iCodeReport = errs.iCodeBase + errs.iCodeLocal
  Set rng = rngMultiKeyBasic(rngRows, Array(colrngCode), Array(errs.iCodeReport))

  If rng Is Nothing Then
    With meta
      .Code = iErrNotFound
      .Routine = ""
      .Message = "Msg Not Found"
      .IsUserFacing = False
      .IsFound = False
      .IsBaseMessage = False
      .IsMalformed = False
    End With
    Exit Function
  End If

  With meta
    .Code = Intersect(rng, colrngCode).Value
    .Routine = Intersect(rng, colrngFunc).Value
    .Message = Intersect(rng, colrngMsg).Value
    .IsFound = True
    .IsBaseMessage = (.Message = sErrBase)
    .IsMalformed = False

    sBool = LCase(Trim(CStr(Intersect(rng, colrngIMsg).Value)))
    If Len(sBool) = 0 Then
      .IsUserFacing = False
    ElseIf sBool = "true" Or sBool = "false" Then
      .IsUserFacing = CBool(sBool)
    Else
      .IsUserFacing = False
      .IsMalformed = True
    End If
  End With

  Exit Function

ErrorExit:
  LoadFromLookup = False
End Function
'--------------------------------------------------------------------------------------
' Validate loaded metadata and normalize malformed message
' JDL 3/11/26
'
Public Function Validate(meta, Optional Locn As String = "") As Boolean
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

