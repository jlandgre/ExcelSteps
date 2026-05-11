VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "tblUnstructured"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'Version 5/7/26

Public wkbk As Workbook
Public sht As String
Public wksht As Worksheet
Public rngTable As Range
'-----------------------------------------------------------------------------------------
' Initialize unstructured table wayfinding attributes
' JDL 5/7/26
'
Public Function Init(tblU, wkbk As Workbook, sht As String, Optional rngTable As Range) As Boolean
  SetErrs Init: If errs.IsHandle Then On Error GoTo ErrorExit

  With tblU
    Set .wkbk = wkbk
    .sht = sht
    Set .wksht = .wkbk.Worksheets(.sht)
    If IsMissing(rngTable) Then
      Set .rngTable = .wksht.UsedRange
    Else
      Set .rngTable = rngTable
    End If
  End With
  Exit Function

ErrorExit:
  errs.RecordErr "tblUnstructured.Init", Init
End Function

