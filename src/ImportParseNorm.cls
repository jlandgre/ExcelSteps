VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ImportParseNorm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'Version 5/7/26

Public colinfo As Object
Public files As Object
Public curTbl As String
Public dParamsImport As Object
Public dParamsParse As Object
Public tblRaw As Object
Public tblParsed As Object
Public tblNorm As Object
'-----------------------------------------------------------------------------------------
' Run import, parse, and normalization sequence for a specified table
' JDL 5/7/26
'
Public Function ImportParseNormProcedure(importtbl, colinfo, files As Object, curTbl As String, _
  dParamsImport As Object, dParamsParse As Object) As Boolean
  SetErrs ImportParseNormProcedure: If errs.IsHandle Then On Error GoTo ErrorExit

  If Not importtbl.Init(importtbl, colinfo, files, curTbl, dParamsImport, dParamsParse) Then GoTo ErrorExit
  If Not importtbl.OpenAndValidateRawProcedure(importtbl) Then GoTo ErrorExit
  If Not importtbl.ParseRawProcedure(importtbl) Then GoTo ErrorExit
  If Not importtbl.NormalizeParsedProcedure(importtbl) Then GoTo ErrorExit
  Exit Function

ErrorExit:
  errs.RecordErr "ImportParseNorm.ImportParseNormProcedure", ImportParseNormProcedure
End Function
'-----------------------------------------------------------------------------------------
' Initialize attributes and ensure colinfo is set for current table
' JDL 5/7/26
'
Public Function Init(importtbl, colinfo, files As Object, curTbl As String, _
  dParamsImport As Object, dParamsParse As Object) As Boolean
  SetErrs Init: If errs.IsHandle Then On Error GoTo ErrorExit
  Dim colinfoLocal As Object

  If errs.IsFail(files Is Nothing, 1) Then GoTo ErrorExit
  If errs.IsFail(Len(files.pfImportFile) = 0, 2) Then GoTo ErrorExit
  If errs.IsFail(dParamsImport Is Nothing, 3) Then GoTo ErrorExit
  If errs.IsFail(dParamsParse Is Nothing, 4) Then GoTo ErrorExit
  If errs.IsFail(Len(curTbl) = 0, 5) Then GoTo ErrorExit

  Set colinfoLocal = colinfo
  If colinfoLocal Is Nothing Then Set colinfoLocal = New ColInfo
  If colinfoLocal.tbl Is Nothing Then
    If Not colinfoLocal.Init(colinfoLocal, files) Then GoTo ErrorExit
  End If
  If Not colinfoLocal.SetCurTbl(colinfoLocal, curTbl) Then GoTo ErrorExit

  With importtbl
    Set .colinfo = colinfoLocal
    Set .files = files
    .curTbl = curTbl
    Set .dParamsImport = dParamsImport
    Set .dParamsParse = dParamsParse
    Set .tblRaw = Nothing
    Set .tblParsed = Nothing
    Set .tblNorm = Nothing
  End With
  Set colinfo = colinfoLocal
  Exit Function

ErrorExit:
  errs.RecordErr "ImportParseNorm.Init", Init
End Function
'-----------------------------------------------------------------------------------------
' Open raw data and run raw structure checks and fill operations
' JDL 5/7/26
'
Public Function OpenAndValidateRawProcedure(importtbl) As Boolean
  SetErrs OpenAndValidateRawProcedure: If errs.IsHandle Then On Error GoTo ErrorExit

  If Not importtbl.OpenRawData(importtbl) Then GoTo ErrorExit
  If Not importtbl.ValidateRawStructure(importtbl) Then GoTo ErrorExit
  If Not importtbl.FillMissingVals(importtbl) Then GoTo ErrorExit
  Exit Function

ErrorExit:
  errs.RecordErr "ImportParseNorm.OpenAndValidateRawProcedure", OpenAndValidateRawProcedure
End Function
'-----------------------------------------------------------------------------------------
' Open raw file and provision raw table object in a temporary workbook copy
' JDL 5/7/26
'
Public Function OpenRawData(importtbl) As Boolean
  SetErrs OpenRawData: If errs.IsHandle Then On Error GoTo ErrorExit
  Dim wkbkOpened As Workbook, wkbkTemp As Workbook
  Dim tblAlias As Object, tblUAlias As Object, shtRaw As String

  If Not OpenFile(importtbl.files.pfImportFile, wkbkOpened) Then GoTo ErrorExit
  If errs.IsFail(wkbkOpened Is Nothing, 1, importtbl.files.pfImportFile) Then GoTo ErrorExit

  wkbkOpened.Worksheets(1).Copy
  Set wkbkTemp = ActiveWorkbook
  wkbkOpened.Close False
  shtRaw = wkbkTemp.Worksheets(1).Name

  If LCase$(GetDictText(importtbl.dParamsParse, "RawShape", "rowscols")) = "rowscols" Then
    Set importtbl.tblRaw = New tblRowsCols
    Set tblAlias = importtbl.tblRaw
    If Not tblAlias.Provision(tblAlias, wkbkTemp, False, sht:=shtRaw, IsSetColRngs:=False) Then GoTo ErrorExit
  Else
    Set importtbl.tblRaw = New tblUnstructured
    Set tblUAlias = importtbl.tblRaw
    If Not tblUAlias.Init(tblUAlias, wkbkTemp, shtRaw) Then GoTo ErrorExit
  End If
  Exit Function

ErrorExit:
  errs.RecordErr "ImportParseNorm.OpenRawData", OpenRawData
End Function
'-----------------------------------------------------------------------------------------
' Validate that required raw columns are present for current colinfo table
' JDL 5/7/26
'
Public Function ValidateRawStructure(importtbl) As Boolean
  SetErrs ValidateRawStructure: If errs.IsHandle Then On Error GoTo ErrorExit
  Dim rngVarNormCol As Range, rngVarRawCol As Range, rngRow As Range, rngHdr As Range
  Dim valNorm As Variant, valRaw As String

  If errs.IsFail(importtbl.tblRaw Is Nothing, 1) Then GoTo ErrorExit
  If errs.IsFail(TypeName(importtbl.tblRaw) <> "tblRowsCols", 2) Then GoTo ErrorExit

  Set rngVarNormCol = importtbl.colinfo.tbl.rngTblHeaderVal(importtbl.colinfo.tbl, "VarNameNorm").EntireColumn
  Set rngVarRawCol = importtbl.colinfo.tbl.rngTblHeaderVal(importtbl.colinfo.tbl, "VarNameRaw").EntireColumn

  For Each rngRow In importtbl.colinfo.rngRowsCurTbl.Rows
    valNorm = Intersect(rngRow, rngVarNormCol).Value2
    If Len(CStr(valNorm)) > 0 Then
      valRaw = CStr(Intersect(rngRow, rngVarRawCol).Value2)
      If errs.IsFail(Len(valRaw) = 0, 3, CStr(valNorm)) Then GoTo ErrorExit

      Set rngHdr = importtbl.tblRaw.rngTblHeaderVal(importtbl.tblRaw, valRaw)
      If errs.IsFail(rngHdr Is Nothing, 4, valRaw) Then GoTo ErrorExit
    End If
  Next rngRow
  Exit Function

ErrorExit:
  errs.RecordErr "ImportParseNorm.ValidateRawStructure", ValidateRawStructure
End Function
'-----------------------------------------------------------------------------------------
' Apply FillVals mappings from colinfo rows to raw data columns (archived version)
' JDL 5/7/26
'
Public Function FillMissingVals_prev(importtbl) As Boolean
  SetErrs FillMissingVals_prev: If errs.IsHandle Then On Error GoTo ErrorExit
  Dim rngVarRawCol As Range, rngFillValsCol As Range, rngRow As Range
  Dim rngRawHdr As Range, rngDataCol As Range, rngCell As Range
  Dim sFill As String, sRaw As String, sKey As Variant, valFill As Variant
  Dim dict As Object, aryKeys As Variant

  If errs.IsFail(importtbl.tblRaw Is Nothing, 1) Then GoTo ErrorExit
  If errs.IsFail(TypeName(importtbl.tblRaw) <> "tblRowsCols", 2) Then GoTo ErrorExit

  Set rngVarRawCol = importtbl.colinfo.tbl.rngTblHeaderVal(importtbl.colinfo.tbl, "VarNameRaw").EntireColumn
  Set rngFillValsCol = importtbl.colinfo.tbl.rngTblHeaderVal(importtbl.colinfo.tbl, "FillVals").EntireColumn

  For Each rngRow In importtbl.colinfo.rngRowsCurTbl.Rows
    sFill = CStr(Intersect(rngRow, rngFillValsCol).Value2)
    If Len(sFill) > 0 Then
      sRaw = CStr(Intersect(rngRow, rngVarRawCol).Value2)
      If errs.IsFail(Len(sRaw) = 0, 3) Then GoTo ErrorExit

      Set rngRawHdr = importtbl.tblRaw.rngTblHeaderVal(importtbl.tblRaw, sRaw)
      If errs.IsFail(rngRawHdr Is Nothing, 4, sRaw) Then GoTo ErrorExit
      Set rngDataCol = Intersect(importtbl.tblRaw.rngRows, rngRawHdr.EntireColumn)

      Set dict = ExcelSteps.New_Dictionary
      If Not dict.ParseStringToDictProcedure(sFill) Then GoTo ErrorExit
      aryKeys = dict.GetKeys
      If IsArray(aryKeys) Then
        For Each sKey In aryKeys
          valFill = dict.Item(CStr(sKey))
          For Each rngCell In rngDataCol.Cells
            If UCase$(CStr(sKey)) = "BLANK" Then
              If Len(CStr(rngCell.Value2)) = 0 Then rngCell.Value2 = valFill
            ElseIf CStr(rngCell.Value2) = CStr(sKey) Then
              rngCell.Value2 = valFill
            End If
          Next rngCell
        Next sKey
      End If
    End If
  Next rngRow
  Exit Function

ErrorExit:
  errs.RecordErr "ImportParseNorm.FillMissingVals_prev", FillMissingVals_prev
End Function
'-----------------------------------------------------------------------------------------
' Apply FillVals mappings from colinfo rows to raw data columns using sort/block ops
' JDL 5/7/26
'
Public Function FillMissingVals(importtbl) As Boolean
  SetErrs FillMissingVals: If errs.IsHandle Then On Error GoTo ErrorExit
  Dim rngVarRawCol As Range, rngFillValsCol As Range, rngRow As Range
  Dim rngRawHdr As Range, rngDataCol As Range
  Dim sFill As String, sRaw As String, dict As Object

  If errs.IsFail(importtbl.tblRaw Is Nothing, 1) Then GoTo ErrorExit
  If errs.IsFail(TypeName(importtbl.tblRaw) <> "tblRowsCols", 2) Then GoTo ErrorExit

  Set rngVarRawCol = importtbl.colinfo.tbl.rngTblHeaderVal(importtbl.colinfo.tbl, "VarNameRaw").EntireColumn
  Set rngFillValsCol = importtbl.colinfo.tbl.rngTblHeaderVal(importtbl.colinfo.tbl, "FillVals").EntireColumn

  For Each rngRow In importtbl.colinfo.rngRowsCurTbl.Rows
    sFill = CStr(Intersect(rngRow, rngFillValsCol).Value2)
    If Len(sFill) > 0 Then
      sRaw = CStr(Intersect(rngRow, rngVarRawCol).Value2)
      If errs.IsFail(Len(sRaw) = 0, 3) Then GoTo ErrorExit

      Set rngRawHdr = importtbl.tblRaw.rngTblHeaderVal(importtbl.tblRaw, sRaw)
      If errs.IsFail(rngRawHdr Is Nothing, 4, sRaw) Then GoTo ErrorExit

      'Sort once so equal values are contiguous and can be replaced in blocks
      If Not importtbl.tblRaw.TblSortBy(importtbl.tblRaw, CStr(rngRawHdr.Value2)) Then GoTo ErrorExit
      Set rngDataCol = Intersect(importtbl.tblRaw.rngRows, rngRawHdr.EntireColumn)

      Set dict = ExcelSteps.New_Dictionary
      If Not dict.ParseStringToDictProcedure(sFill) Then GoTo ErrorExit
      If Not importtbl.ApplyFillMapToSortedColumn(importtbl, rngDataCol, dict) Then GoTo ErrorExit
    End If
  Next rngRow
  Exit Function

ErrorExit:
  errs.RecordErr "ImportParseNorm.FillMissingVals", FillMissingVals
End Function
'-----------------------------------------------------------------------------------------
' Apply parsed FillVals dictionary to a sorted data column
' JDL 5/7/26
'
Public Function ApplyFillMapToSortedColumn(importtbl, rngDataCol As Range, dict As Object) As Boolean
  SetErrs ApplyFillMapToSortedColumn: If errs.IsHandle Then On Error GoTo ErrorExit
  Dim aryKeys As Variant, sKey As Variant, valFill As Variant

  If errs.IsFail(rngDataCol Is Nothing, 1) Then GoTo ErrorExit
  If errs.IsFail(dict Is Nothing, 2) Then GoTo ErrorExit

  aryKeys = dict.GetKeys
  If IsArray(aryKeys) Then
    For Each sKey In aryKeys
      valFill = dict.Item(CStr(sKey))
      If Not importtbl.ApplyFillKeyToSortedColumn(importtbl, rngDataCol, CStr(sKey), valFill) Then GoTo ErrorExit
    Next sKey
  End If
  Exit Function

ErrorExit:
  errs.RecordErr "ImportParseNorm.ApplyFillMapToSortedColumn", ApplyFillMapToSortedColumn
End Function
'-----------------------------------------------------------------------------------------
' Apply one FillVals key/value pair to a sorted data column
' JDL 5/7/26
'
Public Function ApplyFillKeyToSortedColumn(importtbl, rngDataCol As Range, ByVal sKey As String, _
  ByVal valFill As Variant) As Boolean
  SetErrs ApplyFillKeyToSortedColumn: If errs.IsHandle Then On Error GoTo ErrorExit
  Dim rngBlanks As Range, rngArea As Range, cellFirst As Range, cellLast As Range

  If UCase$(sKey) = "BLANK" Then
    Set rngBlanks = Nothing
    On Error Resume Next
    Set rngBlanks = rngDataCol.SpecialCells(xlCellTypeBlanks)
    On Error GoTo ErrorExit
    If Not rngBlanks Is Nothing Then
      For Each rngArea In rngBlanks.Areas
        rngArea.Value2 = valFill
      Next rngArea
    End If
  Else
    Set cellFirst = FindInRange(rngDataCol, sKey)
    If Not cellFirst Is Nothing Then
      Set cellLast = rngDataCol.Find(sKey, After:=rngDataCol.Cells(1), _
        LookAt:=xlWhole, SearchDirection:=xlPrevious)
      If cellLast Is Nothing Then Set cellLast = cellFirst
      importtbl.tblRaw.wksht.Range(cellFirst, cellLast).Value2 = valFill
    End If
  End If
  Exit Function

ErrorExit:
  errs.RecordErr "ImportParseNorm.ApplyFillKeyToSortedColumn", ApplyFillKeyToSortedColumn
End Function
'-----------------------------------------------------------------------------------------
' Parse raw data shape into parsed table object
' JDL 5/7/26
'
Public Function ParseRawProcedure(importtbl) As Boolean
  SetErrs ParseRawProcedure: If errs.IsHandle Then On Error GoTo ErrorExit

  If LCase$(GetDictText(importtbl.dParamsParse, "RawShape", "rowscols")) = "rowscols" Then
    Set importtbl.tblParsed = importtbl.tblRaw
  Else
    If errs.IsFail(True, 1, GetDictText(importtbl.dParamsParse, "RawShape", "rowscols")) Then GoTo ErrorExit
  End If
  Exit Function

ErrorExit:
  errs.RecordErr "ImportParseNorm.ParseRawProcedure", ParseRawProcedure
End Function
'-----------------------------------------------------------------------------------------
' Normalize parsed data and apply row filters
' JDL 5/7/26
'
Public Function NormalizeParsedProcedure(importtbl) As Boolean
  SetErrs NormalizeParsedProcedure: If errs.IsHandle Then On Error GoTo ErrorExit

  If Not importtbl.WriteNormalized(importtbl) Then GoTo ErrorExit
  If Not importtbl.FilterRows(importtbl) Then GoTo ErrorExit
  Exit Function

ErrorExit:
  errs.RecordErr "ImportParseNorm.NormalizeParsedProcedure", NormalizeParsedProcedure
End Function
'-----------------------------------------------------------------------------------------
' Build normalized table from ordered colinfo mappings for current table
' JDL 5/7/26
'
Public Function WriteNormalized(importtbl) As Boolean
  SetErrs WriteNormalized: If errs.IsHandle Then On Error GoTo ErrorExit
  Dim maxOrder As Long
  Dim aryNorm() As String, aryRaw() As String

  If errs.IsFail(importtbl.tblParsed Is Nothing, 1) Then GoTo ErrorExit
  If errs.IsFail(TypeName(importtbl.tblParsed) <> "tblRowsCols", 2) Then GoTo ErrorExit

  If Not importtbl.BuildNormMappings(importtbl, aryNorm, aryRaw, maxOrder) Then GoTo ErrorExit
  If Not importtbl.WriteNormFromMappings(importtbl, aryNorm, aryRaw, maxOrder) Then GoTo ErrorExit
  Exit Function

ErrorExit:
  errs.RecordErr "ImportParseNorm.WriteNormalized", WriteNormalized
End Function
'-----------------------------------------------------------------------------------------
' Build ordered VarNameNorm/VarNameRaw arrays from current table metadata
' JDL 5/7/26
'
Public Function BuildNormMappings(importtbl, ByRef aryNorm() As String, ByRef aryRaw() As String, _
  ByRef maxOrder As Long) As Boolean
  SetErrs BuildNormMappings: If errs.IsHandle Then On Error GoTo ErrorExit
  Dim rngOrderCol As Range, rngVarNormCol As Range, rngVarRawCol As Range, rngRow As Range
  Dim orderNum As Long, valOrder As Variant, sNorm As String, sRaw As String

  Set rngOrderCol = importtbl.colinfo.tbl.rngTblHeaderVal(importtbl.colinfo.tbl, importtbl.curTbl).EntireColumn
  Set rngVarNormCol = importtbl.colinfo.tbl.rngTblHeaderVal(importtbl.colinfo.tbl, "VarNameNorm").EntireColumn
  Set rngVarRawCol = importtbl.colinfo.tbl.rngTblHeaderVal(importtbl.colinfo.tbl, "VarNameRaw").EntireColumn

  maxOrder = 0
  For Each rngRow In importtbl.colinfo.rngRowsCurTbl.Rows
    valOrder = Intersect(rngRow, rngOrderCol).Value2
    sNorm = CStr(Intersect(rngRow, rngVarNormCol).Value2)
    If Len(sNorm) > 0 And IsNumeric(valOrder) Then
      orderNum = CLng(valOrder)
      If orderNum > maxOrder Then maxOrder = orderNum
    End If
  Next rngRow
  If errs.IsFail(maxOrder = 0, 1, importtbl.curTbl) Then GoTo ErrorExit

  ReDim aryNorm(1 To maxOrder)
  ReDim aryRaw(1 To maxOrder)

  For Each rngRow In importtbl.colinfo.rngRowsCurTbl.Rows
    valOrder = Intersect(rngRow, rngOrderCol).Value2
    sNorm = CStr(Intersect(rngRow, rngVarNormCol).Value2)
    sRaw = CStr(Intersect(rngRow, rngVarRawCol).Value2)

    If Len(sNorm) > 0 And IsNumeric(valOrder) Then
      orderNum = CLng(valOrder)
      If orderNum > 0 And orderNum <= maxOrder Then
        If errs.IsFail(Len(sRaw) = 0, 2, sNorm) Then GoTo ErrorExit
        aryNorm(orderNum) = sNorm
        aryRaw(orderNum) = sRaw
      End If
    End If
  Next rngRow
  Exit Function

ErrorExit:
  errs.RecordErr "ImportParseNorm.BuildNormMappings", BuildNormMappings
End Function
'-----------------------------------------------------------------------------------------
' Write normalized sheet based on ordered mapping arrays
' JDL 5/7/26
'
Public Function WriteNormFromMappings(importtbl, aryNorm() As String, aryRaw() As String, _
  ByVal maxOrder As Long) As Boolean
  SetErrs WriteNormFromMappings: If errs.IsHandle Then On Error GoTo ErrorExit
  Dim wkshtNorm As Worksheet, normSht As String
  Dim rngSrcHdr As Range, rngSrc As Range, rngDest As Range, tblAlias As Object
  Dim i As Long, j As Long

  normSht = "norm_"
  If SheetExists(importtbl.tblParsed.wkbk, normSht) Then DeleteSheet importtbl.tblParsed.wkbk, normSht
  Set wkshtNorm = importtbl.tblParsed.wkbk.Worksheets.Add(After:=importtbl.tblParsed.wksht)
  wkshtNorm.Name = normSht

  j = 0
  For i = 1 To maxOrder
    If Len(aryNorm(i)) > 0 Then
      j = j + 1
      wkshtNorm.Cells(1, j).Value2 = aryNorm(i)

      Set rngSrcHdr = importtbl.tblParsed.rngTblHeaderVal(importtbl.tblParsed, aryRaw(i))
      If errs.IsFail(rngSrcHdr Is Nothing, 1, aryRaw(i)) Then GoTo ErrorExit

      Set rngSrc = Intersect(importtbl.tblParsed.rngRows, rngSrcHdr.EntireColumn)
      Set rngDest = wkshtNorm.Cells(2, j).Resize(rngSrc.Rows.Count, 1)
      rngDest.Value = rngSrc.Value
    End If
  Next i
  If errs.IsFail(j = 0, 2, importtbl.curTbl) Then GoTo ErrorExit

  Set importtbl.tblNorm = New tblRowsCols
  Set tblAlias = importtbl.tblNorm
  If Not tblAlias.Provision(tblAlias, importtbl.tblParsed.wkbk, False, sht:=normSht, IsSetColRngs:=False) Then GoTo ErrorExit
  Exit Function

ErrorExit:
  errs.RecordErr "ImportParseNorm.WriteNormFromMappings", WriteNormFromMappings
End Function
'-----------------------------------------------------------------------------------------
' Apply row filters described by FilterVals metadata for current table
' JDL 5/7/26
'
Public Function FilterRows(importtbl) As Boolean
  SetErrs FilterRows: If errs.IsHandle Then On Error GoTo ErrorExit
  Dim rngVarNormCol As Range, rngFilterValsCol As Range, rngRow As Range
  Dim sNorm As String, sFilter As String, dict As Object

  If importtbl.tblNorm Is Nothing Then Exit Function

  Set rngVarNormCol = importtbl.colinfo.tbl.rngTblHeaderVal(importtbl.colinfo.tbl, "VarNameNorm").EntireColumn
  Set rngFilterValsCol = importtbl.colinfo.tbl.rngTblHeaderVal(importtbl.colinfo.tbl, "FilterVals").EntireColumn

  For Each rngRow In importtbl.colinfo.rngRowsCurTbl.Rows
    sFilter = CStr(Intersect(rngRow, rngFilterValsCol).Value2)
    If Len(sFilter) > 0 Then
      sNorm = CStr(Intersect(rngRow, rngVarNormCol).Value2)
      If errs.IsFail(Len(sNorm) = 0, 1) Then GoTo ErrorExit

      Set dict = ExcelSteps.New_Dictionary
      If Not dict.ParseStringToDictProcedure(sFilter) Then GoTo ErrorExit

      If dict.Exists("KeepOnly") Then
        If Not ApplyKeepOnlyFilter(importtbl, sNorm, dict.Item("KeepOnly")) Then GoTo ErrorExit
      End If
    End If
  Next rngRow
  Exit Function

ErrorExit:
  errs.RecordErr "ImportParseNorm.FilterRows", FilterRows
End Function
'-----------------------------------------------------------------------------------------
' Keep only rows in tblNorm where sVarNorm column equals keepVal
' JDL 5/7/26
'
Private Function ApplyKeepOnlyFilter(importtbl, ByVal sVarNorm As String, ByVal keepVal As Variant) As Boolean
  SetErrs ApplyKeepOnlyFilter: If errs.IsHandle Then On Error GoTo ErrorExit
  Dim rngHdr As Range, rngDataCol As Range, cellFirst As Range, cellLast As Range
  Dim cellStart As Range, cellEnd As Range, idx As Long, tblAlias As Object

  If Not importtbl.tblNorm.TblSortBy(importtbl.tblNorm, sVarNorm) Then GoTo ErrorExit

  Set rngHdr = importtbl.tblNorm.rngTblHeaderVal(importtbl.tblNorm, sVarNorm)
  If errs.IsFail(rngHdr Is Nothing, 1, sVarNorm) Then GoTo ErrorExit

  Set rngDataCol = Intersect(importtbl.tblNorm.rngRows, rngHdr.EntireColumn)
  Set cellFirst = FindInRange(rngDataCol, keepVal)

  If cellFirst Is Nothing Then
    rngDataCol.EntireRow.Delete
  Else
    For idx = rngDataCol.Cells.Count To 1 Step -1
      If CStr(rngDataCol.Cells(idx, 1).Value2) = CStr(keepVal) Then
        Set cellLast = rngDataCol.Cells(idx, 1)
        Exit For
      End If
    Next idx
    If errs.IsFail(cellLast Is Nothing, 2, CStr(keepVal)) Then GoTo ErrorExit

    Set cellStart = rngDataCol.Cells(1, 1)
    Set cellEnd = rngDataCol.Cells(rngDataCol.Cells.Count, 1)

    If cellLast.Row < cellEnd.Row Then
      importtbl.tblNorm.wksht.Range(cellLast.Offset(1, 0), cellEnd).EntireRow.Delete
    End If
    If cellFirst.Row > cellStart.Row Then
      importtbl.tblNorm.wksht.Range(cellStart, cellFirst.Offset(-1, 0)).EntireRow.Delete
    End If
  End If

  Set tblAlias = importtbl.tblNorm
  If Not tblAlias.Provision(tblAlias, importtbl.tblNorm.wkbk, False, _
    sht:=importtbl.tblNorm.sht, IsSetColRngs:=False) Then GoTo ErrorExit
  Exit Function

ErrorExit:
  errs.RecordErr "ImportParseNorm.ApplyKeepOnlyFilter", ApplyKeepOnlyFilter
End Function
'-----------------------------------------------------------------------------------------
' Return text value from dictionary key if exists, otherwise default value
' JDL 5/7/26
'
Private Function GetDictText(dict As Object, ByVal key As String, Optional ByVal dflt As String = "") As String
  SetErrs "non-bool": If errs.IsHandle Then On Error GoTo ErrorExit

  GetDictText = dflt
  If dict Is Nothing Then Exit Function
  If dict.Exists(key) Then GetDictText = CStr(dict.Item(key))
  Exit Function

ErrorExit:
  errs.RecordErr "ImportParseNorm.GetDictText"
End Function
