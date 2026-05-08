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
'Version 5/8/26

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
' JDL 5/8/26
'
Public Function ImportParseNormProcedure(imptbl, colinfo, files As Object, curTbl As String, _
  dParamsImport As Object, dParamsParse As Object) As Boolean
  SetErrs ImportParseNormProcedure: If errs.IsHandle Then On Error GoTo ErrorExit

  With imptbl
    If Not .Init(imptbl, colinfo, files, curTbl, dParamsImport, dParamsParse) Then GoTo ErrorExit
    If Not .OpenAndValidateRawProcedure(imptbl) Then GoTo ErrorExit
    If Not .ParseRawProcedure(imptbl) Then GoTo ErrorExit
    If Not .NormalizeParsedProcedure(imptbl) Then GoTo ErrorExit
  End With
  Exit Function

ErrorExit:
  errs.RecordErr "ImportParseNormProcedure", ImportParseNormProcedure
End Function
'-----------------------------------------------------------------------------------------
' Initialize attributes and ensure colinfo is set for current table
' JDL 5/7/26
'
Public Function Init(imptbl, colinfo, files As Object, curTbl As String, _
  dParamsImport As Object, dParamsParse As Object) As Boolean
  SetErrs Init: If errs.IsHandle Then On Error GoTo ErrorExit

  'Check that required params are provided
  If errs.IsFail(files Is Nothing, 1) Then GoTo ErrorExit
  If errs.IsFail(Len(files.pfImportFile) = 0, 2) Then GoTo ErrorExit
  If errs.IsFail(dParamsImport Is Nothing, 3) Then GoTo ErrorExit
  If errs.IsFail(dParamsParse Is Nothing, 4) Then GoTo ErrorExit
  If errs.IsFail(Len(curTbl) = 0, 5) Then GoTo ErrorExit

  'Initialize colinfo and set rngRowsCurTbl for curTbl
  If colinfo Is Nothing Then Set colinfo = New ColInfo
  With colinfo
    If .tbl Is Nothing Then
      If Not .Init(colinfo, files) Then GoTo ErrorExit
    End If
    If Not .SetCurTbl(colinfo, curTbl) Then GoTo ErrorExit
  End With

  With imptbl
    Set .colinfo = colinfo
    Set .files = files
    .curTbl = curTbl
    Set .dParamsImport = dParamsImport
    Set .dParamsParse = dParamsParse
    Set .tblRaw = Nothing
    Set .tblParsed = Nothing
    Set .tblNorm = Nothing
  End With
  Exit Function

ErrorExit:
  errs.RecordErr "ImportParseNorm.Init", Init
End Function
'-----------------------------------------------------------------------------------------
' Open raw data and run raw structure checks and fill operations
' JDL 5/7/26
'
Public Function OpenAndValidateRawProcedure(imptbl) As Boolean
  SetErrs OpenAndValidateRawProcedure: If errs.IsHandle Then On Error GoTo ErrorExit

  With imptbl
    If Not .OpenRawData(imptbl) Then GoTo ErrorExit
    If Not .ValidateRawStructure(imptbl) Then GoTo ErrorExit
    If Not .ReplaceVals(imptbl) Then GoTo ErrorExit
  End With
  Exit Function

ErrorExit:
  errs.RecordErr "OpenAndValidateRawProcedure", OpenAndValidateRawProcedure
End Function
'-----------------------------------------------------------------------------------------
' Open raw file and provision raw table object in a temporary workbook copy
' JDL 5/7/26
'
Public Function OpenRawData(imptbl) As Boolean
  SetErrs OpenRawData: If errs.IsHandle Then On Error GoTo ErrorExit
  Dim wkbkOpened As Workbook, wkbkTemp As Workbook
  Dim tblAlias As Object, tblUAlias As Object, shtRaw As String

  If Not OpenFile(imptbl.files.pfImportFile, wkbkOpened) Then GoTo ErrorExit
  If errs.IsFail(wkbkOpened Is Nothing, 1, imptbl.files.pfImportFile) Then GoTo ErrorExit

  wkbkOpened.Worksheets(1).Copy
  Set wkbkTemp = ActiveWorkbook
  wkbkOpened.Close False
  shtRaw = wkbkTemp.Worksheets(1).Name

  With imptbl
    If LCase$(CStr(.dParamsParse.Item("RawShape", "rowscols"))) = "rowscols" Then
      Set .tblRaw = New tblRowsCols
      Set tblAlias = .tblRaw
      If Not tblAlias.Provision(tblAlias, wkbkTemp, False, sht:=shtRaw, IsSetColRngs:=False) Then GoTo ErrorExit
    Else
      Set .tblRaw = New tblUnstructured
      Set tblUAlias = .tblRaw
      If Not tblUAlias.Init(tblUAlias, wkbkTemp, shtRaw) Then GoTo ErrorExit
    End If
  End With
  Exit Function

ErrorExit:
  errs.RecordErr "OpenRawData", OpenRawData
End Function
'-----------------------------------------------------------------------------------------
' Validate that required raw columns are present for current colinfo table
' JDL 5/7/26
'
Public Function ValidateRawStructure(imptbl) As Boolean
  SetErrs ValidateRawStructure: If errs.IsHandle Then On Error GoTo ErrorExit
  Dim rngVarNormCol As Range, rngVarRawCol As Range, rngRow As Range, rngHdr As Range
  Dim valNorm As Variant, valRaw As String

  With imptbl
  If errs.IsFail(imptbl.tblRaw Is Nothing, 1) Then GoTo ErrorExit
  If errs.IsFail(TypeName(imptbl.tblRaw) <> "tblRowsCols", 2) Then GoTo ErrorExit

  Set rngVarNormCol = imptbl.colinfo.tbl.rngTblHeaderVal(imptbl.colinfo.tbl, "VarNameNorm").EntireColumn
  Set rngVarRawCol = imptbl.colinfo.tbl.rngTblHeaderVal(imptbl.colinfo.tbl, "VarNameRaw").EntireColumn

  For Each rngRow In imptbl.colinfo.rngRowsCurTbl.Rows
    valNorm = Intersect(rngRow, rngVarNormCol).Value2
    If Len(CStr(valNorm)) > 0 Then
      valRaw = CStr(Intersect(rngRow, rngVarRawCol).Value2)
      If errs.IsFail(Len(valRaw) = 0, 3, CStr(valNorm)) Then GoTo ErrorExit

      Set rngHdr = imptbl.tblRaw.rngTblHeaderVal(imptbl.tblRaw, valRaw)
      If errs.IsFail(rngHdr Is Nothing, 4, valRaw) Then GoTo ErrorExit
    End If
  Next rngRow
  End With
  Exit Function

ErrorExit:
  errs.RecordErr "ValidateRawStructure", ValidateRawStructure
End Function
'-----------------------------------------------------------------------------------------
' Apply FillVals mappings from colinfo rows to raw data columns using sort/block ops
' JDL 5/7/26
'
Public Function ReplaceVals(imptbl) As Boolean
  SetErrs ReplaceVals: If errs.IsHandle Then On Error GoTo ErrorExit
  Dim rngRow As Range
  Dim rngRawHdr As Range, rngDataCol As Range
  Dim sFill As String, sRaw As String, dict As Object


  With imptbl
    For Each rngRow In .colinfo.rngRowsCurTbl.Rows

      'Read fill/replace vals JSON string from colinfo
      sFill = CStr(Intersect(rngRow, .colinfo.tbl.colrngFillVals).Value2)
      If Len(sFill) > 0 Then

        ' Set sRaw as colinfo row's raw data variable name
        sRaw = CStr(Intersect(rngRow, .colinfo.tbl.colrngVarRaw).Value2)
        If errs.IsFail(Len(sRaw) = 0, 3) Then GoTo ErrorExit

        ' Locate sRaw column in raw table
        Set rngRawHdr = .tblRaw.rngTblHeaderVal(.tblRaw, sRaw)
        If errs.IsFail(rngRawHdr Is Nothing, 4, sRaw) Then GoTo ErrorExit

        'Sort so equal values are contiguous and can be replaced in blocks
        If Not .tblRaw.TblSortBy(.tblRaw, CStr(rngRawHdr.Value2)) Then GoTo ErrorExit
        Set rngDataCol = Intersect(.tblRaw.rngRows, rngRawHdr.EntireColumn)

        Set dict = New Dictionary
        If Not dict.ParseStringToDictProcedure(sFill) Then GoTo ErrorExit
        If Not .ApplyFillMapToSortedColumn(imptbl, rngDataCol, dict) Then GoTo ErrorExit
      End If
    Next rngRow
  End With
  Exit Function

ErrorExit:
  errs.RecordErr "ReplaceVals", ReplaceVals
End Function
'-----------------------------------------------------------------------------------------
' Apply parsed FillVals dictionary to a sorted data column
' JDL 5/7/26
'
Public Function ApplyFillMapToSortedColumn(imptbl, rngDataCol As Range, dict As Object) As Boolean
  SetErrs ApplyFillMapToSortedColumn: If errs.IsHandle Then On Error GoTo ErrorExit
  Dim aryKeys As Variant, sKey As Variant, valFill As Variant

  If errs.IsFail(rngDataCol Is Nothing, 1) Then GoTo ErrorExit
  If errs.IsFail(dict Is Nothing, 2) Then GoTo ErrorExit

  aryKeys = dict.GetKeys
  If IsArray(aryKeys) Then
    For Each sKey In aryKeys
      valFill = dict.Item(CStr(sKey))
      Debug.Print valFill, sKey
      If Not imptbl.ApplyFillKeyToSortedColumn(imptbl, rngDataCol, CStr(sKey), valFill) Then GoTo ErrorExit
    Next sKey
  End If
  Exit Function

ErrorExit:
  errs.RecordErr "ApplyFillMapToSortedColumn", ApplyFillMapToSortedColumn
End Function
'-----------------------------------------------------------------------------------------
' Apply one FillVals key/value pair to a sorted data column
' JDL 5/7/26
'
Public Function ApplyFillKeyToSortedColumn(imptbl, rngDataCol As Range, ByVal sKey As String, _
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
      imptbl.tblRaw.wksht.Range(cellFirst, cellLast).Value2 = valFill
    End If
  End If
  Exit Function

ErrorExit:
  errs.RecordErr "ApplyFillKeyToSortedColumn", ApplyFillKeyToSortedColumn
End Function
'-----------------------------------------------------------------------------------------
' Parse raw data shape into parsed table object
' JDL 5/7/26
'
Public Function ParseRawProcedure(imptbl) As Boolean
  SetErrs ParseRawProcedure: If errs.IsHandle Then On Error GoTo ErrorExit

  If LCase$(CStr(imptbl.dParamsParse.Item("RawShape", "rowscols"))) = "rowscols" Then
    Set imptbl.tblParsed = imptbl.tblRaw
  Else
    If errs.IsFail(True, 1, CStr(imptbl.dParamsParse.Item("RawShape", "rowscols"))) Then GoTo ErrorExit
  End If
  Exit Function

ErrorExit:
  errs.RecordErr "ParseRawProcedure", ParseRawProcedure
End Function
'-----------------------------------------------------------------------------------------
' Normalize parsed data and apply row filters
' JDL 5/7/26
'
Public Function NormalizeParsedProcedure(imptbl) As Boolean
  SetErrs NormalizeParsedProcedure: If errs.IsHandle Then On Error GoTo ErrorExit

  If Not imptbl.WriteNormalized(imptbl) Then GoTo ErrorExit
  If Not imptbl.FilterRows(imptbl) Then GoTo ErrorExit
  Exit Function

ErrorExit:
  errs.RecordErr "NormalizeParsedProcedure", NormalizeParsedProcedure
End Function
'-----------------------------------------------------------------------------------------
' Build normalized table from ordered colinfo mappings for current table
' JDL 5/7/26
'
Public Function WriteNormalized(imptbl) As Boolean
  SetErrs WriteNormalized: If errs.IsHandle Then On Error GoTo ErrorExit
  Dim maxOrder As Long
  Dim aryNorm() As String, aryRaw() As String

  If errs.IsFail(imptbl.tblParsed Is Nothing, 1) Then GoTo ErrorExit
  If errs.IsFail(TypeName(imptbl.tblParsed) <> "tblRowsCols", 2) Then GoTo ErrorExit

  If Not imptbl.BuildNormMappings(imptbl, aryNorm, aryRaw, maxOrder) Then GoTo ErrorExit
  If Not imptbl.WriteNormFromMappings(imptbl, aryNorm, aryRaw, maxOrder) Then GoTo ErrorExit
  Exit Function

ErrorExit:
  errs.RecordErr "WriteNormalized", WriteNormalized
End Function
'-----------------------------------------------------------------------------------------
' Build ordered VarNameNorm/VarNameRaw arrays from current table metadata
' JDL 5/7/26
'
Public Function BuildNormMappings(imptbl, ByRef aryNorm() As String, ByRef aryRaw() As String, _
  ByRef maxOrder As Long) As Boolean
  SetErrs BuildNormMappings: If errs.IsHandle Then On Error GoTo ErrorExit
  Dim rngOrderCol As Range, rngVarNormCol As Range, rngVarRawCol As Range, rngRow As Range
  Dim orderNum As Long, valOrder As Variant, sNorm As String, sRaw As String

  Set rngOrderCol = imptbl.colinfo.tbl.rngTblHeaderVal(imptbl.colinfo.tbl, imptbl.curTbl).EntireColumn
  Set rngVarNormCol = imptbl.colinfo.tbl.rngTblHeaderVal(imptbl.colinfo.tbl, "VarNameNorm").EntireColumn
  Set rngVarRawCol = imptbl.colinfo.tbl.rngTblHeaderVal(imptbl.colinfo.tbl, "VarNameRaw").EntireColumn

  maxOrder = 0
  For Each rngRow In imptbl.colinfo.rngRowsCurTbl.Rows
    valOrder = Intersect(rngRow, rngOrderCol).Value2
    sNorm = CStr(Intersect(rngRow, rngVarNormCol).Value2)
    If Len(sNorm) > 0 And IsNumeric(valOrder) Then
      orderNum = CLng(valOrder)
      If orderNum > maxOrder Then maxOrder = orderNum
    End If
  Next rngRow
  If errs.IsFail(maxOrder = 0, 1, imptbl.curTbl) Then GoTo ErrorExit

  ReDim aryNorm(1 To maxOrder)
  ReDim aryRaw(1 To maxOrder)

  For Each rngRow In imptbl.colinfo.rngRowsCurTbl.Rows
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
  errs.RecordErr "BuildNormMappings", BuildNormMappings
End Function
'-----------------------------------------------------------------------------------------
' Write normalized sheet based on ordered mapping arrays
' JDL 5/7/26; Updated 5/8/26
'
Public Function WriteNormFromMappings(imptbl, aryNorm() As String, aryRaw() As String, _
  ByVal maxOrder As Long) As Boolean
  SetErrs WriteNormFromMappings: If errs.IsHandle Then On Error GoTo ErrorExit
  Dim wkshtNorm As Worksheet, normSht As String
  Dim rngSrcHdr As Range, rngSrc As Range, rngDest As Range, tblAlias As Object
  Dim i As Long, j As Long

  normSht = "norm_"
  If SheetExists(imptbl.tblParsed.wkbk, normSht) Then DeleteSheet imptbl.tblParsed.wkbk, normSht
  Set wkshtNorm = imptbl.tblParsed.wkbk.Worksheets.Add(After:=imptbl.tblParsed.wksht)
  wkshtNorm.Name = normSht

  j = 0
  For i = 1 To maxOrder
    If Len(aryNorm(i)) > 0 Then
      j = j + 1
      wkshtNorm.Cells(1, j).Value2 = aryNorm(i)

      Set rngSrcHdr = imptbl.tblParsed.rngTblHeaderVal(imptbl.tblParsed, aryRaw(i))
      If errs.IsFail(rngSrcHdr Is Nothing, 1, aryRaw(i)) Then GoTo ErrorExit

      Set rngSrc = Intersect(imptbl.tblParsed.rngRows, rngSrcHdr.EntireColumn)
      Set rngDest = wkshtNorm.Cells(2, j).Resize(rngSrc.Rows.Count, 1)
      rngDest.Value = rngSrc.Value
    End If
  Next i
  If errs.IsFail(j = 0, 2, imptbl.curTbl) Then GoTo ErrorExit

  Set imptbl.tblNorm = New tblRowsCols
  Set tblAlias = imptbl.tblNorm
  If Not tblAlias.Provision(tblAlias, imptbl.tblParsed.wkbk, False, sht:=normSht, _
    IsSetColRngs:=False) Then GoTo ErrorExit
  Exit Function

ErrorExit:
  errs.RecordErr "WriteNormFromMappings", WriteNormFromMappings
End Function
'-----------------------------------------------------------------------------------------
' Apply row filters described by FilterVals metadata for current table
' JDL 5/7/26
'
Public Function FilterRows(imptbl) As Boolean
  SetErrs FilterRows: If errs.IsHandle Then On Error GoTo ErrorExit
  Dim rngVarNormCol As Range, rngFilterValsCol As Range, rngRow As Range
  Dim sNorm As String, sFilter As String, dict As Object

  If imptbl.tblNorm Is Nothing Then Exit Function

  Set rngVarNormCol = imptbl.colinfo.tbl.rngTblHeaderVal(imptbl.colinfo.tbl, "VarNameNorm").EntireColumn
  Set rngFilterValsCol = imptbl.colinfo.tbl.rngTblHeaderVal(imptbl.colinfo.tbl, "FilterVals").EntireColumn

  For Each rngRow In imptbl.colinfo.rngRowsCurTbl.Rows
    sFilter = CStr(Intersect(rngRow, rngFilterValsCol).Value2)
    If Len(sFilter) > 0 Then
      sNorm = CStr(Intersect(rngRow, rngVarNormCol).Value2)
      If errs.IsFail(Len(sNorm) = 0, 1) Then GoTo ErrorExit

      Set dict = New Dictionary
      If Not dict.ParseStringToDictProcedure(sFilter) Then GoTo ErrorExit

      If dict.Exists("KeepOnly") Then
        If Not ApplyKeepOnlyFilter(imptbl, sNorm, dict.Item("KeepOnly")) Then GoTo ErrorExit
      End If
    End If
  Next rngRow
  Exit Function

ErrorExit:
  errs.RecordErr "FilterRows", FilterRows
End Function
'-----------------------------------------------------------------------------------------
' Keep only rows in tblNorm where sVarNorm column equals keepVal
' JDL 5/7/26
'
Private Function ApplyKeepOnlyFilter(imptbl, ByVal sVarNorm As String, ByVal keepVal As Variant) As Boolean
  SetErrs ApplyKeepOnlyFilter: If errs.IsHandle Then On Error GoTo ErrorExit
  Dim rngHdr As Range, rngDataCol As Range, cellFirst As Range, cellLast As Range
  Dim cellStart As Range, cellEnd As Range, idx As Long, tblAlias As Object

  If Not imptbl.tblNorm.TblSortBy(imptbl.tblNorm, sVarNorm) Then GoTo ErrorExit

  Set rngHdr = imptbl.tblNorm.rngTblHeaderVal(imptbl.tblNorm, sVarNorm)
  If errs.IsFail(rngHdr Is Nothing, 1, sVarNorm) Then GoTo ErrorExit

  Set rngDataCol = Intersect(imptbl.tblNorm.rngRows, rngHdr.EntireColumn)
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
      imptbl.tblNorm.wksht.Range(cellLast.Offset(1, 0), cellEnd).EntireRow.Delete
    End If
    If cellFirst.Row > cellStart.Row Then
      imptbl.tblNorm.wksht.Range(cellStart, cellFirst.Offset(-1, 0)).EntireRow.Delete
    End If
  End If

  Set tblAlias = imptbl.tblNorm
  If Not tblAlias.Provision(tblAlias, imptbl.tblNorm.wkbk, False, _
    sht:=imptbl.tblNorm.sht, IsSetColRngs:=False) Then GoTo ErrorExit
  Exit Function

ErrorExit:
  errs.RecordErr "ApplyKeepOnlyFilter", ApplyKeepOnlyFilter
End Function