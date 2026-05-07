VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "BRImport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Dashboard_BRImport_cls.vb
'Version 4/30/26
'Orchestrates automatic BR history import: open/normalize source, validate,
'build weekly and monthly pivot matrices, prepare and transfer to WklyHist and OTB,
'then refresh Monthly date rows.
Option Explicit

' Source file context
Public pathData As String               ' Directory path of source file
Public fBRRaw As String                 ' Source file name

' Source and normalized tables (.Init only; .Provision not called)
Public tblRawBR As Object               ' tblRowsCols for opened source workbook sheet
Public tblNormBR As Object              ' tblRowsCols for location-filtered normalized rows

' Pivot output tables (.Init only; pivot structure does not conform to tblRowsCols)
Public rngPivotMo As Range              ' Monthly aggregate matrix (metric rows, mo_year cols)

' PivotTable class instance for BuildWeeklyPivotProcedure
Public pvt As Object

' Location and metric mapping
Public locationBR As String             ' location_client_BR value from Client_Settings
Public curMoYr As String                ' month_year_current value from Client_Settings
Public curMo As String                  ' month_current value from Client_Settings
Public curYr As Long                    ' year_wkStart value from Client_Settings
Public dictMetrics As Object            ' raw column name -> canonical metric name
Public nPT As Long                      ' unique cleaned prodtype count (set by CleanupProdtypesProcedure)
Public colNameLocn As String            ' Location key field column name in source (from keyFieldsBR)
Public colNamePT As String              ' ProductType key field column name in source (from keyFieldsBR)
Public colNameWk As String              ' Week key field column name in source (from keyFieldsBR)
Public colNameYr As String              ' Year key field column name in source (from keyFieldsBR)

' tblNormBR key column ranges (set in OpenAndNormalizeSourceProcedure)
Public colrngNormWkSerial As Range      ' wk_serial_hist column in tblNormBR
Public colrngNormYear As Range          ' year_hist column in tblNormBR
Public colrngNormProdtype As Range      ' prodtype column in tblNormBR (cleaned by CleanupProdtypes)

' Transfer target ("WklyHist" or "OTB" - set by Prep methods before TransferPreparedRanges)
Public transferTarget As String

' Transfer range attributes - WklyHist
Public rngSrcDates As Range             ' wk_serial_hist header row of sorted tblPivotWk columns
Public rngSrcYears As Range             ' year_hist header row of sorted tblPivotWk columns
Public rngSrcMetricData As Range        ' sales_hist metric block rows for source columns
Public rngFirstScenWk As Range          ' First WklyHist Scenario row cell at colrngFirstScenario
Public nCurMoWks As Long                ' Count of current-month buffer cols in WklyHist
Public rngDestDates As Range            ' wk_serial_hist variable row in mdls.WklyHist
Public rngDestYears As Range            ' year_hist variable row in mdls.WklyHist
Public rngDestMetricData As Range       ' sales_hist_prodtype_x variable rows in mdls.WklyHist

' Transfer range attributes - OTB
Public rngSrcDiscounts As Range         ' Discounts metric block rows for resolved source cols
Public rngSrcMarkdowns As Range         ' Markdowns metric block rows for resolved source cols
Public rngSrcCOGS As Range             ' COGS metric block rows for resolved source cols
Public rngDestDiscounts As Range        ' Discounts_prodtype_x variable rows in mdls.OTB
Public rngDestMarkdowns As Range        ' Markdowns_prodtype_x variable rows in mdls.OTB
Public rngDestCOGS As Range            ' COGS_prodtype_x variable rows in mdls.OTB

'-------------------------------------------------------------------------------------
'-------------------------------------------------------------------------------------
' ImportBRMetricsProcedure - orchestrate full BR automatic import
'-------------------------------------------------------------------------------------
'-------------------------------------------------------------------------------------
' Orchestrate full import: open/normalize/validate/pivot, write WklyHist and OTB
' JDL 04/02/26
'
Public Function ImportBRMetricsProcedure(import As BRImport, mdls, tbls) As Boolean
    setErrs ImportBRMetricsProcedure, ThisWorkbook
    If errs.IsHandle Then On Error GoTo ErrorExit

    With import
        If Not .InitImportContext(import, mdls, tbls) Then GoTo ErrorExit

        ' Open source, filter to location, normalize headers, clean and validate raw data
        If Not .OpenAndNormalizeSourceProcedure(import) Then GoTo ErrorExit
        If Not .ValidateRawDataProcedure(import, tbls) Then GoTo ErrorExit
        If Not .CleanupProdtypesProcedure(import, mdls) Then GoTo ErrorExit

        ' <Other code not included in this snippet>

    End With
    Exit Function

ErrorExit:
    errs.RecordErr "ImportBRMetricsProcedure", ImportBRMetricsProcedure
End Function
'-------------------------------------------------------------------------------------
'-------------------------------------------------------------------------------------
' InitImportContext - initialize models/tables and resolve source location/path
'-------------------------------------------------------------------------------------
'-------------------------------------------------------------------------------------
' Initialize required models, Calendar table, location, and metric dictionary
' JDL 04/02/26
'
Public Function InitImportContext(import As BRImport, mdls, tbls) As Boolean
    setErrs InitImportContext, ThisWorkbook
    If errs.IsHandle Then On Error GoTo ErrorExit
    Dim rawNames As Variant, canonNames As Variant, keyFields As Variant, i As Long

    ' Initialize required mdls and tbls.Calendar
    If Not InitMdls(mdls, IsAll:=True, IsRefresh:=False) Then GoTo ErrorExit
    If Not InitTbls(tbls, IsAll:=False, IsCalendar:=True) Then GoTo ErrorExit

    With import
        ' Read and validate Client_Settings inputs
        If Not .ReadClientSettingsInputs(import) Then GoTo ErrorExit

        ' Validate pathData is specified
        If errs.IsFail(Len(.pathData) < 2, 1) Then GoTo ErrorExit

        ' Build dictionary mapping raw metric column names to canonical names
        Set .dictMetrics = ExcelSteps.New_Dictionary

        'Like Python Zip. Add to Dictionary.cls?
        rawNames = Split(sBRMetricRawNames, ",")
        If IsTest Then rawNames = Split(sBRMetricRawNames_Test, ",")
        canonNames = Split(sBRMetricCanonNames, ",")
        For i = 0 To UBound(rawNames)
            .dictMetrics.Add rawNames(i), canonNames(i)
        Next i

        ' Set key field column name attributes for test/production toggle
        keyFields = Split(keyFieldsBR, ",")
        If IsTest Then keyFields = Split(keyFieldsBR_Test, ",")
        .colNameLocn = keyFields(0)
        .colNamePT = keyFields(1)
        .colNameWk = keyFields(2)
        .colNameYr = keyFields(3)
    End With
    Exit Function

ErrorExit:
    errs.RecordErr "InitImportContext", InitImportContext
End Function
'--------------------------------------------------------------------------------------
' Initialize Client_Settings model; read and validate location and current-period attrs
' JDL 04/14/26; updated 4/20/26
'
Public Function ReadClientSettingsInputs(import As BRImport) As Boolean
    setErrs ReadClientSettingsInputs, ThisWorkbook
    If errs.IsHandle Then On Error GoTo ErrorExit
    Dim mdls As Object, curDate As Variant

    If Not InitMdls(mdls, IsAll:=False, IsClientSet:=True, IsRefresh:=False) Then GoTo ErrorExit
    With mdls.Client_Settings
        import.locationBR = .ScenModelLoc(mdls.Client_Settings, "location_client_BR").Value2
        If errs.IsFail(Len(import.locationBR) < 1, 1) Then GoTo ErrorExit

        curDate = .ScenModelLoc(mdls.Client_Settings, "date_Update").Value2
        If errs.IsFail(Len(curDate) < 1, 2) Then GoTo ErrorExit

        import.curMoYr = .ScenModelLoc(mdls.Client_Settings, "month_year_current").Value2
        import.curMo = .ScenModelLoc(mdls.Client_Settings, "month_current").Value2
        import.curYr = .ScenModelLoc(mdls.Client_Settings, "year_wkStart").Value2
        curDate = .ScenModelLoc(mdls.Client_Settings, "year_wkStart").Value2
    End With
    Exit Function

ErrorExit:
    errs.RecordErr "ReadClientSettingsInputs", ReadClientSettingsInputs
End Function
'-------------------------------------------------------------------------------------
'-------------------------------------------------------------------------------------
' OpenAndNormalizeSourceProcedure - open source and write filtered/mapped rows
'-------------------------------------------------------------------------------------
'-------------------------------------------------------------------------------------
' Orchestrate open, header validation, and location-filtered row normalization
' JDL 04/03/26
'
Public Function OpenAndNormalizeSourceProcedure(import As BRImport) As Boolean
    setErrs OpenAndNormalizeSourceProcedure, ThisWorkbook
    If errs.IsHandle Then On Error GoTo ErrorExit
    Dim dictColIdx As Object

    With import
        If Not .OpenSourceWorkbook(import) Then GoTo ErrorExit
        If Not .ValidateSourceStructure(import, dictColIdx) Then GoTo ErrorExit
        If Not .FillMissingProdtypes(import, dictColIdx) Then GoTo ErrorExit
        If Not .WriteNormalizedRows(import, dictColIdx) Then GoTo ErrorExit
    End With
    Exit Function

ErrorExit:
    errs.RecordErr "OpenAndNormalizeSourceProcedure", OpenAndNormalizeSourceProcedure
End Function

'--------------------------------------------------------------------------------------
' Open BR source workbook and provision tblRawBR to populate nCols, nRows, rngHeader
' JDL 04/03/26
'
Public Function OpenSourceWorkbook(import As BRImport) As Boolean
    setErrs OpenSourceWorkbook, ThisWorkbook
    If errs.IsHandle Then On Error GoTo ErrorExit
    Dim wkbkOpened As Workbook, tblRaw As Object, sht As String

    With import
    ' Instance tblRawBR and open source workbook
        Set .tblRawBR = ExcelSteps.New_tbl
        If Not ExcelSteps.OpenFile(.pathData & .fBRRaw, wkbkOpened) Then GoTo ErrorExit

        ' Provision against first sheet; populates nCols, nRows, rngHeader, rngRows
        Set tblRaw = .tblRawBR
        sht = wkbkOpened.Sheets(1).Name
        If Not tblRaw.Provision(tblRaw, wkbkOpened, False, sht:=sht) Then GoTo ErrorExit
    End With
    Exit Function

ErrorExit:
    errs.RecordErr "OpenSourceWorkbook", OpenSourceWorkbook
End Function

'--------------------------------------------------------------------------------------
' Build column index dictionary from tblRawBR header; validate key and metric columns
' JDL 04/03/26
'
Public Function ValidateSourceStructure(import As BRImport, dictColIdx As Object) As Boolean
  setErrs ValidateSourceStructure, ThisWorkbook
  If errs.IsHandle Then On Error GoTo ErrorExit
  Dim rngCell As Range, sHdr As String, keyField As Variant, keyFields As Variant
  Dim metricKeys() As String, metricKey As Variant, metricVals() As Variant
  Dim validSet As Object, i As Long

  ' Build column name -> column index dictionary from tblRawBR header
  Set dictColIdx = ExcelSteps.New_Dictionary
  For Each rngCell In import.tblRawBR.rngHeader.Cells
      sHdr = rngCell.Value2
      If Len(sHdr) > 0 Then dictColIdx.Add sHdr, rngCell.Column
  Next rngCell

  ' Confirm required key fields present in source header
  keyFields = Split(keyFieldsBR, ",")
  If IsTest Then keyFields = Split(keyFieldsBR_Test, ",")
  For Each keyField In keyFields
      If errs.IsFail(Not dictColIdx.Exists(keyField), 1, keyField) Then GoTo ErrorExit
  Next keyField

  ' Confirm all dictMetrics raw column names present in source header
  metricKeys = import.dictMetrics.GetKeys()
  For i = 0 To UBound(metricKeys)
      metricKey = metricKeys(i)
      If errs.IsFail(Not dictColIdx.Exists(CStr(metricKey)), 2, CStr(metricKey)) Then _
          GoTo ErrorExit
  Next i
  Exit Function

ErrorExit:
    errs.RecordErr "ValidateSourceStructure", ValidateSourceStructure
End Function
'--------------------------------------------------------------------------------------
' Create BRNorm sheet, write canonical header, filter rows by location, set tblNormBR
' JDL 04/03/26
'
Public Function WriteNormalizedRows(import As BRImport, dictColIdx As Object) As Boolean
    setErrs WriteNormalizedRows, ThisWorkbook
    If errs.IsHandle Then On Error GoTo ErrorExit
    Dim normSht As String, wkshtNorm As Worksheet
    Dim colIdxLocn As Long, colIdxPT As Long, colIdxWk As Long, colIdxYr As Long
    Dim metricKeys() As String, hdrArr As Variant, nMetrics As Long, i As Long
    Dim destRow As Long, rngRow As Range, locVal As String
    Dim wkVal As Long, yrVal As Long, ptVal As String, tblNorm As Object

    With import
        ' Create fresh BRNorm sheet in source workbook
        normSht = "BRNorm"
        If SheetExists(.tblRawBR.wkbk, normSht) Then DeleteSheet .tblRawBR.wkbk, normSht
        Set wkshtNorm = .tblRawBR.wkbk.Worksheets.Add(After:=.tblRawBR.wksht)
        wkshtNorm.Name = normSht

        ' Build canonical header: fixed key cols then canonical metric names from dictMetrics
        metricKeys = .dictMetrics.GetKeys()
        nMetrics = UBound(metricKeys) + 1
        hdrArr = Array("wk_serial_hist", "year_hist", "prodtype")
        ReDim Preserve hdrArr(2 + nMetrics)
        For i = 0 To UBound(metricKeys)
            hdrArr(3 + i) = .dictMetrics.Item(metricKeys(i))
        Next i
        wkshtNorm.Range("A1").Resize(1, UBound(hdrArr) + 1).Value = hdrArr

        ' Resolve source column indices for key fields from class attributes
        colIdxLocn = dictColIdx.Item(.colNameLocn)
        colIdxPT = dictColIdx.Item(.colNamePT)
        colIdxWk = dictColIdx.Item(.colNameWk)
        colIdxYr = dictColIdx.Item(.colNameYr)

        ' Iterate source rows; write location-matching rows with valid data to BRNorm
        destRow = 2
        Set rngRow = .tblRawBR.rngrows.Rows(1)
        Do While Not Intersect(rngRow, .tblRawBR.rngrows) Is Nothing
            locVal = Trim(CStr(.tblRawBR.wksht.Cells(rngRow.Row, colIdxLocn).Value2))
            If locVal = .locationBR Then
                wkVal = CLng(.tblRawBR.wksht.Cells(rngRow.Row, colIdxWk).Value2)
                yrVal = CLng(.tblRawBR.wksht.Cells(rngRow.Row, colIdxYr).Value2)
                ptVal = Trim(CStr(.tblRawBR.wksht.Cells(rngRow.Row, colIdxPT).Value2))
                If Len(ptVal) > 0 And wkVal > 0 And yrVal > 0 Then
                    wkshtNorm.Cells(destRow, 1).Value2 = wkVal
                    wkshtNorm.Cells(destRow, 2).Value2 = yrVal
                    wkshtNorm.Cells(destRow, 3).Value2 = ptVal
                    For i = 0 To UBound(metricKeys)
                        wkshtNorm.Cells(destRow, 4 + i).Value2 = _
                            .tblRawBR.wksht.Cells(rngRow.Row, _
                                CLng(dictColIdx.Item(metricKeys(i)))).Value2
                    Next i
                    destRow = destRow + 1
                End If
            End If
            Set rngRow = rngRow.Offset(1, 0)
        Loop

        ' Provision tblNormBR against completed BRNorm sheet
        Set .tblNormBR = ExcelSteps.New_tbl
        Set tblNorm = .tblNormBR
        If Not tblNorm.Provision(tblNorm, .tblRawBR.wkbk, False, sht:=normSht) Then GoTo ErrorExit

        ' Set key column ranges for downstream iteration
        Set .colrngNormWkSerial = wkshtNorm.Columns(1)
        Set .colrngNormYear = wkshtNorm.Columns(2)
        Set .colrngNormProdtype = wkshtNorm.Columns(3)
    End With
    Exit Function

ErrorExit:
    errs.RecordErr "WriteNormalizedRows", WriteNormalizedRows
End Function
'-------------------------------------------------------------------------------------
'-------------------------------------------------------------------------------------
' FillMissingProdtypes - replace blank ProductType cells in tblRawBR with UNKNOWN
'-------------------------------------------------------------------------------------
'-------------------------------------------------------------------------------------
' Replace blank ProductType cells in raw source data with "UNKNOWN" before normalization
' JDL 04/07/26
'
Public Function FillMissingProdtypes(import As BRImport, dictColIdx As Object) As Boolean
    setErrs FillMissingProdtypes, ThisWorkbook
    If errs.IsHandle Then On Error GoTo ErrorExit
    Dim colIdxPT As Long, rngRow As Range, cellPT As Range

    colIdxPT = dictColIdx.Item(import.colNamePT)
    With import
        Set rngRow = .tblRawBR.rngrows.Rows(1)
        Do While Not Intersect(rngRow, .tblRawBR.rngrows) Is Nothing
            Set cellPT = .tblRawBR.wksht.Cells(rngRow.Row, colIdxPT)
            If Len(Trim(CStr(cellPT.Value2))) = 0 Then cellPT.Value2 = "UNKNOWN"
            Set rngRow = rngRow.Offset(1, 0)
        Loop
    End With
    Exit Function

ErrorExit:
    errs.RecordErr "FillMissingProdtypes", FillMissingProdtypes
End Function
'-------------------------------------------------------------------------------------
'-------------------------------------------------------------------------------------
' CleanupProdtypesProcedure - apply ProductTypes replacement mapping to prodtype col
'-------------------------------------------------------------------------------------
'-------------------------------------------------------------------------------------
' Apply ProductTypes lookup replacement/standardization mapping to tblNormBR prodtype col
' JDL 04/02/26
'
Public Function CleanupProdtypesProcedure(import As BRImport, mdls) As Boolean
    setErrs CleanupProdtypesProcedure, ThisWorkbook
    If errs.IsHandle Then On Error GoTo ErrorExit
    Dim tblPT As Object, wkbkPT As Workbook
    Dim fPT As String, colOrig As Range, colReplace As Range, colIgnore As Range
    Dim rngRow As Range, rawPT As String, rngFound As Range, replacePT As String
    Dim ignoreVal As Variant, finalPT As String, dictCleanedProdtypes As Object

    ' Step 1: Load ProductTypes lookup from import.pathData
    Set tblPT = ExcelSteps.New_tbl
    Set dictCleanedProdtypes = ExcelSteps.New_Dictionary
    fPT = mdls.params.ScenModelLoc(mdls.params, "file_prodtypes").Value2
    If errs.IsFail(Len(fPT) < 1, 1, "file_prodtypes") Then GoTo ErrorExit

    With tblPT
        If Not ExcelSteps.OpenFile(import.pathData & fPT, wkbkPT) Then GoTo ErrorExit
        If Not .Provision(tblPT, wkbkPT, False, sht:="ProductTypes") Then GoTo ErrorExit
        Set colOrig = .rngTblHeaderVal(tblPT, "ProductType_orig").EntireColumn
        Set colReplace = .rngTblHeaderVal(tblPT, "Name_replace").EntireColumn
        Set colIgnore = .rngTblHeaderVal(tblPT, "Ignore").EntireColumn
    End With

    ' Step 2: Apply replacement map to prodtype column in tblNormBR.rngRows
    With import
        Set rngRow = .tblNormBR.rngrows.Rows(1)
        Do While Not Intersect(rngRow, .tblNormBR.rngrows) Is Nothing
            rawPT = Trim(CStr(Intersect(rngRow, .colrngNormProdtype).Value2))
            If Len(rawPT) > 0 And rawPT <> "IGNORE" Then
            
                ' Look up rawPT in ProductType_orig column
                Set rngFound = ExcelSteps.FindInRange(Intersect(colOrig, tblPT.rngrows), rawPT)
                If Not rngFound Is Nothing Then
                    ignoreVal = Intersect(rngFound.EntireRow, colIgnore).Value2
                    If ignoreVal = True Then
                        .colrngNormProdtype.Replace What:=rawPT, Replacement:="IGNORE", LookAt:=xlWhole
                    Else
                        replacePT = Trim(CStr(Intersect(rngFound.EntireRow, colReplace).Value2))
                        If Len(replacePT) > 0 Then
                            Intersect(rngRow, .colrngNormProdtype).Value2 = replacePT
                        End If
                    End If
                Else
                    If errs.IsFail(True, 2, rawPT) Then GoTo ErrorExit
                End If
            End If
            finalPT = Trim(CStr(Intersect(rngRow, .colrngNormProdtype).Value2))
            If Len(finalPT) > 0 And finalPT <> "IGNORE" Then
                If Not dictCleanedProdtypes.Exists(LCase(finalPT)) Then dictCleanedProdtypes.Add LCase(finalPT), True
            End If
            Set rngRow = rngRow.Offset(1, 0)
        Loop
    End With
    import.nPT = dictCleanedProdtypes.Size

    ' Step 3: Close ProductTypes workbook
    wkbkPT.Close False
    Exit Function

ErrorExit:
    If Not wkbkPT Is Nothing Then wkbkPT.Close False
    errs.RecordErr "CleanupProdtypesProcedure", CleanupProdtypesProcedure
End Function
'-------------------------------------------------------------------------------------
'-------------------------------------------------------------------------------------
' ValidateRawDataProcedure - coordinate all raw-data quality gates
'-------------------------------------------------------------------------------------
'-------------------------------------------------------------------------------------
' Coordinate all raw-data quality gates; fail-fast on first child failure
' JDL 04/02/26
'
Public Function ValidateRawDataProcedure(import As BRImport, tbls) As Boolean
    setErrs ValidateRawDataProcedure, ThisWorkbook
    If errs.IsHandle Then On Error GoTo ErrorExit

    If Not import.ValidateLocationSubset(import) Then GoTo ErrorExit
    'If Not import.ValidateNoDups(import) Then GoTo ErrorExit
    If Not import.ValidateContiguousDates(import, tbls) Then GoTo ErrorExit
    Exit Function

ErrorExit:
    errs.RecordErr "ValidateRawDataProcedure", ValidateRawDataProcedure
End Function
'--------------------------------------------------------------------------------------
' Confirm selected location is populated and present in tblNormBR rows
' JDL 04/02/26; Updated 4/27/26
'
Public Function ValidateLocationSubset(import As BRImport) As Boolean
    setErrs ValidateLocationSubset, ThisWorkbook
    If errs.IsHandle Then On Error GoTo ErrorExit

    ' Confirm tblNormBR has rows (location filter applied in OpenAndNormalize)
    With import
        If errs.IsFail(.tblNormBR.rngrows Is Nothing, 1, .locationBR) Then GoTo ErrorExit
    End With
    Exit Function

ErrorExit:
    errs.RecordErr "ValidateLocationSubset", ValidateLocationSubset
End Function
'--------------------------------------------------------------------------------------
' Confirm no duplicate rows at import grain using dictionary composite-key approach
' JDL 04/02/26; Updated 4/21/26
'
Public Function ValidateNoDups(import As BRImport) As Boolean
    setErrs ValidateNoDups, ThisWorkbook
    If errs.IsHandle Then On Error GoTo ErrorExit
    Dim dict As Object, rngRow As Range, compositeKey As String
    Dim wkVal As String, yrVal As String, ptVal As String

    ' Instance dictionary; iterate tblNormBR.rngRows
    Set dict = ExcelSteps.New_Dictionary
    With import
        Set rngRow = .tblNormBR.rngrows.Rows(1)
        Do While Not Intersect(rngRow, .tblNormBR.rngrows) Is Nothing
            ptVal = Trim(CStr(Intersect(rngRow, .colrngNormProdtype).Value2))
            wkVal = Trim(CStr(Intersect(rngRow, .colrngNormWkSerial).Value2))
            yrVal = Trim(CStr(Intersect(rngRow, .colrngNormYear).Value2))

            ' Build composite key; fail-fast on first duplicate
            compositeKey = ptVal & "_" & wkVal & "_" & yrVal
            If errs.IsFail(dict.Exists(compositeKey), 1, compositeKey) Then GoTo ErrorExit

            ' Add key and continue
            dict.Add compositeKey, True
            Set rngRow = rngRow.Offset(1, 0)
        Loop
    End With
    Exit Function

ErrorExit:
    errs.RecordErr "ValidateNoDups", ValidateNoDups
End Function
'--------------------------------------------------------------------------------------
' Confirm fiscal weeks are contiguous for the import range using Calendar lookup
' JDL 04/02/26; Updated 4/21/26
'
Public Function ValidateContiguousDates(import As BRImport, tbls) As Boolean
    setErrs ValidateContiguousDates, ThisWorkbook
    If errs.IsHandle Then On Error GoTo ErrorExit
    Dim dictPresent As Object, rngRow As Range
    Dim wkVal As String, yrVal As String, wkYrKey As String
    Dim minWk As Long, maxWk As Long, minYr As Long, maxYr As Long
    Dim minKey As String, maxKey As String
    Dim colrngCalWk As Range, colrngCalYr As Range, colrngCalWkYr As Range
    Dim rngCalRow As Range, calWk As Long, calYr As Long, calKey As String
    Dim inRange As Boolean, pastMax As Boolean

    ' Collect unique (wk_serial_hist, year_hist) pairs into dictionary
    Set dictPresent = ExcelSteps.New_Dictionary
    minWk = 99: maxWk = 0: minYr = 9999: maxYr = 0
    With import
        Set rngRow = .tblNormBR.rngrows.Rows(1)
        Do While Not Intersect(rngRow, .tblNormBR.rngrows) Is Nothing
            wkVal = Trim(CStr(Intersect(rngRow, .colrngNormWkSerial).Value2))
            yrVal = Trim(CStr(Intersect(rngRow, .colrngNormYear).Value2))
            wkYrKey = wkVal & "_" & yrVal
            If Not dictPresent.Exists(wkYrKey) Then dictPresent.Add wkYrKey, True
            ' Step 2: Track min/max year first, then week within year
            Dim wkL As Long, yrL As Long
            wkL = CLng(wkVal): yrL = CLng(yrVal)
            If yrL < minYr Or (yrL = minYr And wkL < minWk) Then
                minWk = wkL: minYr = yrL: minKey = wkYrKey
            End If
            If yrL > maxYr Or (yrL = maxYr And wkL > maxWk) Then
                maxWk = wkL: maxYr = yrL: maxKey = wkYrKey
            End If
            Set rngRow = rngRow.Offset(1, 0)
        Loop
    End With

    With tbls.Calendar
    
        ' Look up Calendar column ranges for Week_Number, Year, wk_year
        Set colrngCalWk = .rngTblHeaderVal(tbls.Calendar, "Week_Number").EntireColumn
        Set colrngCalYr = .rngTblHeaderVal(tbls.Calendar, "Year").EntireColumn
        Set colrngCalWkYr = .rngTblHeaderVal(tbls.Calendar, "wk_year").EntireColumn

        ' Validate min key exists in Calendar
        Dim rngMin As Range
        Set rngMin = ExcelSteps.FindInRange(Intersect(colrngCalWkYr, .rngrows), minKey)
        If errs.IsFail(rngMin Is Nothing, 1, minKey) Then GoTo ErrorExit
    
        ' Step 4: Iterate Calendar rows from min to max; confirm each expected week is present
        inRange = False: pastMax = False
        Set rngCalRow = tbls.Calendar.rngrows.Rows(1)
    
        Do While Not Intersect(rngCalRow, .rngrows) Is Nothing And Not pastMax
            calKey = Trim(CStr(Intersect(rngCalRow, colrngCalWkYr).Value2))
            If calKey = minKey Then inRange = True
            If inRange Then
                If errs.IsFail(Not dictPresent.Exists(calKey), 2, calKey) Then GoTo ErrorExit
                If calKey = maxKey Then pastMax = True
            End If
            Set rngCalRow = rngCalRow.Offset(1, 0)
        Loop
    End With
    Exit Function

ErrorExit:
    errs.RecordErr "ValidateContiguousDates", ValidateContiguousDates
End Function

' <Other code not included in this snippet>