Attribute VB_Name = "tests_ToolboxClasses"
Option Explicit
'Version 5/5/26

' colinfo_ mockup data
Const ColInfo_BRTable_nrows As Long = 9 'n data  rows in BR_Example table in test_data/ColInfo.xlsx
Const ColInfo_BRTable_rngRows As String = "$2:$10" '.tbl.rngRows address for 9 rows
Const ColInfo_aryIndices_BRExample As String = "Location,ProdType,Year,SerialWeek"
Const ColInfo_aryMetrics_BRExample As String = "Net_Sales,Discounts,Markdowns,COGS"
Const ColInfo_colRng_BRExample As String = "$G:$G"
Const ColInfo_dictNormalize_BRExample_Size As Long = 8
Const ColInfo_Locn_VarnameRaw_BRExample As String = "Locn_Raw"
Const ColInfo_Sales_VarnameRaw_BRExample As String = "Sales_Raw"
Const ImportFile_BR_Example As String = "BR_Raw_Mockup.xlsx"
Const ImportFile_SecondTbl As String = "Second_Raw_Mockup.xlsx"
Const ImportNormHeader_BR_Example As String = "Location,ProdType,Year,SerialWeek,Net_Sales,Discounts,Markdowns,COGS"
Const ImportFilteredOnlineRows As Long = 5

'-----------------------------------------------------------------------------------------
' Validate ProjFiles and ColInfo classes
' JDL 4/29/26; Updated 5/6/26
'
Sub TestDriver_ToolBox()
    Dim procs As New Procedures, AllEnabled As Boolean
    With procs
        .Init procs, ThisWorkbook, "ToolBox", "Tests_ToolboxClasses"
        SetApplEnvir False, False, xlCalculationAutomatic

        AllEnabled = False
        .ProjFiles.Enabled = False
        .ColInfo.Enabled = True
        .ImportParseNorm.Enabled = True
    End With

    With procs.ProjFiles
        If .Enabled Or AllEnabled Then
            procs.curProcedure = .name
            test_Init procs
        End If
    End With

    With procs.ColInfo
        If .Enabled Or AllEnabled Then
            procs.curProcedure = .name
            test_ColInfo_Init procs
            test_ColInfo_SetCurTbl procs
            test_ColInfo_YieldAryIndices procs
            test_ColInfo_YieldAryIndices_Empty procs
            test_ColInfo_YieldAryMetrics procs
            test_ColInfo_YieldDNormalize procs
            test_ColInfo_SetCurTblReentrant procs
        End If
    End With

    With procs.ImportParseNorm
        If .Enabled Or AllEnabled Then
            procs.curProcedure = .name
            test_ImportParseNorm_Init procs
            test_ImportParseNorm_OpenRawData procs
            test_ImportParseNorm_ValidateRawStructure procs
            test_ImportParseNorm_BuildNormMappings procs
            test_ImportParseNorm_ApplyFillMapToSortedColumn procs
            test_ImportParseNorm_FillMissingVals procs
            test_ImportParseNorm_WriteNormalized procs
            test_ImportParseNorm_FilterRows procs
        End If
    End With

    procs.EvalOverall procs
End Sub
'-----------------------------------------------------------------------------------------
'-----------------------------------------------------------------------------------------
' procs.ImportParseNorm
'-----------------------------------------------------------------------------------------
'-----------------------------------------------------------------------------------------
' Initialize ImportParseNorm and ensure required attributes are set
' JDL 5/7/26
'
Sub test_ImportParseNorm_Init(procs)
    Dim tst As New Test: tst.Init tst, "test_ImportParseNorm_Init", ThisWorkbook
    Dim importtbl As Object, colinfo As Object, files As Object
    Dim dParamsImport As Object, dParamsParse As Object

    With tst
        InitImportParseNormTest tst, importtbl, colinfo, files, dParamsImport, dParamsParse, "BR_Example"

        .Assert tst, Not importtbl.colinfo Is Nothing
        .Assert tst, Not importtbl.files Is Nothing
        .Assert tst, importtbl.curTbl = "BR_Example"
        .Assert tst, Not importtbl.dParamsImport Is Nothing
        .Assert tst, Not importtbl.dParamsParse Is Nothing
        .Assert tst, Not importtbl.colinfo.rngRowsCurTbl Is Nothing
        .Update tst, procs
    End With
    CloseImportParseNormWkbk importtbl
    CloseColInfoWkbk colinfo
    ExcelSteps.IsTest = False
End Sub
'-----------------------------------------------------------------------------------------
' Build ordered normalization arrays from colinfo metadata
' JDL 5/7/26
'
Sub test_ImportParseNorm_BuildNormMappings(procs)
    Dim tst As New Test: tst.Init tst, "test_ImportParseNorm_BuildNormMappings", ThisWorkbook
    Dim importtbl As Object, colinfo As Object, files As Object
    Dim dParamsImport As Object, dParamsParse As Object
    Dim aryNorm() As String, aryRaw() As String, maxOrder As Long

    With tst
        InitImportParseNormTest tst, importtbl, colinfo, files, dParamsImport, dParamsParse, "BR_Example"

        .Assert tst, importtbl.BuildNormMappings(importtbl, aryNorm, aryRaw, maxOrder)
        .Assert tst, maxOrder = 8
        .Assert tst, aryNorm(1) = "Location"
        .Assert tst, aryRaw(1) = "Locn_Raw"
        .Assert tst, aryNorm(8) = "COGS"
        .Assert tst, aryRaw(8) = "COGS_Raw"
        .Update tst, procs
    End With
    CloseImportParseNormWkbk importtbl
    CloseColInfoWkbk colinfo
    ExcelSteps.IsTest = False
End Sub
'-----------------------------------------------------------------------------------------
' Apply FillVals dictionary to a sorted column via helper method
' JDL 5/7/26
'
Sub test_ImportParseNorm_ApplyFillMapToSortedColumn(procs)
    Dim tst As New Test: tst.Init tst, "test_ImportParseNorm_ApplyFillMapToSortedColumn", ThisWorkbook
    Dim importtbl As Object, colinfo As Object, files As Object
    Dim dParamsImport As Object, dParamsParse As Object, dict As Object
    Dim rngProdHdr As Range, rngProdData As Range

    With tst
        InitImportParseNormTest tst, importtbl, colinfo, files, dParamsImport, dParamsParse, "BR_Example"

        .Assert tst, importtbl.OpenRawData(importtbl)
        Set rngProdHdr = importtbl.tblRaw.rngTblHeaderVal(importtbl.tblRaw, "ProdType_Raw")
        .Assert tst, Not rngProdHdr Is Nothing
        .Assert tst, importtbl.tblRaw.TblSortBy(importtbl.tblRaw, CStr(rngProdHdr.Value2))
        Set rngProdData = Intersect(importtbl.tblRaw.rngRows, rngProdHdr.EntireColumn)

        Set dict = ExcelSteps.New_Dictionary
        .Assert tst, dict.ParseStringToDictProcedure("{BLANK:Unknown,Locn10:Locn1}")
        .Assert tst, importtbl.ApplyFillMapToSortedColumn(importtbl, rngProdData, dict)

        .Assert tst, Not FindInRange(rngProdData, "Unknown") Is Nothing
        .Assert tst, Not FindInRange(rngProdData, "Locn1") Is Nothing
        .Assert tst, FindInRange(rngProdData, "Locn10") Is Nothing
        .Update tst, procs
    End With
    CloseImportParseNormWkbk importtbl
    CloseColInfoWkbk colinfo
    ExcelSteps.IsTest = False
End Sub
'-----------------------------------------------------------------------------------------
' Open raw file into temporary workbook and provision tblRaw
' JDL 5/7/26
'
Sub test_ImportParseNorm_OpenRawData(procs)
    Dim tst As New Test: tst.Init tst, "test_ImportParseNorm_OpenRawData", ThisWorkbook
    Dim importtbl As Object, colinfo As Object, files As Object
    Dim dParamsImport As Object, dParamsParse As Object

    With tst
        InitImportParseNormTest tst, importtbl, colinfo, files, dParamsImport, dParamsParse, "BR_Example"

        .Assert tst, importtbl.OpenRawData(importtbl)
        .Assert tst, Not importtbl.tblRaw Is Nothing
        .Assert tst, TypeName(importtbl.tblRaw) = "tblRowsCols"
        .Assert tst, Not importtbl.tblRaw.wkbk Is Nothing
        .Assert tst, importtbl.tblRaw.wkbk.Name <> files.fColInfo
        .Update tst, procs
    End With
    CloseImportParseNormWkbk importtbl
    CloseColInfoWkbk colinfo
    ExcelSteps.IsTest = False
End Sub
'-----------------------------------------------------------------------------------------
' Validate raw headers satisfy required VarNameRaw fields for current table
' JDL 5/7/26
'
Sub test_ImportParseNorm_ValidateRawStructure(procs)
    Dim tst As New Test: tst.Init tst, "test_ImportParseNorm_ValidateRawStructure", ThisWorkbook
    Dim importtbl As Object, colinfo As Object, files As Object
    Dim dParamsImport As Object, dParamsParse As Object

    With tst
        InitImportParseNormTest tst, importtbl, colinfo, files, dParamsImport, dParamsParse, "BR_Example"

        .Assert tst, importtbl.OpenRawData(importtbl)
        .Assert tst, importtbl.ValidateRawStructure(importtbl)
        .Update tst, procs
    End With
    CloseImportParseNormWkbk importtbl
    CloseColInfoWkbk colinfo
    ExcelSteps.IsTest = False
End Sub
'-----------------------------------------------------------------------------------------
' Apply FillVals from colinfo to raw table values
' JDL 5/7/26
'
Sub test_ImportParseNorm_FillMissingVals(procs)
    Dim tst As New Test: tst.Init tst, "test_ImportParseNorm_FillMissingVals", ThisWorkbook
    Dim importtbl As Object, colinfo As Object, files As Object
    Dim dParamsImport As Object, dParamsParse As Object
    Dim rngProdHdr As Range, rngProdData As Range

    With tst
        InitImportParseNormTest tst, importtbl, colinfo, files, dParamsImport, dParamsParse, "BR_Example"

        .Assert tst, importtbl.OpenRawData(importtbl)
        .Assert tst, importtbl.ValidateRawStructure(importtbl)
        .Assert tst, importtbl.FillMissingVals(importtbl)

        Set rngProdHdr = importtbl.tblRaw.rngTblHeaderVal(importtbl.tblRaw, "ProdType_Raw")
        .Assert tst, Not rngProdHdr Is Nothing
        Set rngProdData = Intersect(importtbl.tblRaw.rngRows, rngProdHdr.EntireColumn)

        ' Check BLANK fill and value replacement from FillVals metadata
        .Assert tst, rngProdData.Cells(2, 1).Value2 = "Unknown"
        .Assert tst, rngProdData.Cells(3, 1).Value2 = "Locn1"
        .Assert tst, rngProdData.Cells(5, 1).Value2 = "Unknown"
        .Update tst, procs
    End With
    CloseImportParseNormWkbk importtbl
    CloseColInfoWkbk colinfo
    ExcelSteps.IsTest = False
End Sub
'-----------------------------------------------------------------------------------------
' Write normalized table using ordered CurTbl metadata columns
' JDL 5/7/26
'
Sub test_ImportParseNorm_WriteNormalized(procs)
    Dim tst As New Test: tst.Init tst, "test_ImportParseNorm_WriteNormalized", ThisWorkbook
    Dim importtbl As Object, colinfo As Object, files As Object
    Dim dParamsImport As Object, dParamsParse As Object

    With tst
        InitImportParseNormTest tst, importtbl, colinfo, files, dParamsImport, dParamsParse, "BR_Example"

        .Assert tst, importtbl.OpenAndValidateRawProcedure(importtbl)
        .Assert tst, importtbl.ParseRawProcedure(importtbl)
        .Assert tst, importtbl.WriteNormalized(importtbl)

        .Assert tst, Not importtbl.tblNorm Is Nothing
        .Assert tst, importtbl.tblNorm.sht = "norm_"
        .Assert tst, ListFromArray(importtbl.tblNorm.rngHeader.Value2) = ImportNormHeader_BR_Example
        .Assert tst, importtbl.tblNorm.nRows = 8
        .Update tst, procs
    End With
    CloseImportParseNormWkbk importtbl
    CloseColInfoWkbk colinfo
    ExcelSteps.IsTest = False
End Sub
'-----------------------------------------------------------------------------------------
' Filter normalized rows based on KeepOnly setting from colinfo metadata
' JDL 5/7/26
'
Sub test_ImportParseNorm_FilterRows(procs)
    Dim tst As New Test: tst.Init tst, "test_ImportParseNorm_FilterRows", ThisWorkbook
    Dim importtbl As Object, colinfo As Object, files As Object
    Dim dParamsImport As Object, dParamsParse As Object
    Dim rngLocnHdr As Range, rngLocnData As Range, cell As Range

    With tst
        InitImportParseNormTest tst, importtbl, colinfo, files, dParamsImport, dParamsParse, "BR_Example"

        .Assert tst, importtbl.OpenAndValidateRawProcedure(importtbl)
        .Assert tst, importtbl.ParseRawProcedure(importtbl)
        .Assert tst, importtbl.WriteNormalized(importtbl)
        .Assert tst, importtbl.FilterRows(importtbl)

        .Assert tst, importtbl.tblNorm.nRows = ImportFilteredOnlineRows
        Set rngLocnHdr = importtbl.tblNorm.rngTblHeaderVal(importtbl.tblNorm, "Location")
        .Assert tst, Not rngLocnHdr Is Nothing
        Set rngLocnData = Intersect(importtbl.tblNorm.rngRows, rngLocnHdr.EntireColumn)
        For Each cell In rngLocnData.Cells
            .Assert tst, cell.Value2 = "Online"
        Next cell
        .Update tst, procs
    End With
    CloseImportParseNormWkbk importtbl
    CloseColInfoWkbk colinfo
    ExcelSteps.IsTest = False
End Sub
'-----------------------------------------------------------------------------------------
'-----------------------------------------------------------------------------------------
' procs.ProjFiles
'-----------------------------------------------------------------------------------------
'-----------------------------------------------------------------------------------------
' Set path/filename attributes for project workbook and test suite locations
' JDL 4/29/26; updated 5/1/26
Sub test_Init(procs)
    Dim tst As New Test: tst.Init tst, "test_Init", ThisWorkbook
    Dim files As Object, sep As String, wkbkStepsAddin As Workbook, aryPath As Variant
    Dim dirLast As String, dirNextToLast

    With tst
        '(project being tested name is ExcelSteps in this case. Its file is XLSteps.xlam)
        Set wkbkStepsAddin = GetWorkbookByVBProjectName(vba_project_name)
        
        'Check Init succeeds
        Set files = ExcelSteps.New_ProjFiles
        .Assert tst, files.Init(files, wkbkStepsAddin)

        'Check pathSrc set to project workbook path with trailing sep
        sep = Application.PathSeparator
        aryPath = Split(files.pathSrc, sep)
        dirNextToLast = aryPath(UBound(aryPath) - 1)
        dirLast = aryPathv(UBound(aryPath))
        .Assert tst, dirLast = "" 'Proves trailing sep present
        .Assert tst, dirNextToLast = "src"
        
        'Check IsDevelopment True (project workbook is in src folder)
        .Assert tst, files.IsDevelopment = True

        'Check fColInfo set
        .Assert tst, files.fColInfo = "ColInfo.xlsx"

        'Check pfColInfo = pathColInfo & fColInfo (trailing sep already on pathColInfo)
        .Assert tst, files.pfColInfo = files.pathColInfo & files.fColInfo
        .Update tst, procs
    End With
End Sub
'-----------------------------------------------------------------------------------------
'-----------------------------------------------------------------------------------------
' procs.ColInfo
'-----------------------------------------------------------------------------------------
'-----------------------------------------------------------------------------------------
' Open ColInfo.xlsx and provision colinfo.tbl
' JDL 5/1/26
'
Sub test_ColInfo_Init(procs)
    Dim tst As New Test: tst.Init tst, "test_ColInfo_Init", ThisWorkbook
    Dim colinfo As Object, files As Object, wkbkStepsAddin As Workbook
    Dim sep as string: sep = Application.PathSeparator

    With tst
        ExcelSteps.IsTest = True
        Set files = ExcelSteps.New_ProjFiles

        ' Initialize files and check directories
        Set wkbkStepsAddin = GetWorkbookByVBProjectName(vba_project_name)
        .Assert tst, files.Init(files, wkbkStepsAddin, "test_data")
        '.Assert tst, Left(files.pfColInfo, 2) = sep & sep
        .Assert tst, Not Dir(files.pfColInfo) = ""
        .Assert tst, Right(files.pfColInfo, Len(files.fColInfo)) = files.fColInfo
        
        ' Instance colinfo and check its Init
        Set colinfo = ExcelSteps.New_ColInfo

        'Check Init succeeds and tbl provisioned
        .Assert tst, colinfo.Init(colinfo, files)
        .Assert tst, Not colinfo.tbl Is Nothing
        .Assert tst, colinfo.tbl.sht = "colinfo_"
        .Assert tst, Not colinfo.tbl.wkbk Is Nothing
        .Assert tst, Not colinfo.tbl.rngRows Is Nothing

        'Check CurTbl empty when curTbl arg not passed
        .Assert tst, colinfo.CurTbl = ""
        .Assert tst, colinfo.rngRowsCurTbl Is Nothing
        .Update tst, procs
    End With
    CloseColInfoWkbk colinfo
    ExcelSteps.IsTest = False
End Sub
'-----------------------------------------------------------------------------------------
' Sort colinfo.tbl and set rngRowsCurTbl for specified table
' JDL 5/1/26
'
Sub test_ColInfo_SetCurTbl(procs)
    Dim tst As New Test: tst.Init tst, "test_ColInfo_SetCurTbl", ThisWorkbook
    Dim colinfo As Object, files As Object

    With tst
        InitColInfoTest tst, colinfo, files, "BR_Example"

        'Check CurTbl set
        .Assert tst, colinfo.CurTbl = "BR_Example"

        'Check rngRowsCurTbl address and n rows
        .Assert tst, Not colinfo.rngRowsCurTbl Is Nothing
        .Assert tst, colinfo.rngRowsCurTbl.Rows.Count =ColInfo_BRTable_nrows
        .Assert tst, colinfo.rngRowsCurTbl.Address = ColInfo_BRTable_rngRows

        'Check that .colrngTblName set to correct column range for "BR_Example"
        .Assert tst, colinfo.tbl.colrngTblName.Address = ColInfo_colRng_BRExample
        .Update tst, procs
    End With
    CloseColInfoWkbk colinfo
    ExcelSteps.IsTest = False
End Sub
'-----------------------------------------------------------------------------------------
' Return index VarNameNorm array for curTbl
' JDL 5/1/26
'
Sub test_ColInfo_YieldAryIndices(procs)
    Dim tst As New Test: tst.Init tst, "test_ColInfo_YieldAryIndices", ThisWorkbook
    Dim colinfo As Object, files As Object, ary As Variant

    With tst
        InitColInfoTest tst, colinfo, files, "BR_Example"

        'Check indices returned
        ary = colinfo.YieldAryIndices(colinfo)
        .Assert tst, IsArray(ary)
        .Assert tst, UBound(ary) >= 0

        'Check first index matches expected (Location is first index in BR_Example)
        .Assert tst, ListFromArray(ary) = ColInfo_aryIndices_BRExample 

        .Update tst, procs
    End With
    CloseColInfoWkbk colinfo
    ExcelSteps.IsTest = False
End Sub
'-----------------------------------------------------------------------------------------
' YieldAryIndices returns empty array when curTbl has no index rows
' JDL 5/6/26
'
Sub test_ColInfo_YieldAryIndices_Empty(procs)
    Dim tst As New Test: tst.Init tst, "test_ColInfo_YieldAryIndices_Empty", ThisWorkbook
    Dim colinfo As Object, files As Object, ary As Variant

    With tst
        InitColInfoTest tst, colinfo, files, "Second_Tbl"

        'Check empty array returned for table with no index rows
        ary = colinfo.YieldAryIndices(colinfo)
        .Assert tst, IsArray(ary)
        .Assert tst, UBound(ary) = -1
        .Update tst, procs
    End With
    CloseColInfoWkbk colinfo
    ExcelSteps.IsTest = False
End Sub
'-----------------------------------------------------------------------------------------
' Return metric VarNameNorm array for curTbl
' JDL 5/1/26
'
Sub test_ColInfo_YieldAryMetrics(procs)
    Dim tst As New Test: tst.Init tst, "test_ColInfo_YieldAryMetrics", ThisWorkbook
    Dim colinfo As Object, files As Object, ary As Variant

    With tst
        InitColInfoTest tst, colinfo, files, "BR_Example"

        'Check metrics returned
        ary = colinfo.YieldAryMetrics(colinfo)
        .Assert tst, IsArray(ary)
        .Assert tst, UBound(ary) >= 0

        'Check metrics match expected
        .Assert tst, ListFromArray(ary) = ColInfo_aryMetrics_BRExample
        .Update tst, procs
    End With
    CloseColInfoWkbk colinfo
    ExcelSteps.IsTest = False
End Sub
'-----------------------------------------------------------------------------------------
' Return VarNameNorm->VarNameRaw dictionary for curTbl
' JDL 5/1/26
'
Sub test_ColInfo_YieldDNormalize(procs)
    Dim tst As New Test: tst.Init tst, "test_ColInfo_YieldDNormalize", ThisWorkbook
    Dim colinfo As Object, files As Object, dict As Object

    With tst
        InitColInfoTest tst, colinfo, files, "BR_Example"

        'Check dict returned with correct number of entries
        Set dict = colinfo.YieldDNormalize(colinfo)
        .Assert tst, dict.Size = ColInfo_dictNormalize_BRExample_Size

        'Check known mapping
        .Assert tst, dict.Item("Location") = ColInfo_Locn_VarnameRaw_BRExample
        .Assert tst, dict.Item("Net_Sales") = ColInfo_Sales_VarnameRaw_BRExample
        .Update tst, procs
    End With
    CloseColInfoWkbk colinfo
    ExcelSteps.IsTest = False
End Sub
'-----------------------------------------------------------------------------------------
' SetCurTbl re-entrant: call twice with different table names
' JDL 5/1/26
'
Sub test_ColInfo_SetCurTblReentrant(procs)
    Dim tst As New Test: tst.Init tst, "test_ColInfo_SetCurTblReentrant", ThisWorkbook
    Dim colinfo As Object, files As Object
    Dim nRowsFirst As Long, nRowsSecond As Long

    With tst
        InitColInfoTest tst, colinfo, files, "BR_Example"
        nRowsFirst = colinfo.rngRowsCurTbl.Rows.Count

        'Switch to second table; check CurTbl and row count update
        .Assert tst, colinfo.SetCurTbl(colinfo, "Second_Tbl")
        .Assert tst, colinfo.CurTbl = "Second_Tbl"
        nRowsSecond = colinfo.rngRowsCurTbl.Rows.Count
        .Assert tst, nRowsSecond > 0
        .Assert tst, nRowsSecond <> nRowsFirst
        .Update tst, procs
    End With
    CloseColInfoWkbk colinfo
    ExcelSteps.IsTest = False
End Sub
'-----------------------------------------------------------------------------------------
' Helper: instance and init files and colinfo using test_data/ColInfo.xlsx
' JDL 5/6/26
Sub InitColInfoTest(tst, colinfo As Object, files As Object, Optional curTbl As String)
    Dim wkbkStepsAddin As Workbook, rngRand As Range
    Set wkbkStepsAddin = GetWorkbookByVBProjectName(vba_project_name)
    ExcelSteps.IsTest = True

    ' Instance and init files
    Set files = ExcelSteps.New_ProjFiles
    tst.Assert tst, files.Init(files, wkbkStepsAddin, "test_data")

    ' Instance and init colinfo
    Set colinfo = ExcelSteps.New_ColInfo
    tst.Assert tst, colinfo.Init(colinfo, files)

    ' Optionally set curTbl
    If Len(curTbl) > 0 Then

        ' Sort by rand column to ensure non-sorted order doesn't mask bugs in SetCurTbl
        Set rngRand = colinfo.tbl.rngTblHeaderVal(colinfo.tbl, "rand")
        tst.Assert tst, Not rngRand Is Nothing
        tst.Assert tst, colinfo.tbl.TblSortBy(colinfo.tbl, CStr(rngRand.Value2))
        tst.Assert tst, colinfo.SetCurTbl(colinfo, curTbl)
    End If
End Sub
'-----------------------------------------------------------------------------------------
' Helper: initialize ImportParseNorm with colinfo/files/params for a curTbl
' JDL 5/7/26
Sub InitImportParseNormTest(tst, importtbl As Object, colinfo As Object, files As Object, _
    dParamsImport As Object, dParamsParse As Object, ByVal curTbl As String)
    Dim wkbkStepsAddin As Workbook

    Set wkbkStepsAddin = GetWorkbookByVBProjectName(vba_project_name)
    ExcelSteps.IsTest = True

    Set files = ExcelSteps.New_ProjFiles
    tst.Assert tst, files.Init(files, wkbkStepsAddin, "test_data")

    If curTbl = "Second_Tbl" Then
        files.pfImportFile = files.pathData & ImportFile_SecondTbl
    Else
        files.pfImportFile = files.pathData & ImportFile_BR_Example
    End If
    tst.Assert tst, Len(Dir(files.pfImportFile)) > 0

    Set dParamsImport = ExcelSteps.New_Dictionary
    dParamsImport.Add "FileType", "xlsx"

    Set dParamsParse = ExcelSteps.New_Dictionary
    dParamsParse.Add "RawShape", "rowscols"

    Set colinfo = ExcelSteps.New_ColInfo
    tst.Assert tst, colinfo.Init(colinfo, files)

    Set importtbl = ExcelSteps.New_ImportParseNorm
    tst.Assert tst, importtbl.Init(importtbl, colinfo, files, curTbl, dParamsImport, dParamsParse)
End Sub
'-----------------------------------------------------------------------------------------
' Helper: close colinfo workbook opened by Init/Provision if present
'
Sub CloseColInfoWkbk(colinfo As Object)
    If colinfo Is Nothing Then Exit Sub
    If colinfo.tbl Is Nothing Then Exit Sub
    If colinfo.tbl.wkbk Is Nothing Then Exit Sub
    colinfo.tbl.wkbk.Close False
End Sub
'-----------------------------------------------------------------------------------------
' Helper: close import workbook(s) opened by ImportParseNorm if present
'
Sub CloseImportParseNormWkbk(importtbl As Object)
    If importtbl Is Nothing Then Exit Sub
    If importtbl.tblRaw Is Nothing Then Exit Sub
    If TypeName(importtbl.tblRaw) = "tblRowsCols" Then
        If importtbl.tblRaw.wkbk Is Nothing Then Exit Sub
        importtbl.tblRaw.wkbk.Close False
    End If
End Sub