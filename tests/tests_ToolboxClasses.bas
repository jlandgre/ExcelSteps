Attribute VB_Name = "tests_ToolboxClasses"
Option Explicit
'Version 5/19/26

' colinfo_ mockup data
Const ColInfo_BRTable_nrows As Long = 9 'n data  rows in BR_Example table in test_data/ColInfo.xlsx
Const ColInfo_BRTable_rngRows As String = "$2:$10" '.tbl.rngRows address for 9 rows
Const ColInfo_aryIndices_BRExample As String = "Location,ProdType,Year,SerialWeek"
Const ColInfo_aryMetrics_BRExample As String = "Net_Sales,Discounts,Markdowns,COGS"
Const ColInfo_colRng_BRExample As String = "$J:$J"
Const ColInfo_dictNormalize_BRExample_Count As Long = 8
Const ColInfo_Locn_VarnameRaw_BRExample As String = "Locn_Raw"
Const ColInfo_Sales_VarnameRaw_BRExample As String = "Sales_Raw"

'ImportParseNorm test data
Const ImportFile_BR_Example As String = "BR_Raw_Mockup.xlsx"
Const ImportFile_BR_KeepExcept As String = "BR_Raw_Mockup_KeepExcept.xlsx"
Const ImportFile_SecondTbl As String = "Second_Raw_Mockup.xlsx"
Const ImportNormHeader_BR_Example As String = "Location,ProdType,Year,SerialWeek,Net_Sales,Discounts,Markdowns,COGS"
Const ImportFilteredOnlineRows As Long = 5
Const ImportFilteredKeepExceptBlankRows As Long = 8
Const ImportFilteredKeepExceptBlankOnlineRows As Long = 3

'-----------------------------------------------------------------------------------------
' Validate ProjFiles and ColInfo classes
' JDL 4/29/26; Updated 5/19/26
'
Sub TestDriver_ToolBox()
    Dim procs As New Procedures, AllEnabled As Boolean
    With procs
        .Init procs, ThisWorkbook, "Tests_ToolBox", "Tests_Toolbox"
        SetApplEnvir False, False, xlCalculationAutomatic

        AllEnabled = True
        .ProjFiles.Enabled = False
        .colinfo.Enabled = False
        .ImportParseNorm.Enabled = True
    End With

    ExcelSteps.IsTest = True

    With procs.ProjFiles
        If .Enabled Or AllEnabled Then
            procs.curProcedure = .name
            test_Init procs
        End If
    End With

    With procs.colinfo
        If .Enabled Or AllEnabled Then
            procs.curProcedure = .name
            test_ColInfo_Init procs
            test_ColInfo_Init_ThisWorkbook procs
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
            test_ImportParseNorm_ReplaceVals procs
            test_ImportParseNorm_WriteNormalized procs
            test_ImportParseNorm_FilterRows procs
            test_ImportParseNorm_KeepExcept_BLANK procs
            test_ImportParseNorm_KeepExcept_BLANK_Online procs
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
    Dim imptbl As Object, colinfo As Object, files As Object
    Dim dParamsImport As Object, dParamsParse As Object

    With tst
        InitImportParseNormTest tst, imptbl, colinfo, files, dParamsImport, dParamsParse, "BR_Example"

        .Assert tst, Not imptbl.colinfo Is Nothing
        .Assert tst, Not imptbl.files Is Nothing
        .Assert tst, imptbl.curTbl = "BR_Example"
        .Assert tst, Not imptbl.dParamsImport Is Nothing
        .Assert tst, Not imptbl.dParamsParse Is Nothing
        .Assert tst, imptbl.colinfo.rngRowsCurTbl.Address = ColInfo_BRTable_rngRows
        .Update tst, procs
    End With
    CloseColInfoWkbk colinfo
End Sub
'-----------------------------------------------------------------------------------------
' Open raw file into temporary workbook and provision tblRaw
' JDL 5/7/26
'
Sub test_ImportParseNorm_OpenRawData(procs)
    Dim tst As New Test: tst.Init tst, "test_ImportParseNorm_OpenRawData", ThisWorkbook
    Dim imptbl As Object, colinfo As Object, files As Object
    Dim dParamsImport As Object, dParamsParse As Object

    With tst
        InitImportParseNormTest tst, imptbl, colinfo, files, dParamsImport, dParamsParse, "BR_Example"

        .Assert tst, imptbl.OpenRawData(imptbl)
        .Assert tst, Not imptbl.tblRaw Is Nothing
        .Assert tst, TypeName(imptbl.tblRaw) = "tblRowsCols"
        .Assert tst, Not imptbl.tblRaw.wkbk Is Nothing
        .Assert tst, imptbl.tblRaw.wkbk.name <> files.fColInfo
        .Update tst, procs
    End With
    CloseImportParseNormWkbk imptbl
    CloseColInfoWkbk colinfo
End Sub
'-----------------------------------------------------------------------------------------
' Validate raw headers satisfy required VarNameRaw fields for current table
' JDL 5/7/26
'
Sub test_ImportParseNorm_ValidateRawStructure(procs)
    Dim tst As New Test: tst.Init tst, "test_ImportParseNorm_ValidateRawStructure", ThisWorkbook
    Dim imptbl As Object, colinfo As Object, files As Object
    Dim dParamsImport As Object, dParamsParse As Object

    With tst
        InitImportParseNormTest tst, imptbl, colinfo, files, dParamsImport, dParamsParse, "BR_Example"

        .Assert tst, imptbl.OpenRawData(imptbl)
        .Assert tst, imptbl.ValidateRawStructure(imptbl)
        .Update tst, procs
    End With
    CloseImportParseNormWkbk imptbl
    CloseColInfoWkbk colinfo
End Sub
'-----------------------------------------------------------------------------------------
' Build ordered normalization arrays from colinfo metadata
' JDL 5/7/26
'
Sub test_ImportParseNorm_BuildNormMappings(procs)
    Dim tst As New Test: tst.Init tst, "test_ImportParseNorm_BuildNormMappings", ThisWorkbook
    Dim imptbl As Object, colinfo As Object, files As Object
    Dim dParamsImport As Object, dParamsParse As Object
    Dim aryNorm() As String, aryRaw() As String, maxOrder As Long

    With tst
        InitImportParseNormTest tst, imptbl, colinfo, files, dParamsImport, dParamsParse, "BR_Example"

        .Assert tst, imptbl.BuildNormMappings(imptbl, aryNorm, aryRaw, maxOrder)
        .Assert tst, maxOrder = 8
        .Assert tst, aryNorm(1) = "Location"
        .Assert tst, aryRaw(1) = "Locn_Raw"
        .Assert tst, aryNorm(8) = "COGS"
        .Assert tst, aryRaw(8) = "COGS_Raw"
        .Update tst, procs
    End With
    CloseImportParseNormWkbk imptbl
    CloseColInfoWkbk colinfo
End Sub
'-----------------------------------------------------------------------------------------
' Apply FillVals dictionary to a sorted column via helper method
' JDL 5/7/26
'
Sub test_ImportParseNorm_ApplyFillMapToSortedColumn(procs)
    Dim tst As New Test: tst.Init tst, "test_ImportParseNorm_ApplyFillMapToSortedColumn", ThisWorkbook
    Dim imptbl As Object, colinfo As Object, files As Object
    Dim dParamsImport As Object, dParamsParse As Object, dict As Object
    Dim rngProdHdr As Range, rngProdData As Range

    With tst
        InitImportParseNormTest tst, imptbl, colinfo, files, dParamsImport, dParamsParse, "BR_Example"

        .Assert tst, imptbl.OpenRawData(imptbl)
        Set rngProdHdr = imptbl.tblRaw.rngTblHeaderVal(imptbl.tblRaw, "ProdType_Raw")
        .Assert tst, Not rngProdHdr Is Nothing
        .Assert tst, imptbl.tblRaw.TblSortBy(imptbl.tblRaw, CStr(rngProdHdr.Value2))
        Set rngProdData = Intersect(imptbl.tblRaw.rngRows, rngProdHdr.EntireColumn)

        Set dict = ExcelSteps.New_Dictionary
        .Assert tst, dict.ParseStringToDictProcedure("{BLANK:Unknown,Locn10:Locn1}")
        .Assert tst, imptbl.ApplyFillMapToSortedColumn(imptbl, rngProdData, dict)

        .Assert tst, Not FindInRange(rngProdData, "Unknown") Is Nothing
        .Assert tst, Not FindInRange(rngProdData, "Locn1") Is Nothing
        .Assert tst, FindInRange(rngProdData, "Locn10") Is Nothing
        .Update tst, procs
    End With
    CloseImportParseNormWkbk imptbl
    CloseColInfoWkbk colinfo
End Sub
'-----------------------------------------------------------------------------------------
' Apply FillVals from colinfo to raw table values
' JDL 5/7/26
'
Sub test_ImportParseNorm_ReplaceVals(procs)
    Dim tst As New Test: tst.Init tst, "test_ImportParseNorm_ReplaceVals", ThisWorkbook
    Dim imptbl As Object, colinfo As Object, files As Object
    Dim dParamsImport As Object, dParamsParse As Object
    Dim rngProdHdr As Range, rngProdData As Range, cTgt As Range

    With tst
        InitImportParseNormTest tst, imptbl, colinfo, files, dParamsImport, dParamsParse, "BR_Example"

        .Assert tst, imptbl.OpenRawData(imptbl)
        .Assert tst, imptbl.ValidateRawStructure(imptbl)
        .Assert tst, imptbl.ReplaceVals(imptbl)

        Set rngProdHdr = imptbl.tblRaw.rngTblHeaderVal(imptbl.tblRaw, "ProdType_Raw")
        .Assert tst, Not rngProdHdr Is Nothing
        Set rngProdData = Intersect(imptbl.tblRaw.rngRows, rngProdHdr.EntireColumn)

        ' Check BLANK and Locn10 fill and value replacement
        FindRawTblVal cTgt, "StoreA", 1, "ProdType_Raw", imptbl.tblRaw
        .Assert tst, cTgt.Value2 = "Unknown"
        FindRawTblVal cTgt, "Online", 3, "ProdType_Raw", imptbl.tblRaw
        .Assert tst, cTgt.Value2 = "Unknown"
        FindRawTblVal cTgt, "Online", 2, "ProdType_Raw", imptbl.tblRaw
        .Assert tst, cTgt.Value2 = "Locn1"
        FindRawTblVal cTgt, "StoreC", 4, "ProdType_Raw", imptbl.tblRaw
        .Assert tst, cTgt.Value2 = "Locn1"
        .Update tst, procs
    End With
    CloseImportParseNormWkbk imptbl
    CloseColInfoWkbk colinfo
End Sub
'-----------------------------------------------------------------------------------------
' Lookup to check tbl val based on Locn, Week keys
' JDL 5/8/26
'

Sub FindRawTblVal(cTgt, ValLocn, ValWeek, HeaderTgt, tbl)
    Dim rngColLocn As Range, rngColWeek As Range, rngRow As Range
    Dim aryCols As Variant, aryVals As Variant
        With tbl
            Set rngColLocn = .rngTblHeaderVal(tbl, "Locn_Raw").EntireColumn
            Set rngColWeek = .rngTblHeaderVal(tbl, "Week_Raw").EntireColumn
            aryCols = Array(rngColLocn, rngColWeek)
        End With
        aryVals = Array(ValLocn, ValWeek)
        Set rngRow = ExcelSteps.rngMultiKeyBasic(tbl.rngRows, aryCols, aryVals)
        Set cTgt = Intersect(rngRow, tbl.rngTblHeaderVal(tbl, HeaderTgt).EntireColumn)
End Sub
'-----------------------------------------------------------------------------------------
' Write normalized table using ordered CurTbl metadata columns
' JDL 5/7/26
'
Sub test_ImportParseNorm_WriteNormalized(procs)
    Dim tst As New Test: tst.Init tst, "test_ImportParseNorm_WriteNormalized", ThisWorkbook
    Dim imptbl As Object, colinfo As Object, files As Object
    Dim dParamsImport As Object, dParamsParse As Object

    With tst
        InitImportParseNormTest tst, imptbl, colinfo, files, dParamsImport, dParamsParse, "BR_Example"

        .Assert tst, imptbl.OpenAndValidateRawProcedure(imptbl)
        .Assert tst, imptbl.ParseRawProcedure(imptbl)
        .Assert tst, imptbl.WriteNormalized(imptbl)

        .Assert tst, Not imptbl.tblNorm Is Nothing
        .Assert tst, imptbl.tblNorm.sht = "norm_"
        .Assert tst, ListFromArray(imptbl.tblNorm.rngHeader.Value2) = ImportNormHeader_BR_Example
        .Assert tst, imptbl.tblNorm.nRows = 8
        .Update tst, procs
    End With
    CloseImportParseNormWkbk imptbl
    CloseColInfoWkbk colinfo
    
End Sub
'-----------------------------------------------------------------------------------------
' Filter normalized rows based on KeepOnly setting from colinfo metadata
' JDL 5/7/26
'
Sub test_ImportParseNorm_FilterRows(procs)
    Dim tst As New Test: tst.Init tst, "test_ImportParseNorm_FilterRows", ThisWorkbook
    Dim imptbl As Object, colinfo As Object, files As Object
    Dim dParamsImport As Object, dParamsParse As Object
    Dim rngLocnHdr As Range, rngLocnData As Range, cell As Range

    With tst
        InitImportParseNormTest tst, imptbl, colinfo, files, dParamsImport, dParamsParse, "BR_Example"

        .Assert tst, imptbl.OpenAndValidateRawProcedure(imptbl)
        .Assert tst, imptbl.ParseRawProcedure(imptbl)
        .Assert tst, imptbl.WriteNormalized(imptbl)
        .Assert tst, imptbl.FilterRows(imptbl)

        .Assert tst, imptbl.tblNorm.nRows = ImportFilteredOnlineRows
        Set rngLocnHdr = imptbl.tblNorm.rngTblHeaderVal(imptbl.tblNorm, "Location")
        .Assert tst, Not rngLocnHdr Is Nothing
        Set rngLocnData = Intersect(imptbl.tblNorm.rngRows, rngLocnHdr.EntireColumn)
        For Each cell In rngLocnData.Cells
            .Assert tst, cell.Value2 = "Online"
        Next cell
        .Update tst, procs
    End With
    CloseImportParseNormWkbk imptbl
    CloseColInfoWkbk colinfo
    
End Sub
'-----------------------------------------------------------------------------------------
' Filter normalized rows with KeepExcept BLANK from colinfo metadata
' JDL 5/19/26
'
Sub test_ImportParseNorm_KeepExcept_BLANK(procs)
    Dim tst As New Test: tst.Init tst, "test_ImportParseNorm_KeepExcept_BLANK", ThisWorkbook
    Dim imptbl As Object, colinfo As Object, files As Object
    Dim dParamsImport As Object, dParamsParse As Object
    Dim rngLocnHdr As Range, rngLocnData As Range, cell As Range, rngLocnRow As Range

    With tst
        InitImportParseNormTest tst, imptbl, colinfo, files, dParamsImport, dParamsParse, "BR_Example"

        files.pfImportFile = files.pathData & ImportFile_BR_KeepExcept
        .Assert tst, Len(Dir(files.pfImportFile)) > 0

        Set rngLocnRow = FindInRange(colinfo.tbl.colrngVarNorm, "Location")
        .Assert tst, Not rngLocnRow Is Nothing
        Intersect(rngLocnRow.EntireRow, colinfo.tbl.colrngFilterVals).Value2 = "{KeepExcept:""BLANK""}"

        .Assert tst, imptbl.OpenAndValidateRawProcedure(imptbl)
        .Assert tst, imptbl.ParseRawProcedure(imptbl)
        .Assert tst, imptbl.WriteNormalized(imptbl)
        .Assert tst, imptbl.FilterRows(imptbl)

        .Assert tst, imptbl.tblNorm.nRows = ImportFilteredKeepExceptBlankRows
        Set rngLocnHdr = imptbl.tblNorm.rngTblHeaderVal(imptbl.tblNorm, "Location")
        .Assert tst, Not rngLocnHdr Is Nothing
        Set rngLocnData = Intersect(imptbl.tblNorm.rngRows, rngLocnHdr.EntireColumn)
        For Each cell In rngLocnData.Cells
            .Assert tst, Len(CStr(cell.Value2)) > 0
        Next cell
        .Update tst, procs
    End With
    CloseImportParseNormWkbk imptbl
    CloseColInfoWkbk colinfo

End Sub
'-----------------------------------------------------------------------------------------
' Filter normalized rows with KeepExcept BLANK and Online from colinfo metadata
' JDL 5/19/26
'
Sub test_ImportParseNorm_KeepExcept_BLANK_Online(procs)
    Dim tst As New Test: tst.Init tst, "test_ImportParseNorm_KeepExcept_BLANK_Online", ThisWorkbook
    Dim imptbl As Object, colinfo As Object, files As Object
    Dim dParamsImport As Object, dParamsParse As Object
    Dim rngLocnHdr As Range, rngLocnData As Range, cell As Range, rngLocnRow As Range
    Dim dictSeen As Object

    With tst
        InitImportParseNormTest tst, imptbl, colinfo, files, dParamsImport, dParamsParse, "BR_Example"

        files.pfImportFile = files.pathData & ImportFile_BR_KeepExcept
        .Assert tst, Len(Dir(files.pfImportFile)) > 0

        Set rngLocnRow = FindInRange(colinfo.tbl.colrngVarNorm, "Location")
        .Assert tst, Not rngLocnRow Is Nothing
        Intersect(rngLocnRow.EntireRow, colinfo.tbl.colrngFilterVals).Value2 = "{KeepExcept:""BLANK,Online""}"

        .Assert tst, imptbl.OpenAndValidateRawProcedure(imptbl)
        .Assert tst, imptbl.ParseRawProcedure(imptbl)
        .Assert tst, imptbl.WriteNormalized(imptbl)
        .Assert tst, imptbl.FilterRows(imptbl)

        .Assert tst, imptbl.tblNorm.nRows = ImportFilteredKeepExceptBlankOnlineRows
        Set rngLocnHdr = imptbl.tblNorm.rngTblHeaderVal(imptbl.tblNorm, "Location")
        .Assert tst, Not rngLocnHdr Is Nothing
        Set rngLocnData = Intersect(imptbl.tblNorm.rngRows, rngLocnHdr.EntireColumn)

        Set dictSeen = ExcelSteps.New_Dictionary
        For Each cell In rngLocnData.Cells
            .Assert tst, CStr(cell.Value2) = "StoreA" Or CStr(cell.Value2) = "StoreB" Or _
              CStr(cell.Value2) = "StoreC"
            dictSeen.Add CStr(cell.Value2), True
        Next cell
        .Assert tst, dictSeen.Count = 3
        .Assert tst, dictSeen.Exists("StoreA")
        .Assert tst, dictSeen.Exists("StoreB")
        .Assert tst, dictSeen.Exists("StoreC")
        .Update tst, procs
    End With
    CloseImportParseNormWkbk imptbl
    CloseColInfoWkbk colinfo

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
        dirLast = aryPath(UBound(aryPath))
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
    Dim sep As String: sep = Application.PathSeparator

    With tst
        ExcelSteps.IsTest = True
        Set files = ExcelSteps.New_ProjFiles

        ' Initialize files and check directories
        Set wkbkStepsAddin = GetWorkbookByVBProjectName(vba_project_name)
        .Assert tst, files.Init(files, wkbkStepsAddin, "test_data")
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
        .Assert tst, colinfo.curTbl = ""
        .Assert tst, colinfo.rngRowsCurTbl Is Nothing
        .Update tst, procs
    End With
    CloseColInfoWkbk colinfo
    
End Sub
'-----------------------------------------------------------------------------------------
' Open colinfo_ from ThisWorkbook and provision colinfo.tbl
' JDL 5/11/26
'
Sub test_ColInfo_Init_ThisWorkbook(procs)
    Dim tst As New Test: tst.Init tst, "test_ColInfo_Init_ThisWorkbook", ThisWorkbook
    Dim colinfo As Object, files As Object, wkbkStepsAddin As Workbook
    Dim shtColInfo As Worksheet

    With tst
        ExcelSteps.IsTest = True
        Set wkbkStepsAddin = GetWorkbookByVBProjectName(vba_project_name)
        Set files = ExcelSteps.New_ProjFiles
        .Assert tst, files.Init(files, wkbkStepsAddin, "test_data")

        On Error Resume Next
        Application.DisplayAlerts = False
        ThisWorkbook.Worksheets("colinfo_").Delete
        Application.DisplayAlerts = True
        On Error GoTo 0

        Set shtColInfo = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        shtColInfo.name = "colinfo_"

        Set colinfo = ExcelSteps.New_ColInfo
        .Assert tst, colinfo.Init(colinfo, files, , ThisWorkbook)
        .Assert tst, Not colinfo.tbl Is Nothing
        .Assert tst, colinfo.tbl.wkbk Is ThisWorkbook
        .Update tst, procs
    End With

    On Error Resume Next
    Application.DisplayAlerts = False
    ThisWorkbook.Worksheets("colinfo_").Delete
    Application.DisplayAlerts = True
    On Error GoTo 0
    
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
        .Assert tst, colinfo.curTbl = "BR_Example"

        'Check rngRowsCurTbl address and n rows
        .Assert tst, Not colinfo.rngRowsCurTbl Is Nothing
        .Assert tst, colinfo.rngRowsCurTbl.Rows.Count = ColInfo_BRTable_nrows
        .Assert tst, colinfo.rngRowsCurTbl.Address = ColInfo_BRTable_rngRows

        'Check that .colrngTblName set to correct column range for "BR_Example"
        .Assert tst, colinfo.tbl.colrngTblName.Address = ColInfo_colRng_BRExample
        .Update tst, procs
    End With
    CloseColInfoWkbk colinfo
    
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
        .Assert tst, dict.Count = ColInfo_dictNormalize_BRExample_Count

        'Check known mapping
        .Assert tst, dict.Item("Location") = ColInfo_Locn_VarnameRaw_BRExample
        .Assert tst, dict.Item("Net_Sales") = ColInfo_Sales_VarnameRaw_BRExample
        .Update tst, procs
    End With
    CloseColInfoWkbk colinfo
    
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
        .Assert tst, colinfo.curTbl = "Second_Tbl"
        nRowsSecond = colinfo.rngRowsCurTbl.Rows.Count
        .Assert tst, nRowsSecond > 0
        .Assert tst, nRowsSecond <> nRowsFirst
        .Update tst, procs
    End With
    CloseColInfoWkbk colinfo
    
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
Sub InitImportParseNormTest(tst, imptbl As Object, colinfo As Object, files As Object, _
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

    Set imptbl = ExcelSteps.New_ImportParseNorm
    tst.Assert tst, imptbl.Init(imptbl, colinfo, files, curTbl, dParamsImport, dParamsParse)
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
Sub CloseImportParseNormWkbk(imptbl As Object)
    If imptbl Is Nothing Then Exit Sub
    If imptbl.tblRaw Is Nothing Then Exit Sub
    If TypeName(imptbl.tblRaw) = "tblRowsCols" Then
        If imptbl.tblRaw.wkbk Is Nothing Then Exit Sub
        imptbl.tblRaw.wkbk.Close False
    End If
End Sub

