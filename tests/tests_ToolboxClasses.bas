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

    procs.EvalOverall procs
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
' Helper: close colinfo workbook opened by Init/Provision if present
'
Sub CloseColInfoWkbk(colinfo As Object)
    If colinfo Is Nothing Then Exit Sub
    If colinfo.tbl Is Nothing Then Exit Sub
    If colinfo.tbl.wkbk Is Nothing Then Exit Sub
    colinfo.tbl.wkbk.Close False
End Sub