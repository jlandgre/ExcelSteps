Attribute VB_Name = "tests_ToolboxClasses"
Option Explicit
'Version 5/1/26
'-----------------------------------------------------------------------------------------
' Validate ProjFiles and ColInfo classes
' JDL 4/29/26
'
Sub TestingDriver_ToolBox()
    Dim procs As New Procedures, AllEnabled As Boolean
    With procs
        .Init procs, ThisWorkbook, "ToolBox", "Tests_ToolboxClasses"
        SetApplEnvir False, False, xlCalculationAutomatic

        AllEnabled = False
        .ProjFiles.Enabled = True
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
' Helper: instance and init files and colinfo using test_data/ColInfo.xlsx
'
Sub InitColInfoTest(tst, colinfo As Object, files As Object, _
                   Optional curTbl As String)
    Dim wkbkStepsAddin As Workbook
    Set wkbkStepsAddin = GetWorkbookByVBProjectName(vba_project_name)
    ExcelSteps.IsTest = True
    Set files = ExcelSteps.New_ProjFiles
    tst.Assert tst, files.Init(files, wkbkStepsAddin, "test_data")
    Set colinfo = ExcelSteps.New_ColInfo
    If Len(curTbl) > 0 Then
        tst.Assert tst, colinfo.Init(colinfo, files, curTbl)
    Else
        tst.Assert tst, colinfo.Init(colinfo, files)
    End If
End Sub
'-----------------------------------------------------------------------------------------
' Open ColInfo.xlsx and provision colinfo.tbl
' JDL 5/1/26
'
Sub test_ColInfo_Init(procs)
    Dim tst As New Test: tst.Init tst, "test_ColInfo_Init", ThisWorkbook
    Dim colinfo As Object, files As Object

    With tst
        'Check Init succeeds and tbl provisioned
        InitColInfoTest tst, colinfo, files
        .Assert tst, Not colinfo.tbl Is Nothing
        .Assert tst, colinfo.tbl.sht = "colinfo_"
        .Assert tst, Not colinfo.tbl.wkbk Is Nothing

        'Check CurTbl empty when curTbl arg not passed
        .Assert tst, colinfo.CurTbl = ""
        .Assert tst, colinfo.rngRowsCurTbl Is Nothing
        .Update tst, procs
    End With
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

        'Check rngRowsCurTbl non-Nothing and has rows
        .Assert tst, Not colinfo.rngRowsCurTbl Is Nothing
        .Assert tst, colinfo.rngRowsCurTbl.Rows.Count > 0
        .Update tst, procs
    End With
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
        .Assert tst, ary(0) = "Location"
        .Update tst, procs
    End With
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

        'Check first metric matches expected (Net_Sales is first non-index in BR_Example)
        .Assert tst, ary(0) = "Net_Sales"
        .Update tst, procs
    End With
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

        'Check dict returned with entries
        Set dict = colinfo.YieldDNormalize(colinfo)
        .Assert tst, Not dict Is Nothing
        .Assert tst, dict.Count > 0

        'Check known mapping
        .Assert tst, dict.Item("Location") = "Locn_Raw"
        .Assert tst, dict.Item("Net_Sales") = "Sales_Raw"
        .Update tst, procs
    End With
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
    ExcelSteps.IsTest = False
End Sub