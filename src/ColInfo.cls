VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ColInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'Version 5/1/26

Public tbl As Object            ' tblRowsCols instance for ColInfo.xlsx
Public CurTbl As String         ' Name of current project table column in colinfo.tbl
Public rngRowsCurTbl As Range   ' Row range for current table's included rows
'-----------------------------------------------------------------------------------------
' Open ColInfo.xlsx and provision colinfo.tbl; optionally call SetCurTbl
' JDL 5/1/26
'
Public Function Init(colinfo, files As Object, Optional curTbl As String) As Boolean
    SetErrs Init: If errs.IsHandle Then On Error GoTo ErrorExit
    Dim wkbkColInfo As Workbook

    If errs.IsFail(Len(files.pfColInfo) = 0, 1) Then GoTo ErrorExit
    If Not OpenFile(files.pfColInfo, wkbkColInfo) Then GoTo ErrorExit
    If errs.IsFail(wkbkColInfo Is Nothing, 2) Then GoTo ErrorExit

    With colinfo
        Set .tbl = New tblRowsCols
        Dim tblAlias As Object: Set tblAlias = .tbl
        If Not tblAlias.Init(tblAlias, wkbkColInfo, "colinfo_") Then GoTo ErrorExit
        If Len(curTbl) > 0 Then
            If Not SetCurTbl(colinfo, curTbl) Then GoTo ErrorExit
        End If
    End With
    Exit Function

ErrorExit:
    errs.RecordErr "ColInfo.Init", Init
End Function
'-----------------------------------------------------------------------------------------
' Sort colinfo.tbl by sTbl column ascending; set rngRowsCurTbl to non-blank rows
' JDL 5/1/26
'
Public Function SetCurTbl(colinfo, sTbl As String) As Boolean
    SetErrs SetCurTbl: If errs.IsHandle Then On Error GoTo ErrorExit
    Dim rngColCurTbl As Range, rngStart As Range

    With colinfo
        .CurTbl = sTbl
        If Not .tbl.TblSortBy(.tbl, sTbl) Then GoTo ErrorExit
        Set rngColCurTbl = .tbl.rngTblHeaderVal(.tbl, sTbl)
        If errs.IsFail(rngColCurTbl Is Nothing, 1) Then GoTo ErrorExit
        Set rngStart = Intersect(.tbl.rngRows, rngColCurTbl.EntireColumn).Cells(1, 1)
        Set .rngRowsCurTbl = rngToExtent(rngStart, True)
    End With
    Exit Function

ErrorExit:
    errs.RecordErr "ColInfo.SetCurTbl", SetCurTbl
End Function
'-----------------------------------------------------------------------------------------
' Return Variant array of VarNameNorm values for index rows in rngRowsCurTbl
' JDL 5/1/26
'
Public Function YieldAryIndices(colinfo) As Variant
    SetErrs YieldAryIndices: If errs.IsHandle Then On Error GoTo ErrorExit
    Dim ary() As Variant, n As Long, rngRow As Range
    Dim rngVarNameNorm As Range, rngIsIndex As Range, cellIsIndex As Range

    With colinfo
        If errs.IsFail(Len(.CurTbl) = 0, 1) Then GoTo ErrorExit
        Set rngVarNameNorm = .tbl.rngTblHeaderVal(.tbl, "VarNameNorm")
        Set rngIsIndex = .tbl.rngTblHeaderVal(.tbl, "IsIndex")
        If errs.IsFail(rngVarNameNorm Is Nothing, 2) Then GoTo ErrorExit
        If errs.IsFail(rngIsIndex Is Nothing, 3) Then GoTo ErrorExit
        n = 0
        ReDim ary(0 To .rngRowsCurTbl.Rows.Count - 1)
        For Each rngRow In .rngRowsCurTbl.Rows
            Set cellIsIndex = Intersect(rngRow, rngIsIndex.EntireColumn)
            If cellIsIndex.Value = True Then
                ary(n) = Intersect(rngRow, rngVarNameNorm.EntireColumn).Value
                n = n + 1
            End If
        Next rngRow
    End With
    ReDim Preserve ary(0 To n - 1)
    YieldAryIndices = ary
    Exit Function

ErrorExit:
    errs.RecordErr "ColInfo.YieldAryIndices", YieldAryIndices
End Function
'-----------------------------------------------------------------------------------------
' Return Variant array of VarNameNorm values for non-index rows in rngRowsCurTbl
' JDL 5/1/26
'
Public Function YieldAryMetrics(colinfo) As Variant
    SetErrs YieldAryMetrics: If errs.IsHandle Then On Error GoTo ErrorExit
    Dim ary() As Variant, n As Long, rngRow As Range
    Dim rngVarNameNorm As Range, rngIsIndex As Range, cellIsIndex As Range

    With colinfo
        If errs.IsFail(Len(.CurTbl) = 0, 1) Then GoTo ErrorExit
        Set rngVarNameNorm = .tbl.rngTblHeaderVal(.tbl, "VarNameNorm")
        Set rngIsIndex = .tbl.rngTblHeaderVal(.tbl, "IsIndex")
        If errs.IsFail(rngVarNameNorm Is Nothing, 2) Then GoTo ErrorExit
        If errs.IsFail(rngIsIndex Is Nothing, 3) Then GoTo ErrorExit
        n = 0
        ReDim ary(0 To .rngRowsCurTbl.Rows.Count - 1)
        For Each rngRow In .rngRowsCurTbl.Rows
            Set cellIsIndex = Intersect(rngRow, rngIsIndex.EntireColumn)
            If cellIsIndex.Value <> True Then
                ary(n) = Intersect(rngRow, rngVarNameNorm.EntireColumn).Value
                n = n + 1
            End If
        Next rngRow
    End With
    ReDim Preserve ary(0 To n - 1)
    YieldAryMetrics = ary
    Exit Function

ErrorExit:
    errs.RecordErr "ColInfo.YieldAryMetrics", YieldAryMetrics
End Function
'-----------------------------------------------------------------------------------------
' Return Dictionary of VarNameNorm->VarNameRaw pairs for rngRowsCurTbl
' JDL 5/1/26
'
Public Function YieldDNormalize(colinfo) As Object
    SetErrs YieldDNormalize: If errs.IsHandle Then On Error GoTo ErrorExit
    Dim dict As New Dictionary, rngRow As Range
    Dim rngVarNameNorm As Range, rngVarNameRaw As Range
    Dim cellNorm As Range, cellRaw As Range

    With colinfo
        If errs.IsFail(Len(.CurTbl) = 0, 1) Then GoTo ErrorExit
        Set rngVarNameNorm = .tbl.rngTblHeaderVal(.tbl, "VarNameNorm")
        Set rngVarNameRaw = .tbl.rngTblHeaderVal(.tbl, "VarNameRaw")
        If errs.IsFail(rngVarNameNorm Is Nothing, 2) Then GoTo ErrorExit
        If errs.IsFail(rngVarNameRaw Is Nothing, 3) Then GoTo ErrorExit
        For Each rngRow In .rngRowsCurTbl.Rows
            Set cellNorm = Intersect(rngRow, rngVarNameNorm.EntireColumn)
            Set cellRaw = Intersect(rngRow, rngVarNameRaw.EntireColumn)
            If errs.IsFail(Len(cellRaw.Value) = 0, 4) Then GoTo ErrorExit
            dict.Add CStr(cellNorm.Value), CStr(cellRaw.Value)
        Next rngRow
    End With
    Set YieldDNormalize = dict
    Exit Function

ErrorExit:
    errs.RecordErr "ColInfo.YieldDNormalize", YieldDNormalize
End Function
