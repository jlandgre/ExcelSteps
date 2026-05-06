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
'Version 5/5/26

Public tbl As Object            ' tblRowsCols instance for ColInfo.xlsx
Public CurTbl As String         ' Name of current project table column in colinfo.tbl
Public rngRowsCurTbl As Range   ' Row range for current table's included rows
'-----------------------------------------------------------------------------------------
' Open ColInfo.xlsx and provision colinfo.tbl; optionally call SetCurTbl
' JDL 5/1/26; updated 5/6/26
'
Public Function Init(colinfo, files As Object, Optional curTbl As String) As Boolean
    SetErrs Init: If errs.IsHandle Then On Error GoTo ErrorExit
    Dim wkbkColInfo As Workbook, tblAlias As Object

    If errs.IsFail(Len(files.pfColInfo) = 0, 1) Then GoTo ErrorExit
    If Not OpenFile(files.pfColInfo, wkbkColInfo) Then GoTo ErrorExit
    If errs.IsFail(wkbkColInfo Is Nothing, 2) Then GoTo ErrorExit

    With colinfo
        Set .tbl = New tblRowsCols
        Set tblAlias = .tbl
        If Not tblAlias.Provision(tblAlias, wkbkColInfo, False, sht:=shtColInfo, _
            IsSetColRngs:=True) Then GoTo ErrorExit

        ' If curTbl name specified, sort tbl by its column and set rngRows for its rows
        If Len(curTbl) > 0 Then
            If Not .SetCurTbl(colinfo, curTbl) Then GoTo ErrorExit
        End If
    End With
    Exit Function

ErrorExit:
    errs.RecordErr "ColInfo.Init", Init
End Function
'-----------------------------------------------------------------------------------------
' Sort colinfo.tbl by sTbl column ascending; set rngRowsCurTbl to non-blank rows
' JDL 5/1/26; updated 5/6/26
'
Public Function SetCurTbl(colinfo, curTbl) As Boolean
    SetErrs SetCurTbl: If errs.IsHandle Then On Error GoTo ErrorExit
    Dim rngStart As Range

    With colinfo
        .CurTbl = curTbl

        ' Check that .curTbl name found in table and set colrng for it
        Set .tbl.colrngTblName = .tbl.rngTblHeaderVal(.tbl, .CurTbl).EntireColumn
        If errs.IsFail(.tbl.colrngTblName Is Nothing, 1) Then GoTo ErrorExit

        ' Sort by .curTbl column and set rngRowsCurTbl to its non-blank rows
        If Not .tbl.TblSortBy(.tbl, .CurTbl) Then GoTo ErrorExit
        Set rngStart = Intersect(.tbl.rngRows, .tbl.colrngTblName.EntireColumn).Cells(1, 1)
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
    SetErrs "non-bool": If errs.IsHandle Then On Error GoTo ErrorExit
    Dim ary() As Variant, n As Long, rngRow As Range

    With colinfo

        ' Check that CurTbl is set
        If errs.IsFail(Len(.CurTbl) = 0, 1) Then GoTo ErrorExit

        ' Build array of index variables from sorted tbl
        n = 0
        ReDim ary(0 To .rngRowsCurTbl.Rows.Count - 1)
        For Each rngRow In .rngRowsCurTbl.Rows
            If Intersect(rngRow, .tbl.colrngIsIndex.EntireColumn).Value2 = True Then
                ary(n) = Intersect(rngRow, .tbl.colrngVarNorm.EntireColumn).Value2
                n = n + 1
            End If
        Next rngRow
    End With
    If n > 0 Then ReDim Preserve ary(0 To n - 1) Else ary = Array()
    YieldAryIndices = ary
    Exit Function

ErrorExit:
    errs.RecordErr "ColInfo.YieldAryIndices"
End Function
'-----------------------------------------------------------------------------------------
' Return Variant array of VarNameNorm values for non-index rows in rngRowsCurTbl
' JDL 5/1/26; updated 5/6/26
'
Public Function YieldAryMetrics(colinfo) As Variant
    SetErrs "non-bool": If errs.IsHandle Then On Error GoTo ErrorExit
    Dim ary() As Variant, n As Long, rngRow As Range, valNorm As Variant

    With colinfo

        ' Check that CurTbl is set
        If errs.IsFail(Len(.CurTbl) = 0, 1) Then GoTo ErrorExit

        n = 0
        ReDim ary(0 To .rngRowsCurTbl.Rows.Count - 1)
        For Each rngRow In .rngRowsCurTbl.Rows
            If Intersect(rngRow, .tbl.colrngIsIndex.EntireColumn).Value2 <> True Then
                valNorm = Intersect(rngRow, .tbl.colrngVarNorm.EntireColumn).Value2
                If Len(CStr(valNorm)) > 0 Then
                    ary(n) = valNorm
                    n = n + 1
                End If
            End If
        Next rngRow
    End With
    ReDim Preserve ary(0 To n - 1)
    YieldAryMetrics = ary
    Exit Function

ErrorExit:
    errs.RecordErr "ColInfo.YieldAryMetrics"
End Function
'-----------------------------------------------------------------------------------------
' Return Dictionary of VarNameNorm->VarNameRaw pairs for rngRowsCurTbl
' JDL 5/1/26
'
Public Function YieldDNormalize(colinfo) As Object
    SetErrs "non-bool": If errs.IsHandle Then On Error GoTo ErrorExit
    Dim dict As New Dictionary, rngRow As Range
    Dim valNorm As Variant, valRaw As Variant

    With colinfo

        ' Check that CurTbl is set
        If errs.IsFail(Len(.CurTbl) = 0, 1) Then GoTo ErrorExit
        
        ' Build dictionary of VarNameNorm->VarNameRaw pairs from sorted tbl
        For Each rngRow In .rngRowsCurTbl.Rows
            valNorm = Intersect(rngRow, .tbl.colrngVarNorm.EntireColumn).Value2
            valRaw = Intersect(rngRow, .tbl.colrngVarRaw.EntireColumn).Value2
            If errs.IsFail(Len(valRaw) = 0, 4) Then GoTo ErrorExit

            ' If variable is in normalized table, add pair to dict
            If Len(valNorm) > 0 Then dict.Add CStr(valNorm), CStr(valRaw)
        Next rngRow
    End With
    Set YieldDNormalize = dict
    Exit Function

ErrorExit:
    errs.RecordErr "ColInfo.YieldDNormalize"
End Function
