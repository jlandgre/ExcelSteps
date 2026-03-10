VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "tblRowsCols"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Version 1/29/26
Option Explicit

'Table Locator Attributes
Public wkbk As Workbook
Public sht As String
Public wksht As Worksheet
Public TblName As String
Public Defn As String
Public NamePrefix As String
Public ZoomSetting As Integer
Public IsBlankSht As Boolean
Public IsNoData As Boolean

'True if non-default/custom table; False if default, homed table
Public IsCustomTbl As Boolean

'Table Ranges
Public cellHome As Range
Public rcHome As String 'Comma-separated row,column for top-left data cell
Public rowHome As Integer
Public colHome As Integer
Public rngHeader As Range
Public rngTable As Range
Public rngRows As Range
Public rngRowsPopulated As Range
Public rowCur As Range
Public colCur As Range

'Row and Column Counts and Offset variables
Public lastcol As Long
Public nCols As Long
Public lastrow As Long
Public nRows As Long
Public iOffsetHeader As Integer
Public iOffsetKeyCol As Integer

'Boolean settings
Public IsSetTblNames As Boolean 'table and header ranges named
Public IsSetColNames As Boolean 'individual column ranges programmatically named
Public IsSetAryCols As Boolean 'Create array of column ranges
Public IsSetColRngs As Boolean 'Set named ranges for columns (project-specific)

Public IsNamePrefix As Boolean 'Range names have prefix (either sht or TableName)
Public IsPrefixSht As Boolean 'Name prefix sheet name (TRUE) or table name (FALSE)?

Public IsNameColRows As Boolean '(7/30/25) Name .rngRows col block instead of whole col

'*** Table-specific attributes ***

'shtErrors col ranges
Public colrngCode As Range
Public colrngMod As Range
Public colrngFunc As Range
Public colrngSMsg As Range
Public colrngSVal As Range
Public colrngIMsg As Range

'shtSteps col ranges
Public colrngSht As Range
Public colrngCol As Range
Public colrngStep As Range
Public colrngStrInput As Range
Public colrngCol2 As Range
Public colrngKeep As Range
Public colrngComment As Range
Public colrngNumFmt As Range
Public colrngWidth As Range

'shtTblImp col ranges
Public colrngMdlName As Range
Public colrngGrp As Range
Public colrngSubgrp As Range
Public colrngDesc As Range
Public colrngVarName As Range
Public colrngUnits As Range
Public colrngFormula As Range
Public colrngScenName As Range
Public colrngValue As Range

'*** Project-specific table attributes ***

'ColInfo
Public colrngColName As Range
Public colrngVarDesc As Range
Public colrngVarUnits As Range
Public colrngVarXlwidth As Range
Public colrngVarXlFormat As Range
Public colrngOrder_Errors As Range

'Add project-specific attributes for project's tables
'Trts, Stains, Substrates, Graders, Protocols
Public colrngID As Range
Public colrngName As Range
Public colrngLot As Range
Public colrngStainID As Range
Public colrngSubstID As Range

'Swatches
Public colrngStain As Range
Public colrngSubstrate As Range
Public colrngRep As Range
Public colrngDup As Range
Public colrngProtocol As Range
Public colrngTrt As Range

'Color
Public colrng_L_Cleaned As Range
Public colrng_a_Cleaned As Range
Public colrng_b_Cleaned As Range

'ColorUnstained
Public colrng_LColor As Range
Public colrng_aColor As Range
Public colrng_bColor As Range

'ColorSummary validation
Public colrngGroup As Range
Public colrngCMCDE As Range


'Array of column ranges - requires Get/Let Properties to work (VBA bug/arch. flaw)
Private pAryColRngs() As Variant
Public Property Get aryColRngs() As Variant
    aryColRngs = pAryColRngs
End Property
Public Property Let aryColRngs(vArray As Variant)
    pAryColRngs = vArray
End Property
'-----------------------------------------------------------------------------------------
' Populate Rows/Cols table attributes
'
' To sufficiently specify a table, call with either wkbk + sht args or wkbk + TblName +
' tbl_tablename Setting defined to supply location and other settings. TblName default
' is worksheet name for "homed" (cellHome = Cells(2,1)) tables or can be other
' string for non-homed tables including sheets containing multiple tables
'
'Inputs:
' wkbk: Workbook object for table
' IsFormat: Boolean flag for whether to format the table
' sht: String name of worksheet for table (specify sht for default, homed table)
' TblName: String name of table (specify to read table definition from Settings)
' Defn: String parseable definition of custom table (must be specified if not sht or TblName)
' cellHome: Default Cells(2,1) for homed table. CellHome is data cell rngupper left corner
' iOffsetHeader: Default -1 but can be other (negative) integer to create header range an
'                arbitrary number of rows above cellHome row
' iOffsetKeyCol: Default 0 but can be positive integer to flag a key column in the tbl
' IsSetTblNames: Default True; toggles whether FormatTable sets Excel table and header name
' IsSetColNames: Default True; toggles whether FormatTable sets column range names
' IsNamePrefix: Default False; toggles whether Excel Range names have sheet name as prefix
' IsPrefixSht: Default True; toggles whether name prefix is sheet name or table name
'
' Created:   4/8/21 JDL      Modified: 10/18/24
'
Public Function Provision(ByRef tbl As tblRowsCols, wkbk, IsFormat, _
    Optional sht, Optional TblName, Optional Defn, Optional rcHome, Optional nRows, _
    Optional nCols, Optional iOffsetKeyCol, Optional iOffsetHeader, Optional IsSetAryCols, _
    Optional IsSetColRngs, Optional IsSetTblNames, Optional IsSetColNames, _
    Optional IsNamePrefix, Optional IsPrefixSht, Optional NamePrefix) As Boolean
    
    SetErrs Provision: If errs.IsHandle Then On Error GoTo ErrorExit
    
    With tbl
    
        'Initialize configuration-related attributes
        If Not .Init(tbl, wkbk, sht, TblName, Defn, rcHome, nRows, nCols, _
            iOffsetKeyCol, iOffsetHeader, IsSetAryCols, IsSetColRngs, IsSetTblNames, _
            IsSetColNames, IsNamePrefix, IsPrefixSht, NamePrefix) Then GoTo ErrorExit
                
        'Set dimensions based either on Defn string or actual data
        If Not .SetDimensions(tbl) Then GoTo ErrorExit
        
        'Set array of columnns and range names
        If Not .SetArysNamesRngs(tbl) Then GoTo ErrorExit

        'Generic and project-specific formatting
        If IsFormat Then
            If Not .Format(tbl) Then GoTo ErrorExit
            'if not .Format_proj tbl Then GoTo ErrorExit
        End If
    End With
    Exit Function
    
ErrorExit:
    errs.RecordErr "tblRowsCols.Provision", Provision
End Function
'-----------------------------------------------------------------------------------------
'-----------------------------------------------------------------------------------------
' Init Procedure
'-----------------------------------------------------------------------------------------
'-----------------------------------------------------------------------------------------
' Initialize configuration attributes for custom or default/homed table
' JDL 5/22/24; Modified 9/30/24 add if condition for .wkbk is Nothing
'
Function Init(tbl, wkbk, Optional sht, Optional TblName, Optional Defn, Optional rcHome, _
    Optional nRows, Optional nCols, Optional iOffsetKeyCol, Optional iOffsetHeader, _
    Optional IsSetAryCols, Optional IsSetColRngs, Optional IsSetTblNames, Optional _
    IsSetColNames, Optional IsNamePrefix, Optional IsPrefixSht, Optional NamePrefix) _
    As Boolean
    
    SetErrs Init: If errs.IsHandle Then On Error GoTo ErrorExit
    
    With tbl
        Set .wkbk = wkbk
        
        'Fix error cause (Excel bug) due to case sensitivity of Sheet names
        If Not IsMissing(sht) And Not .wkbk Is Nothing Then
            If IsShtCaseErr(.wkbk, sht) Then errs.ReportWarningMsg 1, "tblRowsCols.Init", param:=sht
        End If
    
        'Set flag and attr for whether custom or default/homed table
        If Not .SetIsCustomTbl(tbl, TblName, Defn, sht) Then GoTo ErrorExit
        
        'Populate attributes that specify table configuration
        If .IsCustomTbl Then
            If Not .SetCustomTblParams(tbl) Then GoTo ErrorExit
        Else
            If Not .SetHomedTblParams(tbl) Then GoTo ErrorExit
        End If
        
        'Override configuration attributes if specified as arguments
        If Not .OverrideWithArgs(tbl, sht, rcHome, nRows, nCols, iOffsetKeyCol, _
            iOffsetHeader, IsSetAryCols, IsSetColRngs, IsSetTblNames, IsSetColNames, _
            IsNamePrefix, IsPrefixSht, NamePrefix) Then GoTo ErrorExit
    End With
    Exit Function
    
ErrorExit:
    errs.RecordErr "tblRowsCols.Init", Init
End Function
'-----------------------------------------------------------------------------------------
' Set flag and other attrs for custom or default/homed table configuration
' JDL 5/29/24; 12/13/24 redo logic to limit/clarify options for custom/default tables
Function SetIsCustomTbl(tbl, Optional TblName, Optional Defn, Optional sht) As Boolean
    SetErrs SetIsCustomTbl: If errs.IsHandle Then On Error GoTo ErrorExit
    Dim IsOverspecified As Boolean
    
    With tbl
        'Custom tbl specified by defn (with optional override sht arg)
        
        If Not IsMissing(Defn) Then
            .IsCustomTbl = True
            .Defn = Defn
            If Not IsMissing(TblName) Then .TblName = TblName
            If Not IsMissing(sht) Then .sht = sht
            
        'Legacy custom table specified by setting
        ElseIf Not IsMissing(TblName) And IsMissing(sht) Then
            .IsCustomTbl = True
            .TblName = TblName
            
        'Default/homed table (with either default or alt tblname specified as arg)
        ElseIf Not IsMissing(sht) Then
            .IsCustomTbl = False
            .sht = sht
            .TblName = .sht
            If Not IsMissing(TblName) Then .TblName = TblName
            
        'Error if neither sht nor one of TblName or Defn specified
        Else
            errs.iCodeLocal = 1
            GoTo ErrorExit
        End If
    End With
    Exit Function
    
ErrorExit:
    errs.RecordErr "SetIsCustomTbl", SetIsCustomTbl
End Function
'-----------------------------------------------------------------------------------------
' Populate custom table parameters from TableName (e.g. Setting) or Defn arg
'
' The function parses a Defn string describing the table. The string either
' comes directly from the Defn argument if specified or from reading a Setting
' based on TblName argument.
'
' Rows/Cols Table Defn string format:
' Setting name: tbl_TblName (based on TblName argument)
' aryTemp(0)  sht - worksheet name where table resides
' aryTemp(1)  r,c - row,column home cell location on sht
' aryTemp(2 to 7)  Booleans - "T" or "F" (IsSetTblNames, IsSetColNames,
'                      IsNamePrefix, IsPrefixSht, IsSetAryCols, IsSetColRngs)
' aryTemp(8)  cellHOme.Colum + iOffsetKeyCol specifies key col (default 0)
' aryTemp(9)  cellHOme.Colum + iOffsetHeader is header row (default -1)
' aryTemp(10) fixed number of table rows ('0' to disable this feature)
' aryTemp(11) fixed number of table columns ('0' to disable this feature)
'
' Example Defn string: Process:6,2:T:T:T:T:T:T:1:-2:12:5
'
' 4/7/21 JDL    Modified 12/13/24 to allow override sht name
'
Function SetCustomTblParams(tbl) As Boolean
    SetErrs SetCustomTblParams: If errs.IsHandle Then On Error GoTo ErrorExit
    Dim i As Integer, aryRaw As Variant, aryTemp As Variant, aryRCHome As Variant

    With tbl
    
        'If Defn not specified, read Defn from Settings_ sheet
        If .IsCustomTbl And (Not Len(.Defn) > 0) Then _
            If Not .ReadDefnSetting(tbl) Then GoTo ErrorExit
        
        'Parse into temp array; init default values/data types of result array
        aryRaw = Split(.Defn, ":")
        aryTemp = Array(False, False, False, False, False, False, 0, -1, 0, 0)
        
        'Boolean Flags
        For i = 2 To 7
            If aryRaw(i) = "T" Then aryTemp(i - 2) = True
        Next i
    
        'Integers
        For i = 8 To 11
            aryTemp(i - 2) = CInt(aryRaw(i))
        Next i

        'Populate class attributes
        If UBound(aryTemp) <> -1 Then
            
            'Set sht name unless overridden by arg (set in SetIsCustomTbl in this case)
            If Len(.sht) < 1 Then .sht = aryRaw(0)
            
            aryRCHome = Split(aryRaw(1), ",")
            .rowHome = CInt(aryRCHome(0))
            .colHome = CInt(aryRCHome(1))
            .IsSetTblNames = aryTemp(0)
            .IsSetColNames = aryTemp(1)
            .IsNamePrefix = aryTemp(2)
            .IsPrefixSht = aryTemp(3)
            .IsSetAryCols = aryTemp(4)
            .IsSetColRngs = aryTemp(5)
            .iOffsetKeyCol = aryTemp(6)
            .iOffsetHeader = aryTemp(7)
            .nRows = aryTemp(8)
            .nCols = aryTemp(9)
            
            'Default table name is sheet name if not specified as argument
            If Len(.TblName) < 1 Then .TblName = .sht
            .NamePrefix = .TblName
        End If
    End With
    Exit Function
ErrorExit:
    errs.RecordErr "SetCustomTblParams", SetCustomTblParams
End Function
'-----------------------------------------------------------------------------------
' Read custom table definition from Settings
' JDL 5/21/24
'
Function ReadDefnSetting(tbl) As Boolean
    SetErrs ReadDefnSetting: If errs.IsHandle Then On Error GoTo ErrorExit
    
    'Set Defn from the Setting value
    tbl.Defn = ReadSetting(tbl.wkbk, "tbl_" & tbl.TblName)
    
    'Exit with error if the Setting was not found
    If errs.IsFail(Len(tbl.Defn) < 1, 1, tbl.TblName) Then GoTo ErrorExit
    Exit Function
    
ErrorExit:
    errs.RecordErr "ReadDefnSetting", ReadDefnSetting
End Function
'-----------------------------------------------------------------------------------
' Assign default attribute values for homed table
' JDL 5/22/24; 10/9/24 add set .TblName; 10/18/24 change .IsSetColNames default True
' 11/5/24 delete setting .TblName (in SetIsCustomTbl method)
'
Function SetHomedTblParams(tbl) As Boolean
    SetErrs SetHomedTblParams: If errs.IsHandle Then On Error GoTo ErrorExit
    Dim aryAtts As Variant, aryVals As Variant, i As Integer
    
    'Defaults attribute settings for homed/default tables
    With tbl
        .rowHome = 2
        .colHome = 1
        .NamePrefix = .sht
        .IsSetTblNames = False
        .IsSetColNames = True
        .IsNamePrefix = False
        .IsPrefixSht = False
        .IsSetAryCols = False
        .IsSetColRngs = False
        .iOffsetKeyCol = 0
        .iOffsetHeader = -1
        .nRows = 0
        .nCols = 0
    End With
    Exit Function
    
ErrorExit:
    errs.RecordErr "SetHomedTblParams", SetHomedTblParams
End Function
'-----------------------------------------------------------------------------------
' Override attribute values if non-default tblInit arguments are specified
' JDL 5/22/24
'
Function OverrideWithArgs(tbl, Optional sht, Optional rcHome, Optional nRows, _
    Optional nCols, Optional iOffsetKeyCol, Optional iOffsetHeader, Optional _
    IsSetAryCols, Optional IsSetColRngs, Optional IsSetTblNames, Optional _
    IsSetColNames, Optional IsNamePrefix, Optional IsPrefixSht, Optional _
    NamePrefix) As Boolean
    
    SetErrs OverrideWithArgs: If errs.IsHandle Then On Error GoTo ErrorExit
    Dim aryRCHome As Variant
    
    With tbl
    
        'Re-initialize .sht, .wksht if override specified
        If .IsCustomTbl And (Not IsMissing(sht)) Then
            .sht = sht
        End If
        
        '"r,c" string specifying home cell location on .sht
        If Not IsMissing(rcHome) Then
            aryRCHome = Split(rcHome, ",")
            .rowHome = CInt(aryRCHome(0))
            .colHome = CInt(aryRCHome(1))
        End If
        
        'Set other  values
        If Not IsMissing(IsSetTblNames) Then .IsSetTblNames = IsSetTblNames
        If Not IsMissing(IsSetColNames) Then .IsSetColNames = IsSetColNames
        If Not IsMissing(IsNamePrefix) Then .IsNamePrefix = IsNamePrefix
        If Not IsMissing(IsPrefixSht) Then .IsPrefixSht = IsPrefixSht
        If Not IsMissing(IsSetAryCols) Then .IsSetAryCols = IsSetAryCols
        If Not IsMissing(IsSetColRngs) Then .IsSetColRngs = IsSetColRngs
        If Not IsMissing(iOffsetKeyCol) Then .iOffsetKeyCol = iOffsetKeyCol
        If Not IsMissing(iOffsetHeader) Then .iOffsetHeader = iOffsetHeader
        If Not IsMissing(NamePrefix) Then .NamePrefix = NamePrefix
        If Not IsMissing(nRows) Then .nRows = nRows
        If Not IsMissing(nCols) Then .nCols = nCols
    End With
    Exit Function
    
ErrorExit:
    errs.RecordErr "OverrideWithArgs", OverrideWithArgs
End Function
'-----------------------------------------------------------------------------------
'-----------------------------------------------------------------------------------
'SetDimensions procedure
'-----------------------------------------------------------------------------------
'-----------------------------------------------------------------------------------
' Procedure to set table dimensions
' JDL 5/24/24
'
Function SetDimensions(tbl) As Boolean
    SetErrs SetDimensions: If errs.IsHandle Then On Error GoTo ErrorExit

    With tbl
        If Not .SetWkshtAndRanges(tbl) Then GoTo ErrorExit
        If Not .SetIsBlankSheet(tbl) Then GoTo ErrorExit
        If Not .SetIsNoData(tbl) Then GoTo ErrorExit
        If Not .SetNRows(tbl) Then GoTo ErrorExit
        If Not .SetNCols(tbl) Then GoTo ErrorExit
        If Not .SetRngTable(tbl) Then GoTo ErrorExit
    End With

    Exit Function
ErrorExit:
    errs.RecordErr "SetDimensions", SetDimensions
End Function
'-----------------------------------------------------------------------------------
' Set .wksht, .cellHome and .rngTable
' JDL 5/23/24
'
Function SetWkshtAndRanges(tbl) As Boolean
    SetErrs SetWkshtAndRanges: If errs.IsHandle Then On Error GoTo ErrorExit
    
    With tbl
    
        'If .sht doesn't exist in the workbook, add it
        If Not SheetExists(.wkbk, .sht) Then AddSheet .wkbk, .sht
    
        'Set attributes
        Set .wksht = .wkbk.Sheets(.sht)
        Set .cellHome = .wksht.Cells(.rowHome, .colHome)
        Set .rngTable = .cellHome.CurrentRegion
        Set .rngRows = Nothing
        Set .rngHeader = Nothing
        .lastrow = 0
        .lastcol = 0
    End With

    Exit Function
ErrorExit:
    errs.RecordErr "SetWkshtAndRanges", SetWkshtAndRanges
End Function
'-----------------------------------------------------------------------------------
' Sets the .IsBlankSht attribute (True if .sht is blank)
' JDL 5/23/24
'
Function SetIsBlankSheet(tbl) As Boolean
    SetErrs SetIsBlankSheet: If errs.IsHandle Then On Error GoTo ErrorExit

    With tbl
        .IsBlankSht = False
        If .rngTable.Cells.Count < 2 Then .IsBlankSht = True
    End With

    Exit Function
ErrorExit:
    errs.RecordErr "SetIsBlankSheet", SetIsBlankSheet
End Function
'-----------------------------------------------------------------------------------
' Set the .IsNoData attribute to flag case where only header is populated
' JDL 5/23/24
'
Function SetIsNoData(tbl) As Boolean
    SetErrs SetIsNoData: If errs.IsHandle Then On Error GoTo ErrorExit

    With tbl
        .IsNoData = True
        If .IsBlankSht Then Exit Function
        
        'Multiple populated cells in first data row
        If WorksheetFunction.CountA(Intersect(.rngTable, _
            .cellHome.EntireRow)) >= 1 Then
            .IsNoData = False
            
        'OR multiple data rows (can be just in one populated column)
        ElseIf .rngTable.Rows.Count > 2 Then
            .IsNoData = False
        End If
    End With

    Exit Function
ErrorExit:
    errs.RecordErr "SetIsNoData", SetIsNoData
End Function
'-----------------------------------------------------------------------------------
' If not already initialized, set .nrows; set .lastrow and .rngRows
' JDL 5/23/24
'
Function SetNRows(tbl) As Boolean
    SetErrs SetNRows: If errs.IsHandle Then On Error GoTo ErrorExit
    Dim rng As Range

    With tbl

        'If data and .nrows hasn't been set by Defn or nrows Init arg
        If (Not .IsNoData) And (.nRows = 0) Then _
            .nRows = .rngTable.Rows(.rngTable.Rows.Count).Row _
                - .cellHome.Row + 1
        
        'Set .rngRows and .lastrow based on rows populated with data
        If .nRows > 0 Then
            Set rng = Range(.cellHome, .cellHome.Offset(.nRows - 1, 0))
            Set .rngRows = rng.EntireRow
            .lastrow = .cellHome.Row + .nRows - 1
        End If
    End With

    Exit Function
ErrorExit:
    errs.RecordErr "SetNRows", SetNRows
End Function
'-----------------------------------------------------------------------------------
' If not already initialized, set .ncols; set .lastcol, and .rngheader
' JDL 5/23/24
'
Function SetNCols(tbl) As Boolean
    SetErrs SetNCols: If errs.IsHandle Then On Error GoTo ErrorExit
    Dim rowHeader As Range, rngCols As Range

    With tbl
    
        'Set .ncols based on tbl extent (if not preset by Defn or ncols arg)
        If (Not .IsBlankSht) And (.nCols = 0) Then _
            .nCols = .rngTable.Columns(.rngTable.Columns.Count).Column _
                - .cellHome.Column + 1

        'Set .lastcol and .rngHeader
        If .nCols > 0 Then
            .lastcol = .cellHome.Column + .nCols - 1
            Set rowHeader = .cellHome.Offset(.iOffsetHeader, 0).EntireRow
            Set rngCols = Range(.cellHome, _
                .cellHome.Offset(0, .nCols - 1)).EntireColumn
            Set .rngHeader = Intersect(rowHeader, rngCols)
        End If
    End With

    Exit Function
ErrorExit:
    errs.RecordErr "SetNCols", SetNCols
End Function
'-----------------------------------------------------------------------------------
' Set .rngTable (useful for clearing or formatting entire table)
' JDL 5/24/24
'
Function SetRngTable(tbl) As Boolean
    SetErrs SetRngTable: If errs.IsHandle Then On Error GoTo ErrorExit

    With tbl
    
        'If sht blank, and nrows/ncols not specified by Defn or Init args
        If (.nRows = 0) And (.nCols = 0) Then
            Set .rngTable = Nothing
        
        'If there are columns due to specified or header sensing
        ElseIf .nRows = 0 Then
            Set .rngTable = .rngHeader
            
        'Rang() encompasses header and data (incl blank space between)
        Else
            Set .rngTable = Range(.rngHeader, _
                Intersect(.rngHeader.EntireColumn, .rngRows))
        End If
    End With

    Exit Function
ErrorExit:
    errs.RecordErr "SetRngTable", SetRngTable
End Function
'-----------------------------------------------------------------------------------------
'-----------------------------------------------------------------------------------------
' SetArysNamesRngs Procedure
'-----------------------------------------------------------------------------------------
'-----------------------------------------------------------------------------------------
' Procedure to set arrays and ranges for the table
' JDL 5/23/24
'
Function SetArysNamesRngs(tbl) As Boolean
    SetErrs SetArysNamesRngs: If errs.IsHandle Then On Error GoTo ErrorExit

    With tbl
        If .IsSetAryCols Then If Not .SetAryColRngs(tbl) Then GoTo ErrorExit
        If .IsSetTblNames Then If Not .SetTblNames(tbl) Then GoTo ErrorExit
        If .IsSetColNames Then If Not .SetAllColNames(tbl) Then GoTo ErrorExit
        
        'Set generic and project-specific column ranges
        If .IsSetColRngs Then
            If Not SetColRanges(tbl) Then GoTo ErrorExit
            'If Not SetColRanges_Proj(tbl) Then GoTo ErrorExit
        End If
    End With
    Exit Function
ErrorExit:
    errs.RecordErr "SetArysNamesRngs", SetArysNamesRngs
End Function
'-----------------------------------------------------------------------------------------
' Set an array of column ranges relative to cellHome
'
'Note: Arrays in Classes is a problematic for VBA. Although this function successfully
'    sets tbl.aryColRngs when provisioning a tbl, use of the the attribute causes
'    a VBA error when code attempts to use it in a called subroutine (e.g
'    not the subroutine or function other than the one that instances the tblClass.
'   5/24/24 added Let/Get Properties to get this more robust; see notes in unit test
'
'    Error message: "Property let procedure not defined and property get procedure
'                    did not return an object"
'
'JDL 5/4/20  Modified 5/24/24 to directly set .aryColRngs
'
Function SetAryColRngs(tbl) As Boolean
    SetErrs SetAryColRngs: If errs.IsHandle Then On Error GoTo ErrorExit
    Dim colRanges() As Variant, i As Integer
    
    With tbl
    
        'Exit if no columns
        If .nCols = 0 Then Exit Function
    
        ReDim colRanges(0 To .nCols - 1)
        For i = 0 To .nCols - 1
            Set colRanges(i) = .cellHome.Offset(0, i).EntireColumn
        Next
        .aryColRngs = colRanges
    End With
    Exit Function
    
ErrorExit:
    errs.RecordErr "SetAryColRngs", SetAryColRngs
End Function
'-----------------------------------------------------------------------------------------
' Set the table and header range names
' JDL 5/23/24; 10/7/24 add data values range; 10/18/24 cleanup
'
Function SetTblNames(tbl) As Boolean
    SetErrs SetTblNames: If errs.IsHandle Then On Error GoTo ErrorExit
    Dim i As Integer, sNameString As String

    With tbl
    
        'Exit if no header (e.g. if .ncols=0)
        If .rngHeader Is Nothing Then Exit Function
        
        i = .cellHome.Column
        
        'Name the overall table's range (entire columns --different than .rngTable)
        sNameString = "='" & .sht & "'!C" & i & ":C" & i + .rngHeader.Columns.Count - 1
        MakeXLName .wkbk, .NamePrefix, sNameString
        
        'Name the header row
        sNameString = "='" & .sht & "'!R" & .rngHeader.Row
        MakeXLName .wkbk, .NamePrefix & "_Header", sNameString
        
        'Name the data values range
        If .rngRows Is Nothing Then Exit Function
        sNameString = "='" & .sht & "'!R" & .cellHome.Row & "C" & .cellHome.Column _
            & ":R" & .lastrow & "C" & .lastcol
        MakeXLName .wkbk, .NamePrefix & "_DataValues", sNameString
    End With
    Exit Function
    
ErrorExit:
    errs.RecordErr "SetTblNames", SetTblNames
End Function
'-----------------------------------------------------------------------------------------
' Names all table columns
' JDL 12/16/22; Modified 7/30/25 add option to name col row block instead of whole col
' Add I
'
Function SetAllColNames(tbl) As Boolean
    SetErrs SetAllColNames: If errs.IsHandle Then On Error GoTo ErrorExit
    Dim r As Variant
    
    With tbl
        'Exit if no header (e.g. if .ncols=0)
        If .rngHeader Is Nothing Then Exit Function
            
        For Each r In .rngHeader
            If Not IsError(r.Value2) Then
                If Len(r.Value2) > 0 Then
                    If Not .IsNameColRows Then
                        If Not .NameColumn(tbl, r) Then GoTo ErrorExit
                    Else
                        If Not .NameColRows(tbl, r) Then GoTo ErrorExit
                    End If
                End If
            End If
        Next r
    End With
    Exit Function
    
ErrorExit:
    errs.RecordErr "SetAllColNames", SetAllColNames
End Function
'-----------------------------------------------------------------------------------------
' Name a column specified by its header cell
' JDL 12/16/22  Modified 5/29/24; 10/18/24 cleanup
'
Function NameColumn(tbl, r) As Boolean
    SetErrs NameColumn: If errs.IsHandle Then On Error GoTo ErrorExit
    Dim prefix As String, NameString As String, rng_name As String
    
    With tbl
        
        'Construct the string for the name (optionally with prefix)
        prefix = ""
        If .IsNamePrefix Then prefix = .NamePrefix & "_"
        rng_name = prefix & xlName(r.value)
        
        'Add the new name
        NameString = "='" & .sht & "'!C" & r.Column
        .wkbk.Names.Add Name:=rng_name, RefersToR1C1:=NameString
    End With
    Exit Function
    
ErrorExit:
    errs.RecordErr "NameColumn", NameColumn
End Function
'-----------------------------------------------------------------------------------------
' Name a block of column rows specified by its header cell (e.g. instead of whole column)
' JDL 7/30/25
'
Function NameColRows(tbl, r) As Boolean
    SetErrs NameColRows: If errs.IsHandle Then On Error GoTo ErrorExit
    Dim prefix As String, NameString As String, rng_name As String, rng As Range
    
    With tbl
        
        'Construct the string for the name (optionally with prefix)
        prefix = ""
        If .IsNamePrefix Then prefix = .NamePrefix & "_"
        rng_name = prefix & xlName(r.value)
        
        'Get intersection of the column and table rows
        Set rng = Intersect(.wksht.Columns(r.Column), .rngRows)
        
        'Add the new name using the intersection range
        NameString = "='" & .sht & "'!" & rng.Address
        .wkbk.Names.Add Name:=rng_name, RefersTo:=NameString
    End With
    Exit Function
    
ErrorExit:
    errs.RecordErr "NameColRows", NameColRows
End Function

'-----------------------------------------------------------------------------------------
'-----------------------------------------------------------------------------------------
' Format Procedure
'-----------------------------------------------------------------------------------------
'-----------------------------------------------------------------------------------------
' Format table and name workbook ranges
' JDL 10/13/20 JDL    Modified: 8/11/25 Modify to save and restore sheet's visible status
'
Function Format(tbl As tblRowsCols) As Boolean
    SetErrs Format: If errs.IsHandle Then On Error GoTo ErrorExit
    Dim sht_Active As String, originalVisible As XlSheetVisibility

    With tbl
        If .rngHeader Is Nothing Then Exit Function

        'Clear previous borders
        SetBorders .rngHeader.EntireColumn, xlNone, True
        SetBorders .rngHeader, xlContinuous, True
        If Not .rngRows Is Nothing Then _
            SetBorders Intersect(.rngHeader.EntireColumn, .rngRows), xlContinuous, True
    
        'Format rngheader
        .RngColWidthFormat tbl
        .rngHeader.Style = "Accent1"
        .rngHeader.WrapText = True

        'Turn off gridlines - save and restore visibility status
        sht_Active = .wkbk.ActiveSheet.Name
        originalVisible = .rngHeader.Parent.Visible
        .rngHeader.Parent.Visible = xlSheetVisible
        .rngHeader.Parent.Activate
        ActiveWindow.DisplayGridlines = False
        '.wkbk.Sheets(sht_Active).Select
        .rngHeader.Parent.Visible = originalVisible
        
        'Autofit row height
        If Not .rngRows Is Nothing Then .rngRows.EntireRow.AutoFit
    End With
    Exit Function
    
ErrorExit:
    errs.RecordErr "tblRowsCols.Format", Format
End Function
'-----------------------------------------------------------------------------------------
' Set column widths
' JDL 12/15/22
'
Sub RngColWidthFormat(tbl)
    Dim rngCol As Variant
    For Each rngCol In tbl.rngHeader.EntireColumn
        With rngCol
            .ColumnWidth = 120
            .AutoFit
            If .ColumnWidth < 120 Then .ColumnWidth = .ColumnWidth + 2
        End With
    Next rngCol
End Sub
'-----------------------------------------------------------------------------------------
'-----------------------------------------------------------------------------------------
' tbl Utility Functions
'-----------------------------------------------------------------------------------------
'-----------------------------------------------------------------------------------------
' Lookup and return value at intersection of table cell row and column range
' Inputs:   rngCell [Range] single row range (usually single cell from key column)
'           rngCol [Range] table column range
'           iShift [Integer] optional shift row value
'
'Created: 6/20 JDL
'
Function TableLoc(rngCell, rngCol, Optional ishift = 0) As Variant
    TableLoc = Intersect(rngCell.Offset(ishift, 0).EntireRow, rngCol)
End Function
'-----------------------------------------------------------------------------------------
' Set a table value at intersection of table cell row and column range
' Inputs:   rngCell [Range] single row range (usually single cell from key column)
'           rngCol [Range] table column range
'           val [variant] value to set at intersection of rngCell.row and rngCol
'           iShift [Integer] optional shift row value
'
'Created: 6/20 JDL
'
Sub SetTableLoc(rngCell, rngCol, val, Optional ishift = 0)
    Intersect(rngCell.Offset(ishift, 0).EntireRow, rngCol) = val
End Sub
'-----------------------------------------------------------------------------------------
'Return cell range for specified table header string (Used by a utility related to colinfo
'
'Inputs: sVal [String] header string
'
' JDL 4/14/21 Modified: 1/27/22 FindInRange instead of .Find
'                       6/9/23 JDL cleanup
'
Function rngTblHeaderVal(tbl, sVal) As Range
    Set rngTblHeaderVal = FindInRange(tbl.rngHeader, sVal)
End Function
'-----------------------------------------------------------------------------------------
' Use colinfo lookup to set a table's header comments
'
' Inputs:   tbl [Class instance] Description of table being formatted
'
' JDL 5/24/21          Modified: 11/6/24 switch to tbl.TableLoc method
'
Function SetTableComments(tbl) As Boolean
    SetErrs SetTableComments: If errs.IsHandle Then On Error GoTo ErrorExit
    Dim tblCI As New tblRowsCols, cellCIVar As Variant, r As Variant, sComment As String

    With tblCI
        Call .Provision(tblCI, tbl.wkbk, False, sht:=shtColInfo)

        'Turn off autofilter to avoid search issues
        If .wkbk.Sheets(.sht).FilterMode Then .wkbk.Sheets(.sht).AutoFilter.Range.AutoFilter

        'Loop through headers of table being formatted
        For Each r In tbl.rngHeader
        
            'Find the variable in colinfo table
            Set cellCIVar = FindInRange(.colrngColName, r)
            If Not cellCIVar Is Nothing Then
                sComment = .TableLoc(cellCIVar, .colrngVarDesc)
                If Len(sComment) > 0 Then AddComment r, sComment
            End If
        Next r
    End With
    Exit Function
    
ErrorExit:
    errs.RecordErr "SetTableComments", SetTableComments
End Function
'-----------------------------------------------------------------------------------------
' Use colinfo lookup to set a table's column widths and number formats
' JDL 10/29/20  Modified: 11/6/24 switch to tbl.TableLoc method
'
Function SetTableWidthAndNumFormat(tbl) As Boolean
    SetErrs SetTableWidthAndNumFormat: If errs.IsHandle Then On Error GoTo ErrorExit
    Dim tblCI As New tblRowsCols
    Dim cell_CI As Variant, v As Variant, r As Variant, iWidth As Integer, sFormat As String

    With tblCI
        Call .Provision(tblCI, tbl.wkbk, False, sht:=shtColInfo)

        'Turn off autofilter to avoid search issues
        If .wkbk.Sheets(.sht).FilterMode Then .wkbk.Sheets(.sht).AutoFilter.Range.AutoFilter

        'Loop through column headers of table being formatted
        For Each r In tbl.rngHeader
        
            'Find the variable in colinfo table
            Set cell_CI = .colrngColName.Find(r.value, lookat:=xlWhole)
            If Not cell_CI Is Nothing Then
                iWidth = .TableLoc(cell_CI, .colrngVarXlwidth)
                sFormat = .TableLoc(cell_CI, .colrngVarXlFormat)

                'Locate the variable in header of table being formatted and apply col info
                Set v = tbl.rngHeader.Find(r.value, lookat:=xlWhole)
                If iWidth > 0 Then v.EntireColumn.ColumnWidth = iWidth
                If Len(sFormat) > 0 And Not tbl.rngRows Is Nothing Then _
                    Intersect(v.EntireColumn, tbl.rngRows).NumberFormat = sFormat
            End If
        Next r
    End With
    Exit Function
    
ErrorExit:
    errs.RecordErr "SetTableWidthAndNumFormat", SetTableWidthAndNumFormat
End Function
'-----------------------------------------------------------------------------------------
' Set column groupings for arrays of start and end column pairs
'
Public Sub TblGroupColumns(tbl, ByVal ary1, ByVal ary2)
    Dim i As Integer
    With tbl.wkbk.Sheets(tbl.sht)
        .UsedRange.EntireColumn.ClearOutline
        .Outline.SummaryColumn = xlLeft
    End With

    For i = 0 To UBound(ary1)
        Range(ary1(i), ary2(i)).Columns.Group
    Next i
End Sub

'-----------------------------------------------------------------------------------------
'Select table's home cell; retain original active sheet
'
' Created:   1/13/22 JDL
'
Sub SelectHomeCell(tbl)
    Dim wkshtA As Worksheet
    With tbl
        Set wkshtA = .wkbk.ActiveSheet
        .wkbk.Sheets(.sht).Activate
        .cellHome.Select
        wkshtA.Activate
    End With
End Sub
'-----------------------------------------------------------------------------------------
'Freeze tbl.sht Row1
'
'Modified:   3/1/23 JDL
'
Function FreezeRow1(tbl) As Boolean
    SetErrs FreezeRow1: If errs.IsHandle Then On Error GoTo ErrorExit
    Dim wkshtA As Worksheet
    
    Set wkshtA = tbl.wkbk.ActiveSheet

    'Activate to allow direct call
    tbl.wksht.Activate
    With ActiveWindow
        .FreezePanes = False
        .SplitRow = 1
        .FreezePanes = True
    End With
    wkshtA.Activate
    Exit Function
    
ErrorExit:
    errs.RecordErr "FreezeRow1", FreezeRow1
End Function
'-----------------------------------------------------------------------------------------
' Activate tbl's sheet and save its zoom setting for later
' JDL 12/19/22
Sub ShtActivateAndSaveZoom(tbl)
        tbl.wksht.Activate
        tbl.ZoomSetting = ActiveWindow.Zoom
End Sub
'-----------------------------------------------------------------------------------------
' Restore previous zoom setting
' JDL 12/19/22
Sub RestoreZoom(tbl)
    ActiveWindow.Zoom = tbl.ZoomSetting
End Sub
'-----------------------------------------------------------------------------------------
' Format tblImport sheet rows/cols table
'
' Created:   12/15/21 JDL  Modified: 7/14/23 Hard code column headers if use ColInfo
'
Function FormatMdlImport(tbl) As Boolean
    SetErrs FormatMdlImport: If errs.IsHandle Then On Error GoTo ErrorExit
    Dim rng As Range, i As Integer, w As Variant, c As Range, s As String
    
    With tbl
    
        'Refresh column header strings
        s = "Model,Grp,Subgrp,Description,Variable Name,Units,Number Fmt,Formula/Row Type,Scenario Name,Value"
        If Not .RefreshHeaderVals(tbl, s, sStyle:="Accent1", IsWrapText:=True) Then GoTo ErrorExit
        .colrngStrInput.NumberFormat = "@"
        .colrngNumFmt.NumberFormat = "@"
        
        'Ensure formula cells display formulas as text - not calculated values
        If Not .rngRows Is Nothing Then
            For Each w In .rngRows
                Set c = Intersect(w, .colrngStrInput)
                c.value = c.Formula
            Next w
        End If
                
        'Set column widths and freeze row 1
        i = 1
        For Each w In Array(18, 18, 18, 30, 18, 18, 18, 18, 18, 18)
            .wkbk.Sheets(.sht).Columns(i).EntireColumn.ColumnWidth = w
            i = i + 1
        Next w
        If Not .FreezeRow1(tbl) Then GoTo ErrorExit
    End With
    Exit Function
    
ErrorExit:
    errs.RecordErr "FormatMdlImport", FormatMdlImport
End Function
'-----------------------------------------------------------------------------------------
'Clear table cells
'JDL 2/28/23 Modified 4/6/23 check rngTable is Nothing
Function ClearTable(tbl) As Boolean
    SetErrs ClearTable: If errs.IsHandle Then On Error GoTo ErrorExit

    If tbl.rngTable Is Nothing Then Exit Function
    
    'Clear sheet outline and table contents
    tbl.wkbk.Sheets(tbl.sht).UsedRange.EntireColumn.ClearOutline
    With tbl.rngTable
        .ClearContents
        .ClearFormats
        .ClearComments
        'xxx 2/28/23 should clear cell validation rules
    End With
    Exit Function
    
ErrorExit:
    errs.RecordErr "ClearTable", ClearTable
End Function
'-----------------------------------------------------------------------------------------
' Write/refresh the header for a table and set its style (default based on colinfo table)
'
' sStyle [String] an Excel named style such as "Accent1"
' IsWrapText [Boolean] TRUE to format header as wrapped text
'
' Created:   12/17/20 JDL    Modified 4/6/23 Add SetHeaderFromLstCols
'                                     6/9/23 JDL cleanup
'
'Built/validated in Unique Vals_1220.xlsm
'
Function RefreshHeaderVals(tbl, Optional lstcols, Optional sStyle, Optional IsWrapText) As Boolean
    SetErrs RefreshHeaderVals: If errs.IsHandle Then On Error GoTo ErrorExit
    Dim lst As String
    
    With tbl
    
        'Clear outline and read comma-separated header string from Settings
        .wkbk.Sheets(.sht).UsedRange.EntireColumn.ClearOutline
        If IsMissing(lstcols) Then lstcols = BuildHeaderListFromColInfo(.wkbk, .TblName)

        'set/reset ncols and rngheader based on lst
        SetRngHeaderFromColsLst tbl, lstcols
        
        .rngHeader = Split(lstcols, ",")
        If Not IsMissing(sStyle) Then .rngHeader.Style = sStyle
        If Not IsMissing(IsWrapText) Then .rngHeader.WrapText = IsWrapText
    End With
    Exit Function
    
ErrorExit:
    errs.RecordErr "RefreshHeaderVals", RefreshHeaderVals
End Function
'-----------------------------------------------------------------------------------------
'set ncols and rngHeader based on a comma-separated list of columns
'JDL 4/6/23
Sub SetRngHeaderFromColsLst(tbl, lst)
    Dim r As Range
    With tbl
        If .rngHeader Is Nothing Then
            .nCols = UBound(Split(lst, ",")) + 1
            Set r = .cellHome.Offset(.iOffsetHeader, 0)
            Set .rngHeader = Range(r, r.Offset(0, .nCols - 1))
        End If
    End With
End Sub
'-----------------------------------------------------------------------------------------
'Function SortRowRange - Sort a row range
'Moved to tblRowsCols 3/3/22
'
Function SortRowRange(tbl, sSortBy) As Boolean
    Dim r As Variant
    SortRowRange = True

    'Locate the header for the specified Sort-by column
    Set r = FindInRange(tbl.rngHeader, sSortBy)
    
    With tbl.rngHeader.Parent.Sort.SortFields
        .Clear
        .Add key:=r, SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    End With
    With tbl.rngHeader.Parent.Sort
        .SetRange Intersect(tbl.rngHeader.EntireColumn, tbl.rngRows)
        .Header = xlNo
        .MatchCase = False
        .Orientation = xlTopToBottom
        .Apply
        .SortFields.Clear
    End With
End Function
'------------------------------------------------------------------------------------
' Performs multi-key Sort on tbl
'
' Modified 4/30/24; 11/6/24 moved to tblRowsCols class
'
Public Function TblSortBy(tbl, SortBy) As Boolean
    SetErrs TblSortBy: If errs.IsHandle Then On Error GoTo ErrorExit
    Dim r As Range, RSortBy As Range, arySortBy As Variant, i As Integer
        
    If tbl.rngRows Is Nothing Then Exit Function
    
    'Exit sort-by not specified
    If Len(SortBy) < 1 Then GoTo ErrorExit
    
    'Split SortBy string into array of column names
    arySortBy = Split(SortBy, ",")

    'Add each sort column to SortFields; Exit if specified column not found in the table
    With tbl.wksht.Sort.SortFields
        .Clear
        For i = LBound(arySortBy) To UBound(arySortBy)
            Set RSortBy = tbl.rngHeader.Find(arySortBy(i), lookat:=xlWhole)
            If RSortBy Is Nothing Then GoTo ErrorExit
            .Add key:=RSortBy, SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
        Next i
    End With
    
    'Sort the table by the SortFields
    With tbl.rngHeader.Parent.Sort
        .SetRange Intersect(tbl.rngHeader.EntireColumn, tbl.rngRows)
        .Header = xlNo
        .MatchCase = False
        .Orientation = xlTopToBottom
        .Apply
        .SortFields.Clear
    End With
    Exit Function
    
ErrorExit:
    errs.RecordErr "TblSortBy", TblSortBy
End Function
'-----------------------------------------------------------------------------------------
'Below subs and functions are project-specific
'-----------------------------------------------------------------------------------------
' Set custom-named individual column ranges
' JDL Modified 1/27/22 - add plt_idx; 3/1/23 remove tblProcess mdl attributes to separate sub
'              12/13/24 fix typo that causes test failure
'
Public Function SetColRanges(ByRef tbl) As Boolean
    SetErrs SetColRanges: If errs.IsHandle Then On Error GoTo ErrorExit

    With tbl
        If .TblName = shtErrors Then
            Set .colrngCode = .cellHome.Offset(0, 0).EntireColumn
            Set .colrngMod = cellHome.Offset(0, 1).EntireColumn
            Set .colrngFunc = cellHome.Offset(0, 2).EntireColumn
            Set .colrngSMsg = cellHome.Offset(0, 3).EntireColumn
            Set .colrngSVal = cellHome.Offset(0, 4).EntireColumn
            Set .colrngIMsg = cellHome.Offset(0, 5).EntireColumn

        ElseIf .TblName = shtSteps Then
            Set .colrngSht = .cellHome.Offset(0, 0).EntireColumn
            Set .colrngCol = .cellHome.Offset(0, 1).EntireColumn
            Set .colrngStep = .cellHome.Offset(0, 2).EntireColumn
            Set .colrngStrInput = .cellHome.Offset(0, 3).EntireColumn
            Set .colrngCol2 = .cellHome.Offset(0, 4).EntireColumn
            Set .colrngKeep = .cellHome.Offset(0, 5).EntireColumn
            Set .colrngComment = .cellHome.Offset(0, 6).EntireColumn
            Set .colrngNumFmt = .cellHome.Offset(0, 7).EntireColumn
            Set .colrngWidth = .cellHome.Offset(0, 8).EntireColumn

        ElseIf .TblName = shtTblImp Then
            Set .colrngMdlName = .cellHome.Offset(0, 0).EntireColumn
            Set .colrngGrp = .cellHome.Offset(0, 1).EntireColumn
            Set .colrngSubgrp = .cellHome.Offset(0, 2).EntireColumn
            Set .colrngDesc = .cellHome.Offset(0, 3).EntireColumn
            Set .colrngVarName = .cellHome.Offset(0, 4).EntireColumn
            Set .colrngUnits = .cellHome.Offset(0, 5).EntireColumn
            Set .colrngNumFmt = .cellHome.Offset(0, 6).EntireColumn
            Set .colrngStrInput = .cellHome.Offset(0, 7).EntireColumn
            Set .colrngScenName = .cellHome.Offset(0, 8).EntireColumn
            Set .colrngValue = .cellHome.Offset(0, 9).EntireColumn
        End If
    End With
    Exit Function
    
ErrorExit:
    errs.RecordErr "SetColRanges", SetColRanges
End Function






