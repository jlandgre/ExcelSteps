'mdlScenario cls.vb
'Version 10/29/25
Option Explicit
Public cellHome As Range
Public rngRows As Range
Public rngHeader As Range
Public wkbk As Workbook
Public sht As String
Public wksht As Worksheet
Public IsCalc As Boolean
Public IsSuppHeader As Boolean
Public IsLiteModel As Boolean
Public IsMdlNmPrefix As Boolean
Public IsRngNames As Boolean
Public mdlName As String
Public nRows As Integer
Public NamePrefix As String
Public rngPopRows As Range
Public rngFormulaRows As Range
Public rngPopCols As Range
Public rngStepsVars As Range
Public icolCalc As Integer
Public rngMdl As Range
Public rowCur As Range
Public colCur As Range

'Column Ranges
Public colrngHeader As Range 'Multi-column
Public colrngHeaderFmt As Range 'Multi-column
Public colrngModel As Range 'Multi-column if not .IsCalc
Public colrngFirstScenario As Range 'First scenario column, defaults to colrngModel.columns(1)
Public colrngGrp As Range
Public colrngSubgrp As Range
Public colrngDesc As Range
Public colrngVarNames As Range
Public colrngUnits As Range
Public colrngNumFmt As Range
Public colrngFormulas As Range

'ExcelSteps table and dictionaries for performance
Public tblSteps As tblRowsCols
Public dStepsFormulas As Object
Public dStepsNumFormats As Object
Public dStepsColWidths As Object
Public dStepsComments As Object
Public dDropdownList As Object

'Range naming reference prefix for rows
Public NameRefPrefix As String
Public NameRefPrefixCell As String

'Formatting
Public iColorGrpRows As Long 'RGB color to highlight group rows
'---------------------------------------------------------------------------------------
' Initialize Scenario Model location within workbook and set params from arguments
'
' JDL 7/19/23; Updated 11/11/24 delete redundant .Init call of .SetCellHome method
'               2/15/25 update comments
' 1. sht specified; mdlName or Defn not --default model
' 2. sht, mdlName specified; Defn not --custom model; lookup defn from Settings
' 3. defn specified; sht, mdlName not --custom model; parse sht from Defn
' 4. defn and sht/mdlName specified; sht/mdlName args override parsed from defn
' 5. if mdlName not specified and not in Defn, mdlName=xlName(sht)
'
Public Function Init(ByRef mdl As mdlScenario, wkbk, Optional sht, Optional IsCalc, _
        Optional IsSuppHeader, Optional IsRngNames, Optional cellHome, _
        Optional mdlName, Optional Defn, Optional nRows, Optional IsMdlNmPrefix, _
        Optional IsLiteModel) As Boolean
    SetErrs Init: If errs.IsHandle Then On Error GoTo ErrorExit
    
    With mdl
        Set .wkbk = wkbk
        
        'If sht specified as argument, the arg overrides potentialparsed sht name
        If Not IsMissing(sht) Then .sht = sht
        
        'If mdlName specified as argument, the arg overrides parsed Defn mdlName
        If Not IsMissing(mdlName) Then .mdlName = mdlName
                            
        'Init model using Defn or use mdlName to look up definition from setting
        If (Not IsMissing(mdlName)) Or (Not IsMissing(Defn)) Then
            If Not .ParseMdlScenDefn(mdl, Defn) Then GoTo ErrorExit
            
        'Init model from args (covers case of default model specified only by sht)
        Else
            .mdlName = xlName(.sht)
            If Not .SetAttsFromArgs(mdl, IsLiteModel, IsSuppHeader, IsRngNames, _
                    IsCalc, IsMdlNmPrefix, nRows, cellHome) Then GoTo ErrorExit
        End If
    End With
    Exit Function
    
ErrorExit:
    errs.RecordErr "mdlScenario.Init", Init
End Function
'---------------------------------------------------------------------------------------
' Set CellHome range
'
' JDL 7/17/23
'
Public Function SetCellHome(mdl, cellHome) As Boolean
    SetErrs SetCellHome: If errs.IsHandle Then On Error GoTo ErrorExit
    With mdl
    
        'CellHome specified from Init argument
        If Not IsMissing(cellHome) Then
            Set .cellHome = cellHome
            Exit Function
        End If
        
        'Default CellHome
        Set .cellHome = .wksht.Cells(2, 1)
        If .IsSuppHeader Then Set .cellHome = .wksht.Cells(1, 1)
    End With
    Exit Function
    
ErrorExit:
    errs.RecordErr "SetCellHome", SetCellHome
End Function
'---------------------------------------------------------------------------------------
' Populate Scenario Model properties and ranges
' Inputs: mdl [mdlScenario Class instance] Provision returns this populated
'         sht [String] sheet name with model (either sht or mdlName required)
'         wkbk [Workbook] workbook object containing the model
'         IsCalc [Boolean] True if single-column model (cells named instead of rows)
'         IsSuppHeader [Boolean] True to suppress writing a header row
'         IsRngNames [Boolean] True to create row and column range names
'         cellHome [Range] Home cell range just below header row (top left corner)
'         mdlName [String] Model name - for reading config from Settings
'         nRows [Integer] Restrict model to fixed number of rows
'         IsMdlNmPrefix [Boolean] Add model name (sheet name default) to range names
'         IsLiteModel [Boolean] True if compact header columns + ExcelSteps
'                     for variable metadata
'
' Created:   1/11/21 JDL Modified: 11/11/24; 8/8/25 add .colrngFirstScenario
'              10/29/25 use tblSteps attribute; add atts for name prefix strings
'
'
Public Function Provision(ByRef mdl As mdlScenario, wkbk, Optional sht, _
            Optional IsCalc, Optional IsSuppHeader, _
            Optional IsRngNames, Optional cellHome, Optional mdlName, _
            Optional Defn, Optional nRows, Optional IsMdlNmPrefix, _
            Optional IsLiteModel) As Boolean
    SetErrs Provision: If errs.IsHandle Then On Error GoTo ErrorExit
    Dim rng As Range
    Dim tblStepsTemp As New tblRowsCols 'For passing instance as ByRef argument to called function

    With mdl
    
        If Not Init(mdl, wkbk, sht, IsCalc, IsSuppHeader, IsRngNames, cellHome, _
            mdlName, Defn, nRows, IsMdlNmPrefix, IsLiteModel) Then GoTo ErrorExit
                
        'Set header column ranges - Full and "Lite" versions
        If Not SetColRanges(mdl) Then GoTo ErrorExit
        
        'Set model starting column
        Set .colrngModel = .colrngUnits.Offset(0, 2)
        If Not .IsLiteModel Then Set .colrngModel = .colrngFormulas.Offset(0, 2)
        
        'Set range for rows containing variables
        Set rng = Intersect(.cellHome.EntireRow, .colrngVarNames)
        If .nRows = 0 Then
            Set .rngRows = rngToExtent(rng, IsRows:=True)
        Else
            Set .rngRows = Range(.cellHome, .cellHome.Offset(.nRows - 1, 0)).EntireRow
        End If
        
        'Set .nrows as gross number of rows (not just populated) 6/16/23
        If Not .rngRows Is Nothing Then .nRows = .rngRows.Rows.Count
        
        'Set range for model's columns
        If Not .IsCalc Then Set .colrngModel = _
            rngToExtent(Intersect(.cellHome.EntireRow, .colrngModel), IsRows:=False)
        If .IsCalc Then .icolCalc = .colrngModel.Column

        'Set default first scenario column to first column of colrngModel
        If .colrngModel.Columns.Count > 0 Then _
            Set .colrngFirstScenario = .colrngModel.Columns(1)

        'Label the Scenario Name row for multi-column models
        If Not .IsCalc Then
            Intersect(.cellHome.EntireRow, .colrngDesc) = "Scenario Names"
            Intersect(.cellHome.EntireRow, .colrngVarNames) = "Scenario"
        ElseIf Not .IsSuppHeader Then
            Set rng = Intersect(.cellHome.EntireRow, .colrngModel)
            If Len(rng.value) = 0 Then rng.value = "Calculator"
        End If
        
        'Set multicell ranges for model rows (variables) and columns (scenarios)
        Set .rngPopRows = BuildMultiCellRange(.rngRows, .colrngVarNames)
                
        Set .rngPopCols = .colrngModel
        If Not .IsCalc Then Set .rngPopCols = BuildMultiCellRange(.colrngModel, _
            .cellHome.EntireRow)
        
        'Initialize tblSteps and dictionaries for Lite models
        If .IsLiteModel Then
            If Not .PrepStepsForMdl(wkbk, tblStepsTemp) Then GoTo ErrorExit
            Set .tblSteps = tblStepsTemp
        End If
        
        'Set multicell range of rows whose variables are calculated by formula
        If Not .rngPopRows Is Nothing Then .SetRngFormulaRows mdl

        'Set the header range
        If Not .IsSuppHeader And .cellHome.Row > 1 Then _
            Set .rngHeader = Intersect(.cellHome.Offset(-1, 0).EntireRow, .colrngHeader)
            
        'Set range for entire model (not incl. header row)
        Set .rngMdl = .colrngModel.Columns(.colrngModel.Columns.Count)
        Set .rngMdl = Intersect(Range(.cellHome.EntireColumn, .rngMdl), .rngRows)
            
        'Create prefix for variable names
        .NamePrefix = ""
        If .IsMdlNmPrefix Then .NamePrefix = xlName(.mdlName) & "_"
        
        'Set NameRefPrefix for row (Not .IsCalc) and cell (.IsCalc) naming (once per model)
        .NameRefPrefix = "='" & .sht & "'!R"
        .NameRefPrefixCell = "='" & .sht & "'!R"
    End With
    Exit Function
    
ErrorExit:
    errs.RecordErr "Provision", Provision
End Function
'---------------------------------------------------------------------------------------
' Initialize Steps table for use in Lite mdl Refresh
'
' Modified JDL 3/8/23 Set rowCur; 10/18/24 .InitMdl instead of .Init with RefreshAPI
'
Public Function PrepStepsForMdl(wkbk, ByRef tblSteps) As Boolean

    SetErrs PrepStepsForMdl: If errs.IsHandle Then On Error GoTo ErrorExit
    Dim refr As New Refresh, rng As Range
    
    With refr
        If Not .InitMdl(refr, wkbk) Then GoTo ErrorExit
        If Not .PrepExcelStepsSht(refr, tblSteps) Then GoTo ErrorExit
    End With
    
    'Set rowCur attribute to first unused row
    With tblSteps.wkbk.Sheets(tblSteps.sht)
        Set tblSteps.rowCur = tblSteps.colrngCol.Cells(.Rows.Count, 1).End(xlUp).Offset(2, 0)
    End With
    Exit Function
    
ErrorExit:
    errs.RecordErr "PrepStepsForMdl", PrepStepsForMdl
End Function
'---------------------------------------------------------------------------------------
' Format the Scenario Model
' Modified JDL 11/21/23 colrngHeaderFmt max column width 80 for long var names, formulas
'              7/21/25 add iColorGrpRows option
'
Public Sub FormatScenModelClass(mdl)
    Dim rng As Range, str As String, rng2 As Range, colrngFill As Range
    With mdl
    
        'Set or refresh header strings and header column formats
        str = ScenHeader
        If .IsLiteModel Then str = ScenHeaderLite
        If Not .IsSuppHeader Then .rngHeader = Split(str, ",")
        
        'Set text format, column widths and column outline for specific columns
        .colrngModel.Columns(1).Offset(0, -1).ColumnWidth = 4
        If Not .IsLiteModel Then
            Intersect(.rngRows, .colrngNumFmt).NumberFormat = "@"
            Intersect(.rngRows, .colrngFormulas).NumberFormat = "@"
            .colrngGrp.ColumnWidth = 4
            .colrngSubgrp.ColumnWidth = 7
            ColWidthAutofit .colrngDesc
            
            With Range(.colrngNumFmt, .colrngFormulas)
                If HasColOutlining(mdl.wkbk, mdl.sht) Then
                    If mdl.colrngNumFmt.OutlineLevel > 1 Or _
                        mdl.colrngFormulas.OutlineLevel > 1 Then
                        .Columns.Ungroup
                    End If
                End If
                .Columns.Group
            End With
        End If
               
        'Format header row and scenario name row
        If Not .IsSuppHeader Then .rngHeader.Style = "Accent1"
        If Not .rngPopCols Is Nothing Then
            For Each rng In .rngPopCols.EntireColumn
                If Not .IsSuppHeader Then Intersect(.rngHeader.EntireRow, rng).Style = "Accent1"
                If Not .IsCalc Then SetHeaderStyle Intersect(.cellHome.EntireRow, rng)
            Next rng
        End If
        
        'Format Grp rows with bars of specified fill color
        If .iColorGrpRows > 0 And Not .rngRows Is Nothing Then
        
            Set colrngFill = .colrngModel
            If .IsCalc Then Set colrngFill = Range(.colrngGrp, .colrngModel.Offset(0, 3))
            
            'Iterate through model rows and add fill color bars to Grp-populated ones
            Set rng = Intersect(.colrngGrp, .rngRows).Cells(1)
            Do While Not Intersect(rng, .rngRows) Is Nothing
                If Not IsEmpty(rng) Then
                    Intersect(rng.EntireRow, colrngFill).Interior.Color = .iColorGrpRows
                End If
                Set rng = rng.Offset(1, 0)
            Loop
        End If
        
        ColWidthAutofit .colrngDesc, iMaxWidth:=40
        .colrngDesc.WrapText = True
        
        For Each rng In .colrngHeaderFmt
            ColWidthAutofit rng, iMaxWidth:=80, iMinWidth:=10
        Next rng
        If Not .rngPopCols Is Nothing Then
            For Each rng In .rngPopCols
                ColWidthAutofit rng, iMinWidth:=10
            Next rng
        End If
    End With
End Sub
'---------------------------------------------------------------------------------------
' Format column header cells and Scenario Name cells
'
Public Sub SetHeaderStyle(rng)
    rng.Font.Size = 9
    rng.Style = "Note"
End Sub
'---------------------------------------------------------------------------------------
' Set column ranges
'
Function SetColRanges(ByRef mdl) As Boolean

    SetErrs SetColRanges: If errs.IsHandle Then On Error GoTo ErrorExit
    With mdl
        If Not .IsLiteModel Then
            Set .colrngGrp = .cellHome.Offset(0, 0).EntireColumn
            Set .colrngSubgrp = .cellHome.Offset(0, 1).EntireColumn
            Set .colrngDesc = .cellHome.Offset(0, 2).EntireColumn
            Set .colrngVarNames = .cellHome.Offset(0, 3).EntireColumn
            Set .colrngUnits = .cellHome.Offset(0, 4).EntireColumn
            Set .colrngNumFmt = .cellHome.Offset(0, 5).EntireColumn
            Set .colrngFormulas = .cellHome.Offset(0, 6).EntireColumn
            Set .colrngHeader = Range(.colrngGrp, .colrngFormulas)
            Set .colrngHeaderFmt = Range(.colrngVarNames, .colrngFormulas)
        Else
            Set .colrngGrp = .cellHome.Offset(0, 0).EntireColumn
            Set .colrngDesc = .cellHome.Offset(0, 1).EntireColumn
            Set .colrngVarNames = .cellHome.Offset(0, 2).EntireColumn
            Set .colrngUnits = .cellHome.Offset(0, 3).EntireColumn
            Set .colrngHeader = Range(.colrngDesc, .colrngUnits)
            Set .colrngHeaderFmt = Range(.colrngVarNames, .colrngUnits)
        End If
    End With
    Exit Function
    
ErrorExit:
    errs.RecordErr "SetColRanges", SetColRanges
End Function
'---------------------------------------------------------------------------------------
'Parse metadata for a Scenario Model in the workbook
'
'Inputs: mdl [mdlScenario Object]
'      wkbk [Workbook object] workbook containing the Scenario Model
'      Model [String] user-assigned name of the Scenario Model
'
' This returns array of Scenario Model metadata. The function parses a string
' Such definitions can be manually created when the model is created or they can be
' written [refreshed] to Settings sheet and read from there
'
' Scenario Model setting format:
' Setting name: mdl_ModelName (Model function argument)
' Setting value: sht:r,c:IsCalc:IsSuppHeader:IsMdlNmPrefix:IsLiteModel
'              sht - worksheet name where model resides
'              r,c,nrows - row,col home cell on sht; specified nrows (0 to ignore)
'              Booleans - represent as either "T" or "F" in setting
'
'              mdlName is needed for setting range name prefixes and for naming the model's overall range.
'              It must be specified directly as Init arg (.mdlName set in Init) or, if not used here
'              to read Setting with Defn string, it must be included in Defn as optional 9th element
'              (parsed and set here by ParseMdlScenDefn)
'
' Example Definition w/o non-default sName: Process:8,31:0:T:T:T:T:T
'                        non-default sName: Process:8,31:0:T:T:T:T:T:mdlProcess
'
' 4/1/21 JDL    Modified 1/6/22 Add IsRngNames T/F param
'               3/5/23 Add Defn argument and code; 7/17/23 cleanup
'               7/17/23 Refactor and convert to Boolean function; add sName
'               7/27/23 Change criteria for setting mdlName from parsed string
'               2/15/25 Refactor .Init to debug handling sht and mdlName arg overrides
'
Function ParseMdlScenDefn(mdl, Defn) As Boolean
    
    SetErrs ParseMdlScenDefn: If errs.IsHandle Then On Error GoTo ErrorExit
    Dim i As Integer, aryRaw As Variant, aryParams As Variant, aryCellHome As Variant
    With mdl
        
        'Read Setting if model definition not specified as Provision arg
        If IsMissing(Defn) Then Defn = ReadSetting(.wkbk, .mdlName)
        
        'Parse Defn
        aryRaw = Split(Defn, ":")
        
        'Set sheet name if not overridden by .Init arg
        If Len(.sht) = 0 Then .sht = aryRaw(0)
        
        'Set model name (needs to be valid for rng naming); don't override if from arg
        If UBound(aryRaw) = 8 And Len(.mdlName) = 0 Then
            .mdlName = aryRaw(8)
        
        'If mdlName wasn't specified as arg, use sht name
        ElseIf Len(.mdlName) = 0 Then
            .mdlName = xlName(.sht)
        End If

        'Allow for case where sht doesn't exist yet; otherwise, set cellHome
        If SheetExists(.wkbk, .sht) Then
            Set .wksht = .wkbk.Sheets(.sht)
            aryCellHome = Split(aryRaw(1), ",")
            Set .cellHome = .wksht.Cells(CInt(aryCellHome(0)), CInt(aryCellHome(1)))
        End If
    
        'nrows = 0 --> not specified: use extent of var names to set nrows and rngRows)
        .nRows = CInt(aryRaw(2))

        'Booleans: IsCalc, IsSuppHeader, IsRngNames, IsMdlNmPrefix, IsLiteModel
        aryParams = Array(False, False, False, False, False)
        For i = 3 To 7
            If aryRaw(i) = "T" Then aryParams(i - 3) = True
        Next i
        .IsCalc = aryParams(0)
        .IsSuppHeader = aryParams(1)
        .IsRngNames = aryParams(2)
        .IsMdlNmPrefix = aryParams(3)
        .IsLiteModel = aryParams(4)
    End With
    Exit Function
    
ErrorExit:
    errs.RecordErr "ParseMdlScenDefn", ParseMdlScenDefn
End Function
'---------------------------------------------------------------------------------------
' Set Class Attributes from specified, optional arguments
'
' JDL 1/6/22; Modified 11/11/24 fix bug refreshing blank sheet --add cellHome arg
'             1/13/25 .mdlName = xlName(.sht) in case sheet name not valid range name
'             2/15/25 refactor to fix bugs in sht/mdlName arg handling
'
Public Function SetAttsFromArgs(ByRef mdl, IsLiteModel, IsSuppHeader, IsRngNames, IsCalc, _
        IsMdlNmPrefix, nRows, cellHome) As Boolean
    SetErrs SetAttsFromArgs: If errs.IsHandle Then On Error GoTo ErrorExit
    
    With mdl
        
        'Default to name cell (IsCalc) or row (multi-column model)
        .IsRngNames = True
        
        If Not IsMissing(IsLiteModel) Then .IsLiteModel = IsLiteModel
        If Not IsMissing(IsSuppHeader) Then .IsSuppHeader = IsSuppHeader
        If Not IsMissing(IsRngNames) Then .IsRngNames = IsRngNames
        If Not IsMissing(IsCalc) Then .IsCalc = IsCalc
        If Not IsMissing(IsMdlNmPrefix) Then .IsMdlNmPrefix = IsMdlNmPrefix
        .nRows = 0
        If Not IsMissing(nRows) Then .nRows = nRows
        
        'If/Endif to allow for possibility of adding sheet post-init
        If SheetExists(.wkbk, .sht) Then
            Set .wksht = .wkbk.Sheets(.sht)
            If Not .SetCellHome(mdl, cellHome) Then GoTo ErrorExit
        End If
    End With
    Exit Function
    
ErrorExit:
    errs.RecordErr "SetAttsFromArgs", SetAttsFromArgs
End Function
'---------------------------------------------------------------------------------------
' Set Class Attributes from specified, optional arguments
'
' JDL 1/6/22; Modified 5/6/25 bug fix for when .colrngStrInput does not display as text
'               Refactor to improve performance 10/29/25
'
Sub SetRngFormulaRows(ByRef mdl)
    Dim w As Variant, rowsSteps As Range, IsFormula As Boolean, r As Range, r_formula As Range
    Dim sFormula As String, sNumFmt As String, sColWidth As String, sComment As String
    Dim sDropdown As String, sStep As String
    
    With mdl
    
        'Initialize dictionaries for Lite models
        If .IsLiteModel Then
            Set .dStepsFormulas = New Dictionary
            Set .dStepsNumFormats = New Dictionary
            Set .dStepsColWidths = New Dictionary
            Set .dStepsComments = New Dictionary
            Set .dDropdownList = New Dictionary
            
            'Set a row range (multirange of entire rows) for mdl's rows in ExcelSteps
            Set .rngStepsVars = KeyColRng(.tblSteps, Array(.tblSteps.colrngSht), Array(.mdlName))
            If .rngStepsVars Is Nothing Then Exit Sub
            Set .rngStepsVars = Intersect(.rngStepsVars, .tblSteps.colrngCol)
        
            'refresh/Ensure formula cells display as text
            With Intersect(.tblSteps.rngRows, .tblSteps.colrngStrInput)
                .Value2 = .Formula
            End With
        End If

            
        'Initialize to dummy range two rows below .rngRows (Avoid If/Else in loop)
        Set .rngFormulaRows = .rngRows.Rows(.rngRows.Rows.Count).Offset(2, 0).Cells(1, 1)

        'Iterate over rows that contain variables
        For Each w In .rngPopRows
            IsFormula = False
            
            'If not Lite, can just look at Formulas column in mdl
            If Not .IsLiteModel Then
                Set r = Intersect(w.EntireRow, .colrngFormulas)
                
                'Check for a formula in Formula/Row Type column (mod 10/26/22)
                If Not IsEmpty(r) Then
                    r = r.Formula 'Reset in case Text=-formatted cell evals to
                                  '#Name, #N/A or #Spill error
                    If Left(r, 1) = "=" Then IsFormula = True
                End If
            
            'If Lite, search mdl's rows in ExcelSteps and populate dictionaries
            Else
            
                Set r = .rngStepsVars.Find(w.value, lookat:=xlWhole)
                If Not r Is Nothing Then
                    'Check for formula
                    Set r_formula = Intersect(r.EntireRow, .tblSteps.colrngStrInput)
                    sFormula = r_formula.Value2
                    If Left(sFormula, 1) = "=" Then
                        IsFormula = True
                        .dStepsFormulas.Add w.value, sFormula
                    End If
                    
                    'Store number format if present
                    sNumFmt = Intersect(r.EntireRow, .tblSteps.colrngNumFmt).value
                    If Len(sNumFmt) > 0 Then .dStepsNumFormats.Add w.value, sNumFmt
                    
                    'Store column width if present
                    sColWidth = Intersect(r.EntireRow, .tblSteps.colrngWidth).value
                    If Len(sColWidth) > 0 Then .dStepsColWidths.Add w.value, sColWidth
                    
                    'Store comment if present
                    sComment = Intersect(r.EntireRow, .tblSteps.colrngComment).value
                    If Len(sComment) > 0 Then .dStepsComments.Add w.value, sComment
                    
                    'Store dropdown list name if step is col_dropdown
                    sStep = Intersect(r.EntireRow, .tblSteps.colrngStep).Value2
                    If LCase(sStep) = "col_dropdown" Then
                        sDropdown = Intersect(r.EntireRow, .tblSteps.colrngStrInput).value
                        If Len(sDropdown) > 0 Then .dDropdownList.Add w.value, sDropdown
                    End If
                End If
            End If
            
            If IsFormula Then Set .rngFormulaRows = Union(.rngFormulaRows, w)
        Next w
        
        'Remove dummy range from result
        Set .rngFormulaRows = Intersect(.rngFormulaRows, .rngRows)
    End With
End Sub
'---------------------------------------------------------------------------------------
' Set Class Attributes from specified, optional arguments
'
' JDL 1/6/22; Modified 5/6/25 bug fix for when .colrngStrInput does not display as text
'
Sub SetRngFormulaRows_prev(ByRef mdl, tblSteps)
    Dim w As Variant, rowsSteps As Range, IsFormula As Boolean, r As Range, r_formula As Range
    With mdl
    
        'Set a row range (multirange of entire rows) for mdl's rows in ExcelSteps
        If .IsLiteModel Then
            Set .rngStepsVars = KeyColRng(tblSteps, Array(tblSteps.colrngSht), Array(.mdlName))
            If .rngStepsVars Is Nothing Then Exit Sub
            Set .rngStepsVars = Intersect(.rngStepsVars, tblSteps.colrngCol)
        End If

        'Iterate over rows that contain variables
        For Each w In .rngPopRows
            IsFormula = False
            
            'If not Lite, can just look at Formulas column in mdl
            If Not .IsLiteModel Then
                Set r = Intersect(w.EntireRow, .colrngFormulas)
                
                'Check for a formula in Formula/Row Type column (mod 10/26/22)
                If Not IsEmpty(r) Then
                    r = r.Formula 'Reset in case Text=-formatted cell evals to
                                  '#Name, #N/A or #Spill error
                    If Left(r, 1) = "=" Then IsFormula = True
                End If
            
            'If Lite, need to search mdl's rows in ExcelSteps  (mod 10/26/22)
            Else
                Set r = FindInRange(.rngStepsVars, w.value)
                If Not r Is Nothing Then
                
                    'Ensure formula displayed as text (not evaluated formula/error etc.)
                    Set r_formula = Intersect(r.EntireRow, tblSteps.colrngStrInput)
                    r_formula.value = r_formula.Formula
                    IsFormula = (Left(tblSteps.TableLoc(r, tblSteps.colrngStrInput), 1) = "=")
                End If
            End If
            
            'Add variable to multirange for calculated variables
            If IsFormula Then
                If .rngFormulaRows Is Nothing Then
                    Set .rngFormulaRows = w
                Else
                    Set .rngFormulaRows = Union(.rngFormulaRows, w)
                End If
            End If
        Next w
    End With
End Sub
'---------------------------------------------------------------------------------------
' Apply border around model
'
' JDL mod 5/8/23
Sub ApplyBorderAroundModel(mdl, Optional IsBufferRow = False, Optional IsBufferCol = False)
    Dim xlEdge As Variant, rng As Range
    
    Set rng = mdl.rngMdl
    If IsBufferRow Then Set rng = Union(rng, mdl.rngMdl.Offset(1, 0))
    If IsBufferCol Then Set rng = Union(rng, rng.Offset(0, 1))
    
    For Each xlEdge In Array(xlEdgeLeft, xlEdgeTop, xlEdgeBottom, xlEdgeRight)
        With rng.Borders(xlEdge)
            .LineStyle = xlContinuous
            .Weight = xlMedium
        End With
    Next xlEdge
End Sub
'---------------------------------------------------------------------------------------
' Clear model cell values and outline
'
' Modified: 3/13/23 JDL fix bug with IsBufferCol
'
Function ClearModel(mdl, Optional IsBufferRow = False, Optional IsBufferCol = False) As Boolean

    SetErrs ClearModel: If errs.IsHandle Then On Error GoTo ErrorExit
    Dim rng As Range
    With mdl
    
        'Clear model cells
        Set rng = .rngMdl
        If IsBufferRow Then Set rng = Union(rng, .rngMdl.Offset(1, 0))
        If IsBufferCol Then Set rng = Union(rng, rng.Offset(0, 1))
        rng.Clear
        
        'Clear header cells
        If Not .IsSuppHeader Then
            .rngHeader.Clear
            Intersect(.rngHeader.EntireRow, .colrngModel).Clear
        End If
        
        'Clear column outline
        If Not .IsLiteModel Then
            If HasColOutlining(.wkbk, .sht) Then _
                Range(.colrngNumFmt, .colrngFormulas).Columns.Ungroup
        End If
    End With
    Exit Function
    
ErrorExit:
    errs.RecordErr "ClearModel", ClearModel
End Function
'-----------------------------------------------------------------------------------------------
' Clear model input values (non-formula rows) from provisioned mdlScenario instance
' JDL 8/1/25; Modified 10/1/25 to work with multicolumn Scenario Models
'
Public Function ClearMdlInputs(ByRef mdl As mdlScenario) As Boolean
    SetErrs ClearMdlInputs: If errs.IsHandle Then On Error GoTo ErrorExit
    Dim CellVarName As Range, rngIntersect As Range, IsInputRow As Boolean

    With mdl
        If .rngPopCols Is Nothing Then Exit Function
        
        For Each CellVarName In .rngPopRows

            ' Determine if variable is an input row (not in formula rows)
            If .rngFormulaRows Is Nothing Then
                IsInputRow = True
            Else
                IsInputRow = (Intersect(CellVarName, .rngFormulaRows) Is Nothing)
            End If
            
            ' Clear the input row if needed (but don't clear Scenario names)
            If IsInputRow And CellVarName.Value2 <> "Scenario" Then
                Set rngIntersect = Intersect(CellVarName.EntireRow, .rngPopCols.EntireColumn)
                If Not rngIntersect Is Nothing Then rngIntersect.ClearContents
            End If
        Next CellVarName
    End With
    Exit Function
    
ErrorExit:
    errs.RecordErr "ClearMdlInputs", ClearMdlInputs
End Function
'-----------------------------------------------------------------------------------------------
' Import model input values from external file
' JDL 8/1/25; update OpenFile args 9/4/25
'
Public Function ImportMdlInputs(ByVal mdl As mdlScenario, ByVal filepath As String) As Boolean
    SetErrs ImportMdlInputs: If errs.IsHandle Then On Error GoTo ErrorExit
    Dim wkbkImport As Workbook, varname As String, val As Variant
    Dim tblImport As New tblRowsCols
    
    ' Open import file and Provision tblImport
    If Not ExcelSteps.OpenFile(filepath, wkbkImport) Then GoTo ErrorExit
    With tblImport
        If Not .Provision(tblImport, wkbkImport, False, sht:=wkbkImport.Worksheets(1).Name) _
            Then GoTo ErrorExit
    
        ' Iteratively import values from tblImport rows
        Set .rowCur = .wksht.Rows(2)
        Do While Intersect(.rowCur, .wksht.Columns(1)).value <> ""
            varname = Intersect(.rowCur, .wksht.Columns(2)).value
            val = Intersect(.rowCur, .wksht.Columns(3)).value
            mdl.SetScenModelLoc mdl, varname, val
            Set .rowCur = .rowCur.Offset(1, 0)
        Loop
    End With
    
    ' Close the import workbook
    wkbkImport.Close False: Set wkbkImport = Nothing
    Exit Function
    
ErrorExit:
    If Not wkbkImport Is Nothing Then wkbkImport.Close False: Set wkbkImport = Nothing
    errs.RecordErr "ImportMdlInputs", ImportMdlInputs
End Function
'-----------------------------------------------------------------------------------------------
' Export models input values to external file
' JDL 8/4/25; Update to use SafeSaveAs 10/16/25; Simplify to wkbkExisting.Close 10/27/25
'
Public Function ExportMdlInputs(ByVal mdl As mdlScenario, ByVal filepath As String) As Boolean
    SetErrs ExportMdlInputs: If errs.IsHandle Then On Error GoTo ErrorExit
    Dim wkbkExport As Workbook, scenario As String, col As Range
    Dim tblExport As New tblRowsCols, wkbkExisting As Workbook, fileName As String
    
    ' Close the filepath workbook if it's already open (xxx move to utility)
    fileName = Right(filepath, Len(filepath) - InStrRev(filepath, Application.PathSeparator))
    For Each wkbkExisting In Workbooks
        If wkbkExisting.Name = fileName Then
            wkbkExisting.Close False
            Exit For
        End If
    Next wkbkExisting
    
    With mdl
        ' Prepare export workbook and tblExport
        If Not .PrepTblExport(tblExport, wkbkExport) Then GoTo ErrorExit
        
        ' Export data for each scenario column in the model
        If .IsCalc Then
            If Not .ExportInputs(mdl, tblExport, "Calculator", .colrngModel) Then GoTo ErrorExit
        Else
            For Each col In .rngPopCols.Columns
                scenario = Intersect(.ScenModelLoc(mdl, "Scenario"), col).value
                If Len(scenario) > 0 Then
                    If Not .ExportInputs(mdl, tblExport, scenario, col) Then GoTo ErrorExit
                End If
            Next col
        End If
        
        ' Provision and Refresh with formatting
        If Not tblExport.Provision(tblExport, wkbkExport, IsFormat:=True, _
            sht:=wkbkExport.ActiveSheet.Name, IsSetTblNames:=False, IsSetColNames:=False) _
            Then GoTo ErrorExit
        
        ' Save and close the export workbook (overwrite if exists)
        If Not SafeSaveAs(wkbkExport, filepath) Then GoTo ErrorExit
        wkbkExport.Close False
    End With
    Exit Function
    
ErrorExit:
    If Not wkbkExport Is Nothing Then wkbkExport.Close False: Set wkbkExport = Nothing
    errs.RecordErr "ExportMdlInputs", ExportMdlInputs
End Function
'-----------------------------------------------------------------------------------------------
' Clear scenario-containing column ranges from provisioned model (leaves template untouched)
' JDL 8/5/25
'
Public Function ClearMdlScenarios(ByVal mdl As mdlScenario) As Boolean
    SetErrs ClearMdlScenarios: If errs.IsHandle Then On Error GoTo ErrorExit
    Dim colArea As Range, col As Range

    With mdl
        If .rngPopCols Is Nothing Then Exit Function
        
        ' Clear each scenario column
        For Each colArea In .rngPopCols.Areas
            For Each col In colArea.Columns
                Intersect(col.EntireColumn, .rngRows.EntireRow).Clear
            Next col
        Next colArea
    End With
    Exit Function
    
ErrorExit:
    errs.RecordErr "ClearMdlScenarios", ClearMdlScenarios
End Function
'-----------------------------------------------------------------------------------------------
' Prepare export workbook and tblExport with headers
' JDL 8/4/25
'
Public Function PrepTblExport(ByRef tblExport As tblRowsCols, ByRef wkbkExport As Workbook) As Boolean
    SetErrs PrepTblExport: If errs.IsHandle Then On Error GoTo ErrorExit
    Dim sht As String
    
    ' Create new workbook and provision tblExport
    Set wkbkExport = Workbooks.Add
    sht = wkbkExport.Worksheets(1).Name
    
    ' For Mac Excel
    wkbkExport.Activate
    
    If Not tblExport.Provision(tblExport, wkbkExport, False, sht:=sht) Then GoTo ErrorExit

    ' Write headers
    tblExport.wksht.Range("A1:D1").value = Split("Scenario,Variable,Value,Description", ",")

    ' Set rowCur to first data row (row 2)
    Set tblExport.rowCur = tblExport.wksht.Rows(2)

    Exit Function
    
ErrorExit:
    errs.RecordErr "PrepTblExport", PrepTblExport
End Function
'-----------------------------------------------------------------------------------------------
' Export data for one scenario column
' JDL 8/4/25
'
Public Function ExportInputs(ByVal mdl, tblExport As tblRowsCols, _
                            ByVal scenario As String, ByVal rngCol As Range) As Boolean
    SetErrs ExportInputs: If errs.IsHandle Then On Error GoTo ErrorExit
    Dim CellVarName As Range, rngIntersect As Range, IsInputRow As Boolean
    Dim varname As String, val As Variant

    With mdl
        For Each CellVarName In .rngPopRows
            ' Skip the "Scenario" variable itself
            varname = Intersect(CellVarName, .colrngVarNames).value
            If varname <> "Scenario" Then
                ' Determine if variable is an input row (not in formula rows)
                If .rngFormulaRows Is Nothing Then
                    IsInputRow = True
                Else
                    IsInputRow = (Intersect(CellVarName, .rngFormulaRows) Is Nothing)
                End If
                
                ' Export the input row if appropriate
                If IsInputRow Then
                    Set rngIntersect = Intersect(CellVarName.EntireRow, rngCol)
                    If Not rngIntersect Is Nothing Then
                        val = rngIntersect.value

                        If Len(val) > 0 Then
                            With tblExport
                                Intersect(.rowCur, .wksht.Columns(1)).value = scenario
                                Intersect(.rowCur, .wksht.Columns(2)).value = varname
                                Intersect(.rowCur, .wksht.Columns(3)).value = val
                                Intersect(.rowCur, .wksht.Columns(4)).value = _
                                    Intersect(CellVarName.EntireRow, mdl.colrngDesc).value
                                Set .rowCur = .rowCur.Offset(1, 0)
                            End With
                        End If
                    End If
                End If
            End If
        Next CellVarName
    End With
    Exit Function
    
ErrorExit:
    errs.RecordErr "ExportInputs", ExportInputs
End Function
'-----------------------------------------------------------------------------------------------
' Delete mdl Range names
' The .Visible property is used to skip hidden _xlfn.SINGLE name created by dynamic array
' glitch from circa 2020 Excel update. See: https://stackoverflow.com/questions/59121799
'
' JDL 6/20/23
'
Function DeleteMdlRangeNames(mdl, Optional sPrefix As String = "") As Boolean
    
    SetErrs DeleteMdlRangeNames: If errs.IsHandle Then On Error GoTo ErrorExit
    Dim w As Variant, nchars As Integer
    
    'If user doesn't override prefix, use ModelName
    If Len(sPrefix) < 1 Then sPrefix = mdl.mdlName
    
    nchars = Len(sPrefix)
    For Each w In mdl.wkbk.Names
        If (Left(w.Name, nchars) = sPrefix) And w.Visible Then w.Delete
    Next w
    Exit Function
    
ErrorExit:
    errs.RecordErr "DeleteMdlRangeNames", DeleteMdlRangeNames
End Function
'------------------------------------------------------------------------------------------------
' Add a dropdown to a Scenario Model variable value(s)
'
' Created:   1/3/22 JDL      Modified: 2/1/22 FindInRange
'
Sub AddDropdownToVariable(mdl, sVar, sDropdownFormula)
    Dim c As Range
    Set c = FindInRange(mdl.colrngVarNames, sVar)
    If c Is Nothing Then Exit Sub
    
    Set c = Intersect(c.EntireRow, mdl.colrngModel)
    
    'Skip if error such as named range not existing
    On Error GoTo 0
    AddValidationList c, sDropdownFormula
End Sub
'---------------------------------------------------------------------------------------
'Lookup and return Scenario Model value
'Inputs:    sVar [String] Scenario Model variable name for lookup
'         rngCol [Range] table column range
'
' Created: 2/4/21 JDL  Modified: 8/8/25 Use Intersect in case multiple same name in col
'
Function ScenModelLoc(mdl, sVar, Optional rngCol) As Range
    Dim rngRow As Range
    
    With mdl
        Set rngRow = FindInRange(Intersect(.rngRows, .colrngVarNames), sVar)
        If rngRow Is Nothing Then Exit Function
    
        If IsMissing(rngCol) Then Set rngCol = .colrngModel
    End With
    
    Set ScenModelLoc = Intersect(rngRow.EntireRow, rngCol)
End Function
'---------------------------------------------------------------------------------------
' Set value for specified Scenario Model variable
' Inputs: sVar [String] Scenario Model variable name for lookup
'         rngCol [Range] table column range
'         val [variant] value to set at intersection of rngCell.row and rngCol
'
' Created: 2/4/21 JDL  Modified: 8/8/25 Use Intersect in case multiple same name in col
'
Sub SetScenModelLoc(mdl, sVar, val, Optional rngCol)
    Dim rngRow As Range
    With mdl
        Set rngRow = FindInRange(Intersect(.rngRows, .colrngVarNames), sVar)
        If rngRow Is Nothing Then Exit Sub
    
        If IsMissing(rngCol) Then Set rngCol = .colrngModel
    End With
    
    Intersect(rngRow.EntireRow, rngCol) = val
End Sub
'---------------------------------------------------------------------------------------
' Refresh a Scenario model
'
' Created: JDL      Modified: 1/7/22 - Refactor to use mdlRow Class
'                          6/30/22 - refactor mdlRow Class
'                          10/25/25 - fix SetHeaderStyle range issue
'                          10/29/25 - use mdl.tblSteps attribute
Function Refresh(mdl) As Boolean
    SetErrs Refresh: If errs.IsHandle Then On Error GoTo ErrorExit
    Dim r As Object, rngRow As Variant
            
    With mdl
        If Not CkVarAndScenNames(mdl) Then GoTo ErrorExit
        
        'Delete previous range names if using model name prefix
        If .IsRngNames And .IsMdlNmPrefix Then
            If Not DeleteMdlRangeNames(mdl) Then GoTo ErrorExit
        End If
        
        If .IsRngNames Then NameMdlColumns mdl
    End With
    
    'Loop through rows and apply the model - exit if no model rows
    If Not mdl.rngPopRows Is Nothing Then
        For Each rngRow In mdl.rngPopRows
        
            'Instance/initialize mdlRow Class to hold row attributes
            Set r = New mdlRow
            r.Init r, mdl, rngRow
            
            'Name cell/row
            If mdl.IsRngNames Then r.NameRow r, mdl
            
            'Set Number Format and formula
            If Not r.rngMdlCells Is Nothing Then
                r.FormatRow r
                If r.HasLstValidation Then r.AddListValidation r
                If Not r.WriteRowFormula(r, mdl) Then GoTo ErrorExit
            End If
            
            'Format template columns
            SetHeaderStyle Intersect(rngRow.EntireRow, mdl.colrngHeaderFmt)
            SetBorders Intersect(rngRow, mdl.colrngHeaderFmt), xlContinuous, True
        Next rngRow
    End If
    
    mdl.FormatScenModelClass mdl
    Exit Function
    
ErrorExit:
    errs.RecordErr "Refresh", Refresh
End Function
'---------------------------------------------------------------------------------------
' Refresh a Scenario model
' JDL pre 1/22; Modified: 7/17/25 refactor for speedup - not effective and doesn't handle
' formula errors; 10/29/25 use mdl.tblSteps attribute
'
Function RefreshNew(mdl) As Boolean
    SetErrs RefreshNew: If errs.IsHandle Then On Error GoTo ErrorExit
    Dim r As Object, rngRow As Variant
    Dim aryFormulas As Variant, rngFormulas As Range
    Dim aryRows As Variant, rngMdlRows As Range
    
    With mdl
        If Not CkVarAndScenNames(mdl) Then GoTo ErrorExit
        If .IsRngNames Then NameMdlColumns mdl
    End With
    
    'Loop through rows and apply the model - exit if no model rows
    If Not mdl.rngPopRows Is Nothing Then
    
        'If using mdlName range name prefix, mass-delete previous versions of those names
        'If mdl.IsMdlNmPrefix Then DeleteRngNamesWithPrefix mdl.wkbk, mdl.NamePrefix

        For Each rngRow In mdl.rngPopRows
        
            'Instance/initialize mdlRow Class to hold row attributes
            Set r = New mdlRow
            r.Init r, mdl, rngRow
            
            'Name cell/row
            If mdl.IsRngNames Then r.NameRow r, mdl
            
            'Set Number Format and formula
            If Not r.rngMdlCells Is Nothing Then
                r.FormatRow r
                If r.HasLstValidation Then r.AddListValidation r
            End If
            
            'For performance, accumulate formula rows to write as a block
            If r.HasFormula Then
                If Not mdl.UpdateFormulaRng(r, mdl, aryFormulas, rngFormulas) Then GoTo ErrorExit
            End If
        Next rngRow
        
        'Write the last block of formulas if any
        If Not rngFormulas Is Nothing Then
            If Not mdl.WriteRngFormulas(mdl, aryFormulas, rngFormulas) Then GoTo ErrorExit
        End If
    End If
        
    'Format template columns (whole region format for performance)
    With mdl
        If .rngRows Is Nothing Then Exit Function
        
        SetHeaderStyle Intersect(.rngRows, .colrngHeaderFmt)
        SetBorders Intersect(.rngRows, .colrngHeaderFmt), xlContinuous, True
        For Each rngRow In .rngRows
            If IsEmpty(Intersect(rngRow, .colrngVarNames)) Then _
                Intersect(rngRow, .colrngHeaderFmt).ClearFormats
        Next rngRow
    
        'Format overall template
        .FormatScenModelClass mdl
    End With
    Exit Function
    
ErrorExit:
    errs.RecordErr "RefreshNew", RefreshNew
End Function
'-----------------------------------------------------------------------------------------
' Update formula string array and formula-containing rng during .rngPopRows iteration
' JDL 7/17/25
Function UpdateFormulaRng(r, mdl, aryFormulas, rngFormulas) As Boolean
    SetErrs UpdateFormulaRng: If errs.IsHandle Then On Error GoTo ErrorExit
    Dim rngFormulas_temp As Range
    
    ' Initialize aryFormulas and rngFormulas with first formula-containing row
    If rngFormulas Is Nothing Then
        Set rngFormulas = r.rngVarRow
        ReDim aryFormulas(1 To 1)
        aryFormulas(1) = r.sFormula
    Else
        Set rngFormulas_temp = rngFormulas
        Set rngFormulas = Union(rngFormulas, r.rngVarRow)
        
        'Revert and write formulas if have moved to next non-contiguous formula row
        If rngFormulas.Areas.Count > 1 Then
            Set rngFormulas = rngFormulas_temp
            If Not mdl.WriteRngFormulas(mdl, aryFormulas, rngFormulas) Then GoTo ErrorExit
            
            'start a new contiguous formula row block
            Set rngFormulas = r.rngVarRow
            ReDim aryFormulas(1 To 1)
            aryFormulas(1) = r.sFormula
            
        'Add the current formula to aryFormulas and continue with that contiguous range
        Else
            ReDim Preserve aryFormulas(1 To UBound(aryFormulas) + 1)
            aryFormulas(UBound(aryFormulas)) = r.sFormula
        End If
    End If
    Exit Function
    
ErrorExit:
    errs.RecordErr "UpdateFormulaRng", UpdateFormulaRng
End Function
'-----------------------------------------------------------------------------------------
' Write array of formulas for current contiguous row range
' JDL 7/17/25
Function WriteRngFormulas(mdl, aryFormulas, rngFormulas) As Boolean
    SetErrs WriteRngFormulas: If errs.IsHandle Then On Error GoTo ErrorExit
    Dim colBlock As Range, rngDest As Range
    
    If mdl.rngPopCols Is Nothing Then Exit Function
    
    'Iterate over .rngPopCols non-contiguous column blocks
    For Each colBlock In mdl.rngPopCols.Areas
        Set rngDest = Intersect(rngFormulas, colBlock.EntireColumn)
        
        ' Apply formulas to the target range
        rngDest.FormulaR1C1 = Application.Transpose(aryFormulas)
        rngDest.Style = "Calculation"
        rngDest.Font.ColorIndex = xlAutomatic
    Next colBlock
    Exit Function
    
ErrorExit:
    errs.RecordErr "WriteRngFormulas", WriteRngFormulas
End Function
'---------------------------------------------------------------------------------------
' Check row and column name strings are Excel compatible and non-redundant
'
' Created: 1/7/22 JDL
'
Function CkVarAndScenNames(mdl) As Boolean
    CkVarAndScenNames = True
    With mdl
        If Not CheckNames(.rngPopRows) Then CkVarAndScenNames = False
        If Not .IsCalc Then
            If Not CheckNames(.rngPopCols) Then CkVarAndScenNames = False
        End If
    End With
End Function
'---------------------------------------------------------------------------------------
' Name multi-column model column ranges
'
' Created: 1/6/22 JDL   Modified 12/12/24 append .NamePrefix (may be blank)
'
Sub NameMdlColumns(mdl)
    Dim c As Range, rngRow As Range, col_name As String
    With mdl
        Set rngRow = .cellHome.EntireRow
        If Not .IsCalc And Not .rngPopCols Is Nothing Then
            For Each c In .rngPopCols
            
                col_name = .NamePrefix & c.value
                
                MakeXLName .wkbk, .NamePrefix & Intersect(rngRow, c).value, _
                    MakeRefNameString(.sht, icol1:=c.Column)
            Next c
        ElseIf .IsCalc Then
            MakeXLName .wkbk, .mdlName, _
                MakeRefNameString(.sht, 0, 0, .icolCalc, .icolCalc)
        End If
    End With
End Sub
'---------------------------------------------------------------------------------------
'Methods related to SwapModel capability to replace a Scenario Model with a second one
'stored on tblImport Sheet
'---------------------------------------------------------------------------------------
' SwapModels master procedure for transferring a Scenario Model to a rows/cols "input
' deck" on shtTblImport and transferring a rows/cols version to Scenario Model as
' a replacement
'
' Inputs: ModelNew [String] Name of model to swap to mdlDest from tblImport sheet table
'         ModelDest [String] dest mdl name (for defn lookup if not specified)
'         ModelDefnDest [Optional String] destination model Defn (if pre-set)
'
'JDL 3/14/23 JDL    Modified: 7/14/23 Major refactoring; 9/30/24 tblSteps instead of tblSteps
'
Function SwapModels(wkbk As Workbook, Optional ByVal ModelNew As String, _
        Optional ByVal ModelDest As String, Optional ByVal ModelDefnDest As String) As Boolean

    SetErrs SwapModels: If errs.IsHandle Then On Error GoTo ErrorExit
    Dim mdlDest As New mdlScenario, tblImp As New tblRowsCols
    
    'Initialize and Provision Classes for the swap
    If Len(ModelDefnDest) < 1 Then ModelDefnDest = ReadSetting(wkbk, ModelDest)
    If Not InitSwapModels(mdlDest, tblImp, wkbk, ModelDefnDest) Then GoTo ErrorExit
    
    'If there is a previous model, transfer it to tblImport sheet
    If mdlDest.nRows > 1 Then
        If Not TransferToTblImport(mdlDest, tblImp) Then GoTo ErrorExit
    End If
    
    'If requested, transfer new mdl from tblImport sheet table
    If Len(ModelNew) > 0 Then
        If Not TransferToMdlDest(mdlDest, tblImp, ModelNew, ModelDefnDest) Then GoTo ErrorExit
        ApplyBorderAroundModel mdlDest, IsBufferRow:=True, IsBufferCol:=True
    End If
    Exit Function
    
ErrorExit:
    errs.RecordErr "SwapModels", SwapModels
End Function
'---------------------------------------------------------------------------------------
' Initialize a swap or other move between Scenario Model and tblImport
'
'JDL 3/14/23 JDL    Modified: 7/14/23 Refactoring; 9/30/24 tblSteps instead of tblSteps
'                             10/29/25 use mdl.tblSteps attribute
'
Function InitSwapModels(mdlDest As mdlScenario, tblImp As tblRowsCols, _
        ByVal wkbk As Workbook, ModelDefnDest As String) As Boolean

    SetErrs InitSwapModels: If errs.IsHandle Then On Error GoTo ErrorExit
    Dim refr As New Refresh
    Dim tblStepsTemp As New tblRowsCols 'For passing instance as ByRef argument to called function
    
    'Provision/initialize destination model (this will init mdlDest.tblSteps if IsLiteModel)
    With mdlDest
        If Not .Provision(mdlDest, wkbk, Defn:=ModelDefnDest) Then GoTo ErrorExit
        Set .rowCur = .cellHome.EntireRow
    End With
        
    'Provision/initialize tblImport sheet table (xxx mod to use .End(xlUp) for .rowCur???
    With tblImp
        If Not .Provision(tblImp, wkbk, False, shtTblImp, nCols:=10, IsSetColRngs:=True) Then GoTo ErrorExit
        
        'Initialize rowCur to first blank row; initialize rngRows anyway if no data
        If .rngRows Is Nothing Then
            Set .rowCur = .cellHome.EntireRow
            Set .rngRows = .rowCur
        Else
            Set .rowCur = .rngRows.Rows(.rngRows.Count).Offset(1, 0)
        End If
    End With
     
    'Provision/Initialize ExcelSteps sheet if needed (for non-Lite models or if not yet done)
    If mdlDest.tblSteps Is Nothing Then
        If Not refr.InitTbl(refr, wkbk:=wkbk, IsReplace:=True, IsTblFormat:=False) Then GoTo ErrorExit
        If Not refr.PrepExcelStepsSht(refr, tblStepsTemp) Then GoTo ErrorExit
        Set mdlDest.tblSteps = tblStepsTemp
        With mdlDest.tblSteps
            Set .rowCur = .colrngCol.Cells(.wksht.Rows.Count, 1).End(xlUp).Offset(2, 0).EntireRow
        End With
    End If
    Exit Function

ErrorExit:
    errs.RecordErr "InitSwapModels", InitSwapModels
End Function
'-----------------------------------------------------------------------------------------------
'-----------------------------------------------------------------------------------------------
' TransferToTblImport Procedure - transfer model from mdlDest region to tblImp rows/cols
'-----------------------------------------------------------------------------------------------
'-----------------------------------------------------------------------------------------------
' Procedure - Transfer model from mdlDest Scenario Model region to tblImport sheet rows/cols
'
' JDL 8/22/23; 9/30/24 tblSteps instead of tblSteps; 10/29/25 use mdl.tblSteps attribute
'
Function TransferToTblImport(ByVal mdlDest As mdlScenario, ByRef tblImp As tblRowsCols) As Boolean

    SetErrs TransferToTblImport: If errs.IsHandle Then On Error GoTo ErrorExit
    Dim ModelPrev As String
    With mdlDest
        If Not .ReadModelName(mdlDest, ModelPrev) Then GoTo ErrorExit
        If Not .TblImportDeleteModel(tblImp, ModelPrev) Then GoTo ErrorExit
        If Not .TransferMdlDestRows(mdlDest, tblImp, ModelPrev) Then GoTo ErrorExit
        If Not .DeleteTblImpTrailingBlankRows(tblImp) Then GoTo ErrorExit
        If Not .ClearModel(mdlDest, IsBufferRow:=True, IsBufferCol:=True) Then GoTo ErrorExit
        If Not .StepsDeleteMdl(mdlDest) Then GoTo ErrorExit
    End With
    Exit Function
    
ErrorExit:
    errs.RecordErr "TransferToTblImport", TransferToTblImport
End Function
'-----------------------------------------------------------------------------------------------
' Set ModelPrev by reading mdl_name variable value from mdlDest; set mdlDest.rngStepsVars
'
' JDL 7/27/23   Modified 8/22/23 Add set mdlDest.mdlName and .rngStepsVars
'                        9/30/24 tblSteps instead of tblSteps; 10/29/25 use mdl.tblSteps
'
Function ReadModelName(ByVal mdlDest As mdlScenario, ModelPrev As String) As Boolean

    SetErrs ReadModelName: If errs.IsHandle Then On Error GoTo ErrorExit
    Dim rng As Range
    
    With mdlDest
        Set rng = FindInRange(.colrngVarNames, "mdl_name")
        If errs.IsFail(rng Is Nothing, 1) Then GoTo ErrorExit
        
        ModelPrev = Intersect(rng.EntireRow, .colrngModel)
        
        'Set ExcelSteps range for ModelPrev
        .mdlName = ModelPrev
        If Not .tblSteps Is Nothing Then
            Set .rngStepsVars = KeyColRng(.tblSteps, Array(.tblSteps.colrngSht), Array(.mdlName))
        End If
    End With
    Exit Function
    
ErrorExit:
    errs.RecordErr "ReadModelName", ReadModelName
End Function
'---------------------------------------------------------------------------------------
' Delete a model from tblImport sheet
'
' Created: 3/6/23 JDL Modified argument and tblImp name 8/17/23
'
Function TblImportDeleteModel(tblImp, Model) As Boolean

    SetErrs TblImportDeleteModel: If errs.IsHandle Then On Error GoTo ErrorExit
    Dim rngModel As Range, wkbk As Workbook

    'Exit if tblImport sheet is empty
    If tblImp.rngRows Is Nothing Then Exit Function
    
    'Set range for model on tblImport sheet and delete its rows if any
    Set rngModel = rngKeycolRows(tblImp, tblImp.colrngMdlName, Model)
    If rngModel Is Nothing Then Exit Function
    rngModel.Delete
    
    'Re-initialize tblImp
    Set wkbk = tblImp.wkbk
    Set tblImp = New tblRowsCols
    If Not tblImp.Provision(tblImp, wkbk, False, sht:=shtTblImp, nCols:=10, IsSetColRngs:=True) _
            Then GoTo ErrorExit
    Exit Function
ErrorExit:
    errs.RecordErr "tblImportDeleteModel", TblImportDeleteModel
End Function
'-----------------------------------------------------------------------------------------------
' Transfer mdlDest Scenario Model rows to tblImport rows/columns table
'
' JDL 7/27/23; 9/30/24 tblSteps instead of tblSteps; 10/29/25 use mdl.tblSteps
'
Function TransferMdlDestRows(mdlDest, tblImp As tblRowsCols, _
        ByVal ModelPrev As String) As Boolean

    SetErrs TransferMdlDestRows: If errs.IsHandle Then On Error GoTo ErrorExit
    Dim R_MI As New mdlImportRow, r As Range
   
    With mdlDest
        Set .rowCur = .cellHome.EntireRow
        R_MI.Model = ModelPrev
        
        'Iterate over mdlDest rows; Use mdlImportRow Class to transfer to tblImport sheet
        For Each r In .rngRows
            With R_MI
                If Not .Init(R_MI) Then GoTo ErrorExit
                If Not .ReadMdlDestRow(R_MI, mdlDest) Then GoTo ErrorExit
                If Not .ReadStepsRow(R_MI, mdlDest.tblSteps, mdlDest.rngStepsVars) Then GoTo ErrorExit
                If Not .SetBooleanFlags(R_MI) Then GoTo ErrorExit
                If Not .ToTblWriteRow(R_MI, mdlDest, tblImp) Then GoTo ErrorExit
            End With
        Next r
    End With
    Exit Function
    
ErrorExit:
    errs.RecordErr "TransferMdlDestRows", TransferMdlDestRows
End Function
'---------------------------------------------------------------------------------------
' Delete trailing blank rows, if any, from tblImp
'
' JDL 3/14/23 JDL    Modified 8/22/23
'
Function DeleteTblImpTrailingBlankRows(ByRef tblImp As tblRowsCols) As Boolean

    SetErrs DeleteTblImpTrailingBlankRows: If errs.IsHandle Then On Error GoTo ErrorExit

    With tblImp
        Set .rowCur = .rngRows.Rows(.rngRows.Rows.Count)
        Do While Intersect(.rowCur, .colrngVarName) = "<blank>"
            .rowCur.Clear
            Set .rowCur = .rowCur.Offset(-1, 0)
            Set .rngRows = Range(.rngRows.Rows(1), .rowCur)
        Loop
    End With
    Exit Function

ErrorExit:
    errs.RecordErr "DeleteTblImpTrailingBlankRows", DeleteTblImpTrailingBlankRows
End Function
'---------------------------------------------------------------------------------------
' Delete Steps rows for a model
'
' Created: 3/14/23 JDL; 9/30/24 tblSteps instead of tblSteps; 10/29/25 use mdl.tblSteps
'
Function StepsDeleteMdl(mdl) As Boolean

    SetErrs StepsDeleteMdl: If errs.IsHandle Then On Error GoTo ErrorExit
    Dim i As Integer
    With mdl
    
        'Set multirange for model's rows on ExcelSteps sheet
        If Not .SetStepsRowRange(mdl) Then GoTo ErrorExit
        If .rngStepsVars Is Nothing Then Exit Function
        
        'Loop in reverse order and delete
        For i = .rngStepsVars.Rows.Count To 1 Step -1
            .rngStepsVars.Rows(i).Delete
        Next i
    End With
Exit Function
    
ErrorExit:
    errs.RecordErr "StepsDeleteMdl", StepsDeleteMdl
End Function
'---------------------------------------------------------------------------------------
' Set a row range (multirange of entire rows) for mdl's rows in ExcelSteps
'
' Modified: 6/20/23 for "ProcessParams2" (sMdlProcess2) in ExcelSteps Sheet column
'           9/30/24 tblSteps instead of tblSteps; 10/29/25 use mdl.tblSteps
'
Function SetStepsRowRange(mdl) As Boolean

    SetErrs SetStepsRowRange: If errs.IsHandle Then On Error GoTo ErrorExit
    
    With mdl
        If Not .tblSteps Is Nothing Then
            Set .rngStepsVars = KeyColRng(.tblSteps, Array(.tblSteps.colrngSht), Array(.mdlName))
        End If
    End With
    Exit Function
ErrorExit:
    errs.RecordErr "SetStepsRowRange", SetStepsRowRange
End Function
'---------------------------------------------------------------------------------------
'---------------------------------------------------------------------------------------
' TransferToMdlDest Procedure - transfer model from tblImport sheet to mdlDest region
'---------------------------------------------------------------------------------------
'---------------------------------------------------------------------------------------
' Transfer a model from tblImport sheet rows/cols "input deck" to mdlDest Scenario Model
'
' JDL 12/13/21   Modified: 8/22/23 JDL; 9/30/24 tblSteps instead of tblSteps
'                          10/29/25 use mdl.tblSteps attribute
'
Function TransferToMdlDest(mdlDest, tblImp, ModelNew, ModelDefnDest) As Boolean

    SetErrs TransferToMdlDest: If errs.IsHandle Then On Error GoTo ErrorExit
    Dim R_MI As New mdlImportRow
    
    With mdlDest
        
        'Clear previous mdlDest and set rng and rowCur in tblImport table
        If Not .InitTransferToMdl(mdlDest, tblImp, ModelNew) Then GoTo ErrorExit
        
        'Transfer ModelNew rows from tblImport table to mdlDest
        If Not .TransferTblImportRows(R_MI, mdlDest, tblImp) Then GoTo ErrorExit
        
        'Post-Transfer, delete ModelNew rows from tblImport table; Refresh mdlDest
        If Not .ResetPostTransfer(mdlDest, tblImp.rngRowsPopulated, ModelNew, _
                ModelDefnDest) Then GoTo ErrorExit
    End With
    Exit Function
    
ErrorExit:
    errs.RecordErr "TransferToMdlDest", TransferToMdlDest
End Function
'---------------------------------------------------------------------------------------
' Init transferring a model from tblImport sheet to Scenario Model
'
' JDL 7/25/23
'
Function InitTransferToMdl(ByVal mdlDest As mdlScenario, ByRef tblImp As tblRowsCols, _
        ByVal ModelNew As String)

    SetErrs InitTransferToMdl: If errs.IsHandle Then On Error GoTo ErrorExit
    
    'Clear mdlDest region and reset .rowCur
    With mdlDest
        .ClearModel mdlDest
        Set .rowCur = .cellHome.EntireRow
    End With
        
    'Set range for model in tblImport (.rngRowsPopulated attribute) - Err if not found
    With tblImp
        Set .rngRowsPopulated = rngKeycolRows(tblImp, .colrngMdlName, ModelNew)
        If errs.IsFail(.rngRowsPopulated Is Nothing, 1) Then GoTo ErrorExit
        Set .rowCur = .rngRowsPopulated.Rows(1)
    End With
    Exit Function
    
ErrorExit:
    errs.RecordErr "InitTransferToMdl", InitTransferToMdl
End Function
'---------------------------------------------------------------------------------------
' Transfer ModelNew rows from tblImport table to mdlDest
'
' JDL 7/25/23   Modified 8/24/23; 9/30/24 tblSteps instead of tblSteps
'                        10/29/25 use mdl.tblSteps attribute
'
Function TransferTblImportRows(ByRef R_MI As mdlImportRow, ByRef mdlDest As mdlScenario, _
        ByVal tblImp As tblRowsCols)

    SetErrs TransferTblImportRows: If errs.IsHandle Then On Error GoTo ErrorExit
    Dim r As Variant
    
    'Iterate over tblImp rows for model; write table vals to mdlDest and mdlDest.tblSteps
    For Each r In tblImp.rngRowsPopulated.Rows
        With R_MI
            If Not .Init(R_MI) Then GoTo ErrorExit
            If Not .ReadRow(R_MI, tblImp) Then GoTo ErrorExit
            If Not .SetBooleanFlags(R_MI) Then GoTo ErrorExit
            If Not .SetStepType(R_MI) Then GoTo ErrorExit
            If Not .WriteRowToMdl(R_MI, mdlDest) Then GoTo ErrorExit
            If Not .WriteRowToSteps(R_MI, mdlDest.tblSteps) Then GoTo ErrorExit
        End With
        Set mdlDest.rowCur = mdlDest.rowCur.Offset(1, 0)
        Set tblImp.rowCur = tblImp.rowCur.Offset(1, 0)
    Next r
    Exit Function
    
ErrorExit:
    errs.RecordErr "TransferTblImportRows", TransferTblImportRows
End Function
'---------------------------------------------------------------------------------------
' Post-Transfer, delete ModelNew rows from tblImport table; Refresh mdlDest
'
' JDL 7/25/23   Modified 8/22/23
'
Function ResetPostTransfer(ByRef mdlDest As mdlScenario, ByVal rngModel As Range, _
        ByVal ModelNew As String, ByVal ModelDefnDest As String) As Boolean

    SetErrs ResetPostTransfer: If errs.IsHandle Then On Error GoTo ErrorExit
    Dim wkbk As Workbook
    
    'Delete the model from tblImport sheet
    rngModel.Delete
    
    'Re-initialize mdlDest to clear mdlName and other params from pre-transfer
    Set wkbk = mdlDest.wkbk
    Set mdlDest = New mdlScenario
    
    'Re-provision/refresh dest model with customized model name
    With mdlDest
        If Not .Provision(mdlDest, wkbk, mdlName:=ModelNew, _
                Defn:=ModelDefnDest) Then GoTo ErrorExit
        If Not .Refresh(mdlDest) Then GoTo ErrorExit
        .ApplyBorderAroundModel mdlDest, IsBufferRow:=True, IsBufferCol:=True
    End With
    Exit Function
    
ErrorExit:
    errs.RecordErr "ResetPostTransfer", ResetPostTransfer
End Function