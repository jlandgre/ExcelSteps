'mdlRow_cls.vb
'Version 10/29/25
Option Explicit

Public rngVar As Range
Public rngVarRow As Range
Public rngMdlCells As Range
Public iRow As Integer
Public sVar As String

Public HasFormula As Boolean
Public sFormula As String

Public HasNumFmt As Boolean
Public NumFmt As String

Public HasLstValidation As Boolean
Public DropdownLstName As String
'-----------------------------------------------------------------------------------------------------
'Purpose: Initialize the class instance by reading all attributes
'
'JDL 6/30/22; Modified 7/17/25; 10/29/25 use mdl.tblSteps attribute; NameRefPrefix from mdl
'
Sub Init(r, mdl, rngVar)
    SetRowProps r, mdl, rngVar
    
    If mdl.IsLiteModel Then
        r.ReadPropsLite r, mdl
    Else
        r.ReadPropsNonLite r, mdl
    End If
End Sub
'-----------------------------------------------------------------------------------------------------
'Purpose: Set properties related to locations within variable's row
'
'JDL 6/30/22
'
Sub SetRowProps(ByRef r, mdl, rngVar)
    Dim s As String
    With r
        Set .rngVar = rngVar
        .sVar = .rngVar.Value2
        Set .rngVarRow = .rngVar.EntireRow
        .iRow = .rngVar.Row
        
        'set multirange for populated columns
        If mdl.rngPopCols Is Nothing Then Exit Sub
        Set .rngMdlCells = Intersect(.rngVarRow, mdl.rngPopCols.EntireColumn)
    End With
End Sub
'-----------------------------------------------------------------------------------------------------
'Purpose: Read/set properties from the row's column header region
'
'JDL 6/30/22
'
Sub ReadPropsNonLite(ByRef r, mdl)
    Dim s As String
    With mdl
        r.NumFmt = Intersect(r.rngVarRow, .colrngNumFmt).Value2
        If Len(r.NumFmt) > 0 Then r.HasNumFmt = True
        
        r.sFormula = Intersect(r.rngVarRow, .colrngFormulas).Value2
        If Left(r.sFormula, 1) = "=" Then
            r.HasFormula = True
        Else
            'Dropdown instruction is in Formula/Row Type header cell
            s = Intersect(r.rngVarRow, .colrngFormulas).Value2
            If InStr(LCase(s), "dropdown") > 0 Then
                r.DropdownLstName = Split(s, ":")(1)
                If Len(r.DropdownLstName) > 0 Then r.HasLstValidation = True
            End If
        End If
    End With
End Sub
'-----------------------------------------------------------------------------------------------------
'Set step attributes by reading from dictionaries populated in SetRngFormulaRows
'JDL 10/29/25
'
Sub ReadPropsLite(ByRef r, mdl)
    
    'Exit if mdl is not on ExcelSteps (e.g. no recipe)
    If mdl.rngStepsVars Is Nothing Then Exit Sub
    
    'Read properties from dictionaries if variable exists
    If mdl.dStepsFormulas.Exists(r.sVar) Then
        r.HasFormula = True
        r.sFormula = mdl.dStepsFormulas.Item(r.sVar)
    End If
    
    If mdl.dStepsNumFormats.Exists(r.sVar) Then
        r.HasNumFmt = True
        r.NumFmt = mdl.dStepsNumFormats.Item(r.sVar)
    End If
    
    If mdl.dDropdownList.Exists(r.sVar) Then
        r.HasLstValidation = True
        r.DropdownLstName = mdl.dDropdownList.Item(r.sVar)
    End If
End Sub
'-----------------------------------------------------------------------------------------------------
'Set step attributes by reading from ExcelSteps (previous version)
'JDL 6/30/22; Modified 5/9/25 to exit if mdl not on ExcelSteps; 10/20/25 Eliminate Or for HasFormula
' 10/27/25 switch to .Find for performance versus .FindInrange
'
Sub ReadPropsLite_prev(ByRef r, mdl, tblSteps)
    Dim sVal As String, sInstruction As String, rngRowSteps As Range
    
    ' Exit if mdl is not on ExcelSteps (e.g. no recipe)
    If mdl.rngStepsVars Is Nothing Then Exit Sub
    
    'Read the instruction from ExcelSteps; no properties to set if R.sVar not on Steps
    'Set rngRowSteps = FindInRange(mdl.rngStepsVars, r.sVar)
    Set rngRowSteps = mdl.rngStepsVars.Find(r.sVar, lookat:=xlWhole)
    If rngRowSteps Is Nothing Then Exit Sub
    
    'Read instruction
    sInstruction = Intersect(rngRowSteps.EntireRow, tblSteps.colrngStep)
    sVal = Intersect(rngRowSteps.EntireRow, tblSteps.colrngStrInput)
    
    'Number format
    r.NumFmt = Intersect(rngRowSteps.EntireRow, tblSteps.colrngNumFmt)
    If Len(r.NumFmt) > 0 Then r.HasNumFmt = True
    
    'HasFormula based on "=" first character or Col_Insert Step
    'If Left(sVal, 1) = "=" Or LCase(sInstruction) = LCase(sAInsert) Then
    If Left(sVal, 1) = "=" Then
        r.HasFormula = True
        r.sFormula = sVal
        
    'Instruction determins List Validation (refactor to allow other validation criteria)
    ElseIf LCase(sInstruction) = "col_dropdown" Then
        r.HasLstValidation = True
        r.DropdownLstName = sVal
    Else
        'unsupported instruction/ExcelSteps syntax error
    End If
End Sub

'-----------------------------------------------------------------------------------------------------
'Write formula in sVar row's populated column cells
'
'JDL 6/30/22    Modified 5/6/25
'
Function WriteRowFormula(r, mdl) As Boolean
    SetErrs WriteRowFormula: If errs.IsHandle Then On Error GoTo ErrorExit
    Const Locn As String = "WriteRowFormula"
    With r
    
        'Proceed if there's a formula to write
        If Not .HasFormula Then Exit Function
    
        'Write the formula to model column(s)
        On Error Resume Next
        .rngMdlCells.Formula = .sFormula
        
        'If Excel formula syntax error, flag with comment and exit
        If errs.IsFail(err.Number = 1004, 1) Then
            .rngMdlCells.Clear
            GoTo FlagWithComment
        End If
        
        .rngMdlCells.Style = "Calculation"
        
        '5/6/25 change color to black but keep bold and background from Calculation
        rngMdlCells.Font.ColorIndex = xlAutomatic

    End With
    Exit Function
    
FlagWithComment:
    errs.LookupCommentMsg Intersect(r.rngVarRow, mdl.colrngVarNames), Locn, False
ErrorExit:
    errs.RecordErr Locn, WriteRowFormula
End Function

'-----------------------------------------------------------------------------------------------------
'Purpose: Format row's model cells
'
'JDL 6/30/22
'
Sub FormatRow(r)
    With r
        If .HasNumFmt Then .rngMdlCells.NumberFormat = .NumFmt
        SetBorders .rngMdlCells, xlContinuous, True
    End With
End Sub
'-----------------------------------------------------------------------------------------------------
'Purpose: Name row or row's calculator cell
'
'1/6/22 JDL; 7/18/25 refactor for performance
'
Sub NameRow(r, mdl)
    Dim sName As String
    With mdl
        sName = .NamePrefix & r.sVar
        If .IsCalc Then
            MakeXLName .wkbk, sName, .NameRefPrefixCell & r.iRow & "C" & .icolCalc
            'MakeXLName .wkbk, sName, MakeRefNameString(.sht, r.iRow, r.iRow, .icolCalc, .icolCalc)
        Else
            .wkbk.Names.Add Name:=sName, RefersToR1C1:=.NameRefPrefix & r.iRow & ":R" & r.iRow
        End If
    End With
End Sub
'------------------------------------------------------------------------------------------------------
'Purpose: Add a dropdown to model cells in row
'
'Created:   6/30/22 JDL
'
Sub AddListValidation(r)
    AddValidationList r.rngMdlCells, "=" & r.DropdownLstName
End Sub