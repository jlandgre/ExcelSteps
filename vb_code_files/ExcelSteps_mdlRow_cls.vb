'ExcelSteps_mdlRow_cls.vb
'Version 10/29/25 Performance optimization
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
'Write formula in sVar row's populated column cells
'
'JDL 6/30/22    Modified 10/30/25 With.rngMdlCells is significant performance improvement
'
Function WriteRowFormula(r, mdl) As Boolean
    SetErrs WriteRowFormula: If errs.IsHandle Then On Error GoTo ErrorExit
    Const Locn As String = "WriteRowFormula"
    With r.rngMdlCells
        
        'Write the formula to model column(s)
        On Error Resume Next
        .Value2 = r.sFormula
        
        'If Excel formula syntax error, flag with comment and exit
        If errs.IsFail(err.Number = 1004, 1) Then
            .Clear
            GoTo FlagWithComment
        End If
        
        '5/6/25 change color to black but keep bold and background from Calculation
        .Style = "Calculation"
        .Font.ColorIndex = xlAutomatic
    End With
    Exit Function
    
FlagWithComment:
    errs.LookupCommentMsg Intersect(r.rngVarRow, mdl.colrngVarNames), Locn, False
ErrorExit:
    errs.RecordErr Locn, WriteRowFormula
End Function
'-----------------------------------------------------------------------------------------------------
'Format row's model cells
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
'1/6/22 JDL; 10/29/25 refactor for performance
'
Sub NameRow(r, mdl)
    Dim sName As String, sRefersTo As String
    With mdl
        If .IsCalc Then
            sRefersTo = .NameRefPrefixCell & r.iRow & "C" & .icolCalc
        Else
            sRefersTo = .NameRefPrefix & r.iRow & ":R" & r.iRow
        End If
        sName = .NamePrefix & r.sVar
        .wkbk.Names.Add Name:=sName, RefersToR1C1:=sRefersTo
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

