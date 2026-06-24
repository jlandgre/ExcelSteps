VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "mdlRow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Version 6/24/26
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

Public IsBlockVar As Boolean
Public IsInputVarBlock As Boolean
Public IsFormulaVarBlock As Boolean
Public nBlockVars As Long
Public sBaseVar As String
Public sVarStem As String
Public sFormulaBase As String
Public sStep As String
Public sComment As String
Public rngBlockRows As Range
Public rngRecipeRow As Range
Public nExistingBlockVars As Long
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
'Initialize block variable attributes from an ExcelSteps recipe row
'JDL 6/24/26; Updated 6/24/26
'
Function InitBlockVar(r, mdl, rngRecipeRow) As Boolean
    SetErrs InitBlockVar: If errs.IsHandle Then On Error GoTo ErrorExit
    Dim paramsDict As Object, sParams As String

    With r
        Set .rngRecipeRow = rngRecipeRow
        .IsBlockVar = True
        .sBaseVar = Intersect(rngRecipeRow, mdl.tblSteps.colrngCol).Value2
        .sStep = Intersect(rngRecipeRow, mdl.tblSteps.colrngStep).Value2
        .sFormulaBase = Intersect(rngRecipeRow, mdl.tblSteps.colrngStrInput).Formula
        .NumFmt = Intersect(rngRecipeRow, mdl.tblSteps.colrngNumFmt).Value2
        .sComment = Intersect(rngRecipeRow, mdl.tblSteps.colrngComment).Value2
        .IsInputVarBlock = (.sStep = sAInputVarBlock)
        .IsFormulaVarBlock = (.sStep = sAFormulaVarBlock)

        If errs.IsFail(Not .IsInputVarBlock And Not .IsFormulaVarBlock, 1) Then GoTo ErrorExit
        If errs.IsFail(Right(.sBaseVar, 3) <> "_xx", 2, .sBaseVar) Then GoTo ErrorExit

        sParams = Intersect(rngRecipeRow, mdl.tblSteps.colrngParams).Value2
        If errs.IsFail(Len(sParams) = 0, 3, .sBaseVar) Then GoTo ErrorExit
        Set paramsDict = New Dictionary
        If Not paramsDict.ParseStringToDictProcedure(sParams) Then GoTo ErrorExit
        If errs.IsFail(Not paramsDict.Exists("nBlockVars"), 4, sParams) Then GoTo ErrorExit
        If errs.IsFail(paramsDict.Item("nBlockVars") <> CLng(paramsDict.Item("nBlockVars")), _
            5, sParams) Then GoTo ErrorExit
        .nBlockVars = CLng(paramsDict.Item("nBlockVars"))
        If errs.IsFail(.nBlockVars < 1, 6, sParams) Then GoTo ErrorExit
        If errs.IsFail(.IsFormulaVarBlock And Len(.sFormulaBase) = 0, 7, .sBaseVar) Then _
            GoTo ErrorExit

        .sVarStem = Left(.sBaseVar, Len(.sBaseVar) - 3)
        .HasNumFmt = Len(.NumFmt) > 0
        .HasFormula = .IsFormulaVarBlock
    End With
    Exit Function

ErrorExit:
    errs.RecordErr "InitBlockVar", InitBlockVar
End Function
'-----------------------------------------------------------------------------------------------------
'Return indexed block variable name from a base _xx variable
'JDL 6/24/26; Updated 6/24/26
'
Function IndexedVarName(r, idx As Long) As String
    SetErrs "non-bool": If errs.IsHandle Then On Error GoTo ErrorExit

    IndexedVarName = r.sVarStem & "_" & idx
    Exit Function

ErrorExit:
    errs.RecordErr "IndexedVarName"
End Function
'-----------------------------------------------------------------------------------------------------
'Return indexed block formula from a _xx formula template
'JDL 6/24/26; Updated 6/24/26
'
Function IndexedFormula(r, idx As Long) As String
    SetErrs "non-bool": If errs.IsHandle Then On Error GoTo ErrorExit

    IndexedFormula = Replace(r.sFormulaBase, "_xx", "_" & idx)
    Exit Function

ErrorExit:
    errs.RecordErr "IndexedFormula"
End Function
'-----------------------------------------------------------------------------------------------------
'Apply existing mdlRow refresh behavior to every row in an indexed block
'JDL 6/24/26; Updated 6/24/26
'
Function RefreshVarBlock(r, mdl) As Boolean
    SetErrs RefreshVarBlock: If errs.IsHandle Then On Error GoTo ErrorExit
    Dim rngVar As Range, rIdx As mdlRow, idx As Long

    idx = 1
    For Each rngVar In r.rngBlockRows
        Set rIdx = New mdlRow
        rIdx.Init rIdx, mdl, rngVar
        rIdx.sVar = r.IndexedVarName(r, idx)
        rIdx.HasFormula = r.IsFormulaVarBlock
        If rIdx.HasFormula Then rIdx.sFormula = r.IndexedFormula(r, idx)
        If r.HasNumFmt Then
            rIdx.HasNumFmt = True
            rIdx.NumFmt = r.NumFmt
        End If
        If mdl.IsRngNames Then rIdx.NameRow rIdx, mdl
        If Not rIdx.rngMdlCells Is Nothing Then
            If rIdx.HasFormula Then If Not rIdx.WriteRowFormula(rIdx, mdl) Then GoTo ErrorExit
            rIdx.FormatRow rIdx
        End If
        idx = idx + 1
    Next rngVar
    Exit Function

ErrorExit:
    errs.RecordErr "RefreshVarBlock", RefreshVarBlock
End Function
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

