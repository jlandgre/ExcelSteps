VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "mdlImportRow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Version 1/29/26
'         8/25/23 - Class for properties of tblImport sheet row and transfer to mdlDest
Option Explicit
Public HasFormula As Boolean
Public HasDropdown As Boolean
Public HasNumFmt As Boolean
Public HasStepsRow As Boolean
Public Step As String
Public rngStep As Range
Public IsInputVal As Boolean
Public IsNewGrp As Boolean
Public IsNewsubGrp As Boolean
Public IsBlankRow As Boolean

Public Model As String
Public rngModel As Range
Public Grp As String
Public GrpPrev As String
Public Subgrp As String
Public SubGrpPrev As String
Public Desc As String
Public varname As String
Public StrInput As String
Public NumFmt As String
Public Units As String
Public ScenName As String
Public value As Variant
Public DropdownLstName As String
'---------------------------------------------------------------------------------------------
'---------------------------------------------------------------------------------------------
' Methods used by both mdlScenario.TransferTblImportRows() Procedure
'                  and mdlScenario.TransferMdlDestRows() Procedure
'---------------------------------------------------------------------------------------------
'---------------------------------------------------------------------------------------------
' Set mdlImportRow attributes for previous Group and Subgroup
' JDL 8/21/23
'
Function Init(R_MI As mdlImportRow) As Boolean

    SetErrs Init: If errs.IsHandle Then On Error GoTo ErrorExit
    Dim GrpPrev As String, SubGrpPrev As String, Model
    
    'Store previous Group and Subgroup; re-initialize R_MI
    GrpPrev = R_MI.Grp
    SubGrpPrev = R_MI.Subgrp
    Model = R_MI.Model
    Set R_MI = New mdlImportRow
    
    'Set attributes for previous Group and Subgroup
    R_MI.GrpPrev = GrpPrev
    R_MI.SubGrpPrev = SubGrpPrev
    R_MI.Model = Model
    Exit Function
    
ErrorExit:
    errs.RecordErr "mdlImportRow.Init", Init
End Function
'---------------------------------------------------------------------------------------------
' Set Boolean flags describing a mdlDest row for either read from or write to tblImp
' JDL 8/21/23
'
Function SetBooleanFlags(R_MI As mdlImportRow) As Boolean

    SetErrs SetBooleanFlags: If errs.IsHandle Then On Error GoTo ErrorExit
    
    With R_MI
    
        '(Read from tblImp) If mdl row is to be blank, nothing else matters so exit
        If .varname = "<blank>" Then
            .IsBlankRow = True
            Exit Function
        End If
        
        'Determine StrInput (unused for write) is input value, formula or dropdown list
        If .StrInput = "Input" Then
            .IsInputVal = True
        ElseIf Left(.StrInput, 1) = "=" Then
            .HasFormula = True
        ElseIf Len(.StrInput) > 0 Then
            .HasDropdown = True
            .DropdownLstName = .StrInput
        End If
                
        'Number formatting for variable?
        If Len(.NumFmt) > 0 Then .HasNumFmt = True
        
        'ExcelSteps row for variable?
        .HasStepsRow = .HasFormula Or .HasDropdown Or .HasNumFmt
        
        'Did Group or Subgroup change versus previous row?
        If .Grp <> .GrpPrev Then .IsNewGrp = True
        If .Subgrp <> .SubGrpPrev Then .IsNewsubGrp = True
    End With
    Exit Function
ErrorExit:
    errs.RecordErr "SetBooleanFlags", SetBooleanFlags
End Function
'---------------------------------------------------------------------------------------------
'---------------------------------------------------------------------------------------------
' Methods used by mdlScenario.TransferMdlDestRows() Procedure
'---------------------------------------------------------------------------------------------
'---------------------------------------------------------------------------------------------
' Read a mdlDest row into tblImportRow attributes
' JDL 8/21/23
'
Function ReadMdlDestRow(R_MI As mdlImportRow, ByVal mdlDest As mdlScenario) As Boolean

    SetErrs ReadMdlDestRow: If errs.IsHandle Then On Error GoTo ErrorExit
    
    With mdlDest
        'Set flag and exit if row is empty
        If IsEmpty(Intersect(.rowCur, .colrngGrp)) And _
                IsEmpty(Intersect(.rowCur, .colrngVarNames)) Then
            R_MI.IsBlankRow = True
            R_MI.Grp = R_MI.GrpPrev
            Set .rowCur = .rowCur.Offset(1, 0)
            Exit Function
        End If

        R_MI.Grp = Intersect(.rowCur, .colrngGrp)
        If Len(R_MI.Grp) < 1 Then R_MI.Grp = R_MI.GrpPrev
        R_MI.Desc = Intersect(.rowCur, .colrngDesc)
        
        R_MI.varname = Intersect(.rowCur, .colrngVarNames)
        R_MI.Units = Intersect(.rowCur, .colrngUnits)
        R_MI.value = Intersect(.rowCur, .colrngModel)
                
        'Increment to next mdlDest row
        Set .rowCur = .rowCur.Offset(1, 0)
    End With
    Exit Function
ErrorExit:
    errs.RecordErr "ReadMdlDestRow", ReadMdlDestRow
End Function
'---------------------------------------------------------------------------------------------
' Read a mdlDest row's attributes from tblS ExcelSteps table
' JDL 8/21/23
'
Function ReadStepsRow(R_MI As mdlImportRow, ByVal tblS As tblRowsCols, _
        ByVal rngStepsVars As Range) As Boolean
        
    SetErrs ReadStepsRow: If errs.IsHandle Then On Error GoTo ErrorExit
    
    'Exit if no Steps rows for this Scenario Model or if mdlDest row is not a variable
    If rngStepsVars Is Nothing Or Len(R_MI.varname) < 1 Then Exit Function
    
    With R_MI
        Set .rngStep = FindInRange(Intersect(rngStepsVars, tblS.colrngCol), .varname)
        If .rngStep Is Nothing Then Exit Function
        
        'Set params for row's formula and number format
        .StrInput = Intersect(tblS.colrngStrInput, .rngStep.EntireRow)
        .NumFmt = Intersect(tblS.colrngNumFmt, .rngStep.EntireRow)
    End With
    Exit Function
    
ErrorExit:
    errs.RecordErr "ReadStepsRow", ReadStepsRow
End Function
'---------------------------------------------------------------------------------------------
' Write a mdlDest row to the tblImp rows/cols table on tblImport sheet
' JDL 8/21/23
'
Function ToTblWriteRow(ByVal R_MI As mdlImportRow, ByVal mdlDest As mdlScenario, _
        tblImp As tblRowsCols) As Boolean

    SetErrs ToTblWriteRow: If errs.IsHandle Then On Error GoTo ErrorExit
     Dim sVar As String
    
    'Skip writing if row is a group name or if group name has yet to be set at start
    If R_MI.IsNewGrp Or Len(R_MI.Grp) < 1 Then Exit Function
            
    With tblImp
        Intersect(.rowCur, .colrngMdlName) = mdlDest.mdlName
        Intersect(.rowCur, .colrngGrp) = R_MI.Grp
        Intersect(.rowCur, .colrngSubgrp) = R_MI.Subgrp
        Intersect(.rowCur, .colrngDesc) = R_MI.Desc
        
        'Write variable name or blank row marker
        sVar = R_MI.varname
        If R_MI.IsBlankRow Then sVar = "<blank>"
        Intersect(.rowCur, .colrngVarName) = sVar
        
        'Done writing if blank row
        If Not R_MI.IsBlankRow Then
            Intersect(.rowCur, .colrngUnits) = R_MI.Units
            Intersect(.rowCur, .colrngNumFmt) = R_MI.NumFmt
            
            'Write formula, dropdown list name or "Input"
            sVar = R_MI.StrInput
            If R_MI.IsInputVal Then sVar = "Input"
            Intersect(.rowCur, .colrngStrInput) = sVar
            
            'Write value if no formula
            If Not R_MI.HasFormula Then Intersect(.rowCur, .colrngValue) = R_MI.value
        End If
            
        'Increment to next row and add newly-written row to rngrows
        Set .rngRows = Union(.rngRows, .rowCur)
        Set .rowCur = .rowCur.Offset(1, 0)
    End With
    Exit Function
    
ErrorExit:
    errs.RecordErr "ToTblWriteRow", ToTblWriteRow
End Function
'---------------------------------------------------------------------------------------------
'---------------------------------------------------------------------------------------------
' Methods used by mdlScenario.TransferTblImportRows() Procedure
'---------------------------------------------------------------------------------------------
'---------------------------------------------------------------------------------------------
'Read tblImp rowCur values
'JDL 3/7/23 Modified: 8/24/23
'
Function ReadRow(R_MI As mdlImportRow, ByVal tblImp As tblRowsCols) As Boolean

    SetErrs ReadRow: If errs.IsHandle Then On Error GoTo ErrorExit
    
    With tblImp
        R_MI.Model = Intersect(.rowCur, .colrngMdlName)
        R_MI.Grp = Intersect(.rowCur, .colrngGrp)
        R_MI.Subgrp = Intersect(.rowCur, .colrngSubgrp)
        R_MI.Desc = Intersect(.rowCur, .colrngDesc)
        R_MI.varname = Intersect(.rowCur, .colrngVarName)
        R_MI.StrInput = Intersect(.rowCur, .colrngStrInput).Formula
        R_MI.NumFmt = Intersect(.rowCur, .colrngNumFmt)
        R_MI.Units = Intersect(.rowCur, .colrngUnits)
        R_MI.ScenName = Intersect(.rowCur, .colrngScenName)
        R_MI.value = Intersect(.rowCur, .colrngValue)
    End With
    Exit Function
    
ErrorExit:
    errs.RecordErr "ReadRow", ReadRow
End Function
'---------------------------------------------------------------------------------------------
' Set the type of ExcelSteps recipe step
'
' JDL 3/7/23 Modified: 8/24/23
'
Function SetStepType(R_MI As mdlImportRow) As Boolean

    SetErrs SetStepType: If errs.IsHandle Then On Error GoTo ErrorExit
    
    With R_MI
        .Step = "Col_Format"
        If .HasFormula Then
            .Step = "Col_Insert"
        ElseIf .HasDropdown Then
            .Step = "Col_Dropdown"
        End If
    End With
    Exit Function
    
ErrorExit:
    errs.RecordErr "SetStepType", SetStepType
End Function
'---------------------------------------------------------------------------------------------
' Write a tblImport table row to destination Scenario Model
'
'JDL 12/13/21   Modified: 8/24/23 JDL
'
Function WriteRowToMdl(ByVal R_MI As mdlImportRow, ByVal mdlDest As mdlScenario) As Boolean

    SetErrs WriteRowToMdl: If errs.IsHandle Then On Error GoTo ErrorExit
    With R_MI
    
        'If new Group or SubGroup, write that heading and increment mdlDest.rowCur
        If .Grp <> .GrpPrev Then
            If Not .WriteGrpToMdlDest(R_MI, mdlDest) Then GoTo ErrorExit
        ElseIf .Subgrp <> .SubGrpPrev Then
            If Not .WriteSubGrpToMdlDest(R_MI, mdlDest) Then GoTo ErrorExit
        End If
        
        'Write variable name, metadata and value
        If Not .IsBlankRow Then
            If Not .WriteVarToMdlDest(R_MI, mdlDest) Then GoTo ErrorExit
        End If
        Intersect(mdlDest.rowCur, mdlDest.colrngModel) = .value
    End With
    Exit Function
    
ErrorExit:
    errs.RecordErr "WriteRowToMdl", WriteRowToMdl
End Function
'---------------------------------------------------------------------------------------------
' Write new Grp name during Scenario Model transfer from tblImport table
' JDL 12/17/21; Updated 8/24/23
'
Function WriteGrpToMdlDest(ByVal R_MI As mdlImportRow, ByVal mdlDest As mdlScenario) As Boolean
    
    SetErrs WriteGrpToMdlDest: If errs.IsHandle Then On Error GoTo ErrorExit
    Dim c As Range
    
    With mdlDest
    
        'Write Grp to mdlDest rowCur
        Set c = Intersect(.rowCur, .colrngGrp)
        c.value = R_MI.Grp
        
        'Format the Grp string in mdlDest rowCur
        c.Font.Bold = True
        c.Font.Size = 12
        Set .rowCur = .rowCur.Offset(1, 0)
    End With
    Exit Function
    
ErrorExit:
    errs.RecordErr "WriteGrpToMdlDest", WriteGrpToMdlDest
End Function
'---------------------------------------------------------------------------------------------
' Write new Subgroup name during Scenario Model transfer from tblImport table
' JDL 3/7/23     Updated 7/25/23
'
Function WriteSubGrpToMdlDest(ByVal R_MI As mdlImportRow, ByVal mdlDest As mdlScenario) As Boolean
    
    SetErrs WriteSubGrpToMdlDest: If errs.IsHandle Then On Error GoTo ErrorExit
    
    With mdlDest
        Intersect(.rowCur, .colrngSubgrp) = R_MI.Subgrp
        Set .rowCur = .rowCur.Offset(1, 0)
    End With
    Exit Function
    
ErrorExit:
    errs.RecordErr "WriteSubGrpToMdlDest", WriteSubGrpToMdlDest
End Function
'---------------------------------------------------------------------------------------------
' Write variable's metadata to mdlDest Scenario Model as transfer from tblImport sheet
' JDL 3/7/23     Updated 7/25/23
'
Function WriteVarToMdlDest(ByVal R_MI As mdlImportRow, ByVal mdlDest As mdlScenario) As Boolean
    
    SetErrs WriteVarToMdlDest: If errs.IsHandle Then On Error GoTo ErrorExit
    With mdlDest
    
        'Write variable's Description, Variable Name and Units to mdlDest
        Intersect(.rowCur, .colrngDesc) = R_MI.Desc
        Intersect(.rowCur, .colrngVarNames) = R_MI.varname
        Intersect(.rowCur, .colrngUnits) = R_MI.Units
    End With
    Exit Function
    
ErrorExit:
    errs.RecordErr "WriteVarToMdlDest", WriteVarToMdlDest
End Function
'---------------------------------------------------------------------------------------------
' Write Lite Scenario Model variable's info to ExcelSteps row
' Reminder: InitSwapModels presets .rowCur to 1st blank row
'
' JDL 12/17/21; Refactored 3/6/23 for mdlImportRow Class
'
Function WriteRowToSteps(ByVal R_MI As mdlImportRow, tblS As tblRowsCols) As Boolean

    SetErrs WriteRowToSteps: If errs.IsHandle Then On Error GoTo ErrorExit

    If Not R_MI.HasStepsRow Then Exit Function
    With tblS
    
        'Write R_MI.Model and .VarName to Steps rowCur
        .SetTableLoc .rowCur, .colrngSht, R_MI.Model
        .SetTableLoc .rowCur, .colrngCol, R_MI.varname
        
        'Write Step type, Formula/Dropdown list name and Number Format
        .SetTableLoc .rowCur, .colrngStep, R_MI.Step
        If R_MI.HasFormula Or R_MI.HasDropdown Then _
            .SetTableLoc .rowCur, .colrngStrInput, R_MI.StrInput
        If R_MI.HasNumFmt Then .SetTableLoc .rowCur, .colrngNumFmt, R_MI.NumFmt

        'Extend .rngRows to include new row and increment Steps tbl.rowCur
        Set .rngRows = Range(.wkbk.Sheets(.sht).Rows(2), .rowCur)
        Set .rowCur = .rowCur.Offset(1, 0)
    End With
    Exit Function
    
ErrorExit:
    errs.RecordErr "WriteRowToSteps", WriteRowToSteps
End Function
