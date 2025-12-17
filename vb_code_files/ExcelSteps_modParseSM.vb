'ExcelSteps_modParseSM.vb
'Version 12/10/25
Option Explicit
'------------------------------------------------------------------------------------------------------
'Create a rows/columns version of a scenario model
'
'Created: 9/23/21 JDL; Rewritten for updated architecture 11/19/25
'
Function ParseMdl(mdl) As Boolean
    SetErrs ParseMdl: If errs.IsHandle Then On Error GoTo ErrorExit
    Dim icolEnd As Integer, wkshtParsed As Worksheet, rowVarNames As Range, rowDelFlags As Range
    
    'Validate sheet containing multicolumn or calculator model
    If Not ValidateForParseMdl(mdl) Then GoTo ErrorExit
    
    'Create new workbook for parsed data
    If Not CreateWkbkParsedData(mdl) Then GoTo ErrorExit
    
    'Transpose paste model onto new workbook
    If Not TransposeScenarioMdl(mdl) Then GoTo ErrorExit
          
    Set wkshtParsed = mdl.wkbkParsed.Sheets(mdl.mdlName)
    
    'Set ranges and delete initial, unneeded rows and unused columns
    Set rowVarNames = wkshtParsed.Rows(3).EntireRow
    Set rowDelFlags = wkshtParsed.Rows(8).EntireRow
    If Not ParsedDataCleanup(mdl, rowVarNames, icolEnd) Then GoTo ErrorExit
    
    'Delete variables flagged with "d" in blank col to left of first mdl column
    If Not DeleteFlaggedVariables(mdl, rowDelFlags, rowVarNames, icolEnd) Then GoTo ErrorExit
            
    'Miscellaneous cleanup and Set column widths
    If Not MiscCleanup(mdl, rowVarNames, icolEnd) Then GoTo ErrorExit
    ColWidthAutofit wkshtParsed.Rows(1)
    wkshtParsed.Cells(1, 1).Select
    Exit Function
    
ErrorExit:
    errs.RecordErr "ParseMdl", ParseMdl
End Function
'------------------------------------------------------------------------------------------------------
'Validate scenario model for parsing
'Created: 9/23/21 JDL
'
Function ValidateForParseMdl(mdl) As Boolean
    SetErrs ValidateForParseMdl: If errs.IsHandle Then On Error GoTo ErrorExit
    Dim irow_scenario As Integer

        With mdl
            'Check that mdl is provisioned
            If errs.IsFail(.colrngVarNames Is Nothing, 1) Then GoTo ErrorExit
            
            'Check that model has both variable rows and scenario columns
            If errs.IsFail(.rngRows Is Nothing, 2) Then GoTo ErrorExit
            If errs.IsFail(.rngPopCols Is Nothing, 2) Then GoTo ErrorExit
        End With
    Exit Function
    
ErrorExit:
    errs.RecordErr "ValidateForParseMdl", ValidateForParseMdl
End Function
'------------------------------------------------------------------------------------------------------
'Create workbook for parsed data
'Created: 11/17/25 JDL; 12/10/25 Switch to DeleteExtraneousShts call
'
Function CreateWkbkParsedData(mdl) As Boolean
    SetErrs CreateWkbkParsedData: If errs.IsHandle Then On Error GoTo ErrorExit
    Dim s As Variant
    
    'Make a new workbook and Delete any extra sheets
    With mdl
        Set .wkbkParsed = Workbooks.Add
        ActiveSheet.Name = .mdlName
        If Not DeleteExtraneousShts(.wkbkParsed, .mdlName) Then GoTo ErrorExit
    End With
    Exit Function
    
ErrorExit:
    errs.RecordErr "CreateWkbkParsedData", CreateWkbkParsedData
End Function
'------------------------------------------------------------------------------------------------------
' Transpose paste Scenario Model subparts onto new workbook sheet
' 11/18/25 JDL
'
Function TransposeScenarioMdl(mdl) As Boolean
    SetErrs TransposeScenarioMdl: If errs.IsHandle Then On Error GoTo ErrorExit
    Dim icolEnd As Integer, irowEnd As Integer, icol_paste As Integer
    Dim rngSrc As Range, cellDest As Range
    Dim arySrc As Variant, aryDest As Variant, idx As Integer

        With mdl
        
            ' Copy/Transpose paste Variable Descs, Names, Units, Data and delete flags
            arySrc = Array(Intersect(.colrngDesc, .rngRows), Intersect(.colrngVarNames, .rngRows), _
                        Intersect(.colrngUnits, .rngRows), Intersect(.rngRows, .colrngModel), _
                        Intersect(.colrngModel.Columns(1).Offset(0, -1), .rngRows))
            With .wkbkParsed.Sheets(mdl.mdlName)
                aryDest = Array(.Cells(3, 2), .Cells(4, 2), .Cells(5, 2), .Cells(9, 2), .Cells(8, 2))
            End With
            
            For idx = 0 To UBound(arySrc)
                arySrc(idx).Copy
                aryDest(idx).PasteSpecial Paste:=xlPasteValuesAndNumberFormats, Transpose:=True
            Next idx
            
            'Copy/Transpose Paste Scenario descriptions (if header not suppressed)
            If Not .IsSuppHeader Then
                Set rngSrc = Intersect(.rngHeader.EntireRow, .colrngModel)
                Set cellDest = .wkbkParsed.Sheets(mdl.mdlName).Cells(9, 1)
                rngSrc.Copy
                cellDest.PasteSpecial Paste:=xlPasteValuesAndNumberFormats, Transpose:=True
           End If
        End With
    Exit Function
    
ErrorExit:
    errs.RecordErr "TransposeScenarioMdl", TransposeScenarioMdl
End Function
'------------------------------------------------------------------------------------------------------
'Cleanup rows/columns
'Created: 11/19/25 JDL
'
Function ParsedDataCleanup(mdl, rowVarNames, icolEnd) As Boolean
    SetErrs ParsedDataCleanup: If errs.IsHandle Then On Error GoTo ErrorExit
    Dim c As Range, i As Integer
    
    With mdl.wkbkParsed.Sheets(mdl.mdlName)
        Range(.Rows(1).EntireRow, .Rows(2).EntireRow).Delete
    
        'Delete unused columns (n total cols based on mdl.rngRows)
        For i = mdl.rngRows.Rows.Count + 1 To 2 Step -1
            Set c = Intersect(.Columns(i), rowVarNames)
            If IsEmpty(c) Then .Columns(i).EntireColumn.Delete
        Next i
        
        'Set column number of last parsed data column
        icolEnd = .Cells(rowVarNames.Row, .Columns.Count).End(xlToLeft).Column
    End With
    Exit Function
    
ErrorExit:
    errs.RecordErr "ParsedDataCleanup", ParsedDataCleanup
End Function
'------------------------------------------------------------------------------------------------------
'Create workbook for parsed data
'Created: 10/1/22; Moved to function 11/19/25 JDL
'
Function DeleteFlaggedVariables(mdl, rowDelFlags, rowVarNames, icolEnd) As Boolean
    SetErrs DeleteFlaggedVariables: If errs.IsHandle Then On Error GoTo ErrorExit
    Dim i As Integer
    
    'Delete flagged columns (10/21/22) - allows retaining subset
    With mdl.wkbkParsed.Sheets(mdl.mdlName)
        For i = icolEnd To 2 Step -1
            If .Cells(rowDelFlags.Row, i) = "d" Then .Columns(i).EntireColumn.Delete
        Next i
        
        'Reset icolEnd post deletions
        icolEnd = .Cells(rowVarNames.Row, .Columns.Count).End(xlToLeft).Column
    End With
    Exit Function
    
ErrorExit:
    errs.RecordErr "DeleteFlaggedVariables", DeleteFlaggedVariables
End Function
'------------------------------------------------------------------------------------------------------
'Miscellaneous cleanup
'Created: 11/19/25 JDL
'
Function MiscCleanup(mdl, rowVarNames, icolEnd) As Boolean
    SetErrs MiscCleanup: If errs.IsHandle Then On Error GoTo ErrorExit
    Dim i As Integer, s As String, irowEnd As Integer
    
    With mdl.wkbkParsed.Sheets(mdl.mdlName)
    
        'Ensure Calculator model has Scenario column
        If mdl.IsCalc Then
            If rowVarNames.Find("Scenario", lookat:=xlWhole) Is Nothing Then
                .Columns(2).Insert
                .Cells(2, 2).Value2 = "Scenario"
                icolEnd = icolEnd + 1
            End If
            If IsEmpty(.Cells(7, 2)) Then .Cells(7, 2).Value2 = "Calculator"
        End If
                
        'Move Description and units to comments; delete those rows
        For i = 2 To icolEnd
            s = .Cells(1, i).value
            If Len(.Cells(3, i)) > 0 Then
                s = s & ", " & .Cells(3, i)
            End If
            If Len(s) > 0 Then AddComment .Cells(2, i), s
        Next i
        .Rows(1).EntireRow.Delete
        Range(.Rows(2).EntireRow, .Rows(5).EntireRow).Delete
        
        'Label Scenario Description column (but delete blank col if header suppressed)
        .Cells(1, 1) = "Scenario Description"
        If mdl.IsSuppHeader Then .Columns(1).Delete
        
        'Delete unused rows (aka transposed mdl columns that don't contain scenario)
        irowEnd = .Cells(.Rows.Count, 2).End(xlUp).Row
        For i = irowEnd To 2 Step -1
            If Len(.Cells(i, 2)) < 1 Then .Rows(i).EntireRow.Delete
        Next i
    End With
    Exit Function
    
ErrorExit:
    errs.RecordErr "MiscCleanup", MiscCleanup
End Function

