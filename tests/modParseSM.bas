Attribute VB_Name = "modParseSM"
Option Explicit
'------------------------------------------------------------------------------------------------
'Purpose: Create a rows/columns version of a scenario model (active sheet)
'
'Created: 9/23/21 JDL
'
Sub ParseScenarioModel()
    Dim wkbkSM As Workbook, wkbkRC As Workbook, sht As String
    Dim icolModel As Integer, icolEnd As Integer, irowEnd As Integer
    Dim s As Variant, i As Integer, rngTable As Range
        
    Set wkbkSM = ActiveWorkbook
    icolModel = 9
    
    'Exit if active sheet is not a Scenario Model
    If wkbkSM.ActiveSheet.Cells(1, 4) <> "Variable Names" Then Exit Sub
    sht = wkbkSM.ActiveSheet.Name
    
    With wkbkSM.Sheets(sht)
    
        'Find extent of scenario columns
        icolEnd = .Cells(2, .Columns.Count).End(xlToLeft).Column
        If icolEnd < icolModel Then Exit Sub

        'Find extent of variable rows
        irowEnd = .Cells(.Rows.Count, 4).End(xlUp).Row
        If irowEnd < 3 Then Exit Sub
        
        'Make a new workbook and Delete extraneous sheets
        Set wkbkRC = Workbooks.Add
        ActiveSheet.Name = sht
        For Each s In wkbkRC.Sheets
            If s.Name <> sht Then DeleteSheet wkbkRC, s
        Next s

        'Transpose paste Scenario Model to new workbook
        Range(.Cells(1, 1), .Cells(irowEnd, icolEnd)).Copy
        wkbkRC.Sheets(sht).Cells(1, 1).PasteSpecial _
            Paste:=xlPasteValuesAndNumberFormats, Transpose:=True
    End With
            
    'Clean up the rows/columns sheet
    With wkbkRC.Sheets(sht)
        Range(.Rows(1).EntireRow, .Rows(2).EntireRow).Delete
        
        'Delete unused columns
        icolEnd = .Cells(2, .Columns.Count).End(xlToLeft).Column
        For i = icolEnd To 2 Step -1
            If Len(.Cells(2, i)) < 1 Then .Columns(i).EntireColumn.Delete
        Next i
        icolEnd = .Cells(2, .Columns.Count).End(xlToLeft).Column
        
        'Move Description and units to comments; delete those rows
        For i = 3 To icolEnd
            s = .Cells(1, i).Value
            If Len(.Cells(3, i)) > 0 Then
                s = s & ", " & .Cells(3, i)
                AddComment .Cells(2, i), s
            End If
        Next i
        .Rows(1).EntireRow.Delete
        Range(.Rows(2).EntireRow, .Rows(5).EntireRow).Delete
        .Cells(1, 1) = "Scenario Description"
        
        'Delete unused rows
        irowEnd = .Cells(.Rows.Count, 2).End(xlUp).Row
        For i = irowEnd To 2 Step -1
            If Len(.Cells(i, 2)) < 1 Then .Rows(i).EntireRow.Delete
        Next i
        irowEnd = .Cells(.Rows.Count, 2).End(xlUp).Row
        
        'Set column widths
        Set rngTable = Range(.Columns(1).EntireColumn, .Columns(icolEnd).EntireColumn)
        rngTable.EntireColumn.ColumnWidth = 50
        rngTable.EntireColumn.AutoFit
        .Cells(1, 1).Select
    End With
End Sub



