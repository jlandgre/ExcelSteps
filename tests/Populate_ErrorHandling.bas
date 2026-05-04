Attribute VB_Name = "Populate_ErrorHandling"
Option Explicit
'Version 3/13/26
'--------------------------------------------------------------------------------------
' Populate mock Errors_ table for ErrorHandling tests
' JDL 3/13/26
'
Sub Populate_Errs_Default()
    Dim wksht As Worksheet

    If Not SheetExists(ExcelSteps.errs.wkbkE, ExcelSteps.shtErrors) Then _
        AddSheet ExcelSteps.errs.wkbkE, ExcelSteps.shtErrors, _
            ExcelSteps.errs.wkbkE.Sheets(ExcelSteps.errs.wkbkE.Sheets.Count).name

    Set wksht = ExcelSteps.errs.wkbkE.Sheets(ExcelSteps.shtErrors)

    With wksht
        .Cells.Clear

        'Set cleaned Errors_ headers
        .Range("A1:F1").Value = Split("iCodeReport,Module,Routine,Message,IsUserFacing,VBAProject", ",")

        'Base and detail rows for TestProc
        .Cells(2, 1).Value = 100
        .Cells(2, 2).Value = "Utilities"
        .Cells(2, 3).Value = "TestProc"
        .Cells(2, 4).Value = ExcelSteps.sErrBase
        .Cells(2, 5).Value = False
        .Cells(2, 6).Value = "ExcelSteps"

        .Cells(3, 1).Value = 101
        .Cells(3, 2).Value = "Utilities"
        .Cells(3, 3).Value = "TestProc"
        .Cells(3, 4).Value = "User visible: "
        .Cells(3, 5).Value = True
        .Cells(3, 6).Value = "ExcelSteps"

        .Cells(4, 1).Value = 102
        .Cells(4, 2).Value = "Utilities"
        .Cells(4, 3).Value = "TestProc"
        .Cells(4, 4).Value = "Developer detail: "
        .Cells(4, 5).Value = False
        .Cells(4, 6).Value = "ExcelSteps"

        'Malformed row path (invalid IsUserFacing)
        .Cells(5, 1).Value = 201
        .Cells(5, 2).Value = "Utilities"
        .Cells(5, 3).Value = "BadProc"
        .Cells(5, 4).Value = ""
        .Cells(5, 5).Value = "maybe"
        .Cells(5, 6).Value = "ExcelSteps"
    End With
End Sub


