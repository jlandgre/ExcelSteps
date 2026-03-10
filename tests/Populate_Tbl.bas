Attribute VB_Name = "Populate_Tbl"
Option Explicit
'-----------------------------------------------------------------------------------------------
'This section populates rows/columns tables
'-----------------------------------------------------------------------------------------------
'Prep ExcelSteps sheet and set Refresh Class attributes (see test_PrepExcelStepsSht)
'
'Created: 3/6/23 JDL; Modified 7/18/23 add IsReformat arg for refactored Prep method
'                              10/18/24 refactor for refr.InitTbl()
'
Sub PrepBlankStepsForTesting(wkbk, ByRef refr, ByRef tblSteps)
    Const sMsg  As String = "Sub PrepBlankStepsForTesting Error"
    Set refr = Excelsteps.New_Refresh
    Set tblSteps = Excelsteps.New_tbl
    
    'Initialize Refresh class and point it to test workbook
    With refr
        If Not .InitTbl(refr, wkbk, IsReplace:=True, IsTblFormat:=True) Then MsgBox sMsg
        
        'Clear previous ExcelSteps sheet if any
        If SheetExists(.wkbk, shtSteps) Then .wkbk.Sheets(shtSteps).Cells.Clear
        
        'ExcelSteps Prep function in Refresh Class
        If Not .PrepExcelStepsSht(refr, tblSteps, IsReformat:=True) Then MsgBox sMsg
    End With
End Sub

'-----------------------------------------------------------------------------------------------
'Purpose:   5 x 6 table
'
'Created:   12/14/21 JDL      Modified: 10/17/24 add rowHome and colHome args
'
Sub PopulateTbl2(wkbk, sht, Optional IsHeader = True, Optional IsData = True, _
        Optional rowHome = 2, Optional colHome = 1)
    Dim ary As Variant, s As String, LstVals() As Variant, i As Integer
    Dim wksht As Worksheet, nCols As Integer, nRows As Integer
    Set wksht = wkbk.Sheets(sht)
    
    With wkbk.Sheets(sht)
        ClearTestSheetAndNames wksht
        
        nCols = 6
        nRows = 5
        
        'Write Header
        If IsHeader Then
            s = "Desc,Desc2,Desc3,Data_1,Data_2,Data_3"
            Range(.Cells(rowHome - 1, colHome), .Cells(rowHome - 1, colHome + nCols - 1)) = Split(s, ",")
        End If
        
        'Populate lists of column values
        If IsData Then
            ReDim LstVals(1 To 9)
            LstVals(1) = "A,A,B,C,D"
            LstVals(2) = "DD,EE,CC,BB,AA"
            LstVals(3) = "BBB,HHH,EEE,FFF,GGG"
            LstVals(4) = "1.05,1.43,0.27,0.1,0.005"
            LstVals(5) = "27,22,34,19,18"
            LstVals(6) = "12.6,14.7,15.2,15.8,9.6"
            
            'Populate by columns
            For i = 1 To nCols
                ary = Split(LstVals(i), ",")
                Range(.Cells(rowHome, colHome + i - 1), .Cells(rowHome + nRows - 1, colHome + i - 1)) = WorksheetFunction.Transpose(ary)
            Next i
        End If
    End With
End Sub
'-----------------------------------------------------------------------------------------------
'Purpose:   basic 3 x 3 table
'
'Created:   12/14/21 JDL
'
Sub PopulateTbl(wkbk, sht)
    Dim ary As Variant, s As String, LstVals() As Variant, i As Integer
    Dim wksht As Worksheet
    Set wksht = wkbk.Sheets(sht)
    
    With wkbk.Sheets(sht)
        ClearTestSheetAndNames wksht
        
        'Write Header
        s = "Col_A,Col_B,Col_C"
        Range(.Cells(1, 1), .Cells(1, 3)) = Split(s, ",")
        
        'Populate lists of column values
        ReDim LstVals(1 To 9)
        LstVals(1) = "a,aa,aaa"
        LstVals(2) = "b,bb,bbb"
        LstVals(3) = "10,20,30"
        
        'Populate by columns
        For i = 1 To 3
            ary = Split(LstVals(i), ",")
            Range(.Cells(2, i), .Cells(4, i)) = WorksheetFunction.Transpose(ary)
        Next i
    End With
End Sub
'-----------------------------------------------------------------------------------------------
Sub PopulateStepsTblRefresh(wkbk, sht)
    With wkbk.Sheets(shtSteps)
        .Cells(2, 1) = sht
        .Cells(2, 2) = "Data_2"
        .Cells(2, 3) = "Col_Format"
        .Cells(2, 8) = "0.000"
        
        .Cells(3, 1) = sht
        .Cells(3, 2) = "Data_4"
        .Cells(3, 3) = "Col_Insert"
        .Cells(3, 4) = "=@Data_2 + @Data_3"
        .Cells(3, 5) = "Data_3"
        .Cells(3, 6) = True
        .Cells(3, 7) = "Calculated column"
        .Cells(3, 8) = "0.00"
        .Cells(3, 9) = 15
    End With
End Sub

