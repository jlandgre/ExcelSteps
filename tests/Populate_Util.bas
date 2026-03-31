Attribute VB_Name = "Populate_Util"
'Tests_Populate_Util.vb
Option Explicit
'
'-----------------------------------------------------------------------------
'Simulate ExcelSteps recipe
'Modified 1/5/22
Sub PopulateSimRecipe(wkbk, sht)
    Dim rowCur As Long, i As Long

    'Turn off sheet's filter if on and clear cells
    With wkbk.Sheets(sht)
        .AutoFilterMode = False
        .Cells.Clear
        Range(.Columns(1), .Columns(20)).Delete
        
        'Write headers and set formula column format as text
        Range(.Cells(1, 1), .Cells(1, 4)) = _
            Split("Sheet,Column,Step,Formula/List Name/Sort-by", ",")
        .Columns(4).NumberFormat = "@"
            
        'Simulate a circa 50-row recipe
        rowCur = 2
        Range(.Cells(rowCur, 1), .Cells(rowCur + 4, 1)).Value2 = "aaa_bbb"
        Range(.Cells(rowCur + 2, 4), .Cells(rowCur + 4, 4)).Value2 = "=2"
        
        rowCur = 8
        Range(.Cells(rowCur, 1), .Cells(rowCur + 10, 1)).Value2 = "bbb_ccc"
        Range(.Cells(rowCur, 4), .Cells(rowCur + 10, 4)).Value2 = "=2"
        
        rowCur = 20
        Range(.Cells(rowCur, 1), .Cells(rowCur + 10, 1)).Value2 = "ccc_ddd"
        Range(.Cells(rowCur, 4), .Cells(rowCur + 10, 4)).Value2 = "=2"
        
        rowCur = 32
        Range(.Cells(rowCur, 1), .Cells(rowCur + 4, 1)).Value2 = "tgt_sheet"
        Range(.Cells(rowCur + 2, 4), .Cells(rowCur + 4, 4)).Value2 = "=2"
        For i = 32 To 36
            .Cells(i, 2).Value2 = "var_" & (i - 31)
        Next i

        rowCur = 40
        Range(.Cells(rowCur, 1), .Cells(rowCur + 10, 1)).Value2 = "tgt_sheet"
        Range(.Cells(rowCur + 2, 4), .Cells(rowCur + 4, 4)).Value2 = "=2"
        For i = 40 To 50
            .Cells(i, 2).Value2 = "var_" & (i - 39 + 5)
        Next i

        
        rowCur = 52
        Range(.Cells(rowCur, 1), .Cells(rowCur + 10, 1)).Value2 = "eee_fff"
        Range(.Cells(rowCur, 4), .Cells(rowCur + 10, 4)).Value2 = "=2"
        
        'Add Boolean to check if formula
        Range(.Cells(2, 5), .Cells(62, 5)).Formula = "=LEFT(D2,1)=""="""
    End With
End Sub

'-----------------------------------------------------------------------------------------------
'This section populates a table for testing rngMultiKey (4-key lookup)
'-----------------------------------------------------------------------------------------------
'Purpose: Populate lookup table
'
'Created:   12/15/21 JDL      Modified: 1/27/22 for rngMultiKey validation
'
Sub PopulateRngMultiKeyLookupTbl(wksht, cellHome)
    Dim ary As Variant, s As String, LstVals() As Variant, i As Integer
    Dim nRows As Integer, nCols As Integer
    
    'Clear previous
    s = cellHome.Address
    ClearTestSheetAndNames wksht
    Set cellHome = wksht.Range(s)
    nRows = 10
    nCols = 5
                
    'Populate lists of column values
    ReDim LstVals(1 To nCols)
    LstVals(1) = "Key_1,A,C,A,A,B,B,B,B,B"
    LstVals(2) = "Key_2,AA,AB,AA,AB,BC,BD,BC,BC,BD"
    LstVals(3) = "Key_3,X,X,Y,Y,X,X,X,Y,Y"
    LstVals(4) = "Key_4,,,,,1,,2,,"
    LstVals(5) = "Vals,1,2,3,4,5,7,6,8,9"
        
    'Populate by columns
    With cellHome
        For i = 1 To nCols
            ary = Split(LstVals(i), ",")
            Range(.Offset(0, i - 1), .Offset(nRows - 1, i - 1)) = _
                WorksheetFunction.Transpose(ary)
        Next i
    End With
End Sub
'-----------------------------------------------------------------------------
'This section populates values for testing LastPopulatedCell()
'-----------------------------------------------------------------------------
'Populate intermittent values in row
'Modified 1/5/22
Sub PopulateRow(wkbk, sht, icolStart)

    'Turn off sheet's filter if on and clear cells
    With wkbk.Sheets(sht)
        .AutoFilterMode = False
        .Cells.Clear
        Range(.Columns(1), .Columns(20)).Delete
        Range(.Cells(1, icolStart), .Cells(1, icolStart + 5)) = Array("a", "b", , "c", , "d")
    End With
End Sub
'Populate intermittent values in column
Sub PopulateCol(wkbk, sht, irowStart)

    'Turn off sheet's filter if on and clear cells
    RevealWkshtCells wkbk.Sheets(sht)
    With wkbk.Sheets(sht)
        .Cells.Clear
        Range(.Columns(1), .Columns(20)).Delete
        Range(.Cells(irowStart, 2), .Cells(irowStart + 5, 2)) = _
            Application.Transpose(Split("a,b,,c,,d", ","))
    End With
End Sub
'Populate intermittent values in column
Sub PopulateCol2(wkbk, sht)

    'Turn off sheet's filter if on and clear cells
    RevealWkshtCells wkbk.Sheets(sht)
    With wkbk.Sheets(sht)
        .Cells.Clear
        Range(.Columns(1), .Columns(20)).Delete
        Range(.Cells(2, 2), .Cells(2 + 5, 2)) = _
            Application.Transpose(Split("a,b,,c,,d", ","))
    End With
End Sub
'Populate block of values in column
Sub PopulateCol3(wkbk, sht)
    Dim nvals As Long, ary As Variant

    'Turn off sheet's filter if on and clear cells
    RevealWkshtCells wkbk.Sheets(sht)
    With wkbk.Sheets(sht)
        .Cells.Clear
        Range(.Columns(1), .Columns(20)).Delete
        ary = Split("a,b,c,d,e,f", ",")
        nvals = UBound(ary) + 1
        Range(.Cells(1, 1), .Cells(1 + nvals - 1, 1)) = Application.Transpose(ary)
    End With
End Sub
'Populate 2-D block values for FindInRange tests
Sub PopulateCells(wkbk, sht)
    Dim wksht As Worksheet

    Set wksht = wkbk.Sheets(sht)
    RevealWkshtCells wksht
    ClearTestSheetAndNames wksht

    With wksht
        .Range("A1:C1").Value2 = Split("Key,Val,Amount", ",")
        .Range("A2:C2").Value2 = Split("row1,c,10", ",")
        .Range("A3:C3").Value2 = Split("row2,d,20", ",")
        .Range("A4:C4").Value2 = Split("row3,e,30", ",")
    End With
End Sub


