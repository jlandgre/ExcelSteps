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