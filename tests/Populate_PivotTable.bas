Attribute VB_Name = "Populate_PivotTable"
Option Explicit
'-----------------------------------------------------------------------------------------------
' Populate simple rows/columns data for PivotTable tests
' JDL 3/27/26
'
Sub PopulatePivotTableSimple(wkbk, sht)
    Dim wksht As Worksheet

    If Not SheetExists(wkbk, sht) Then AddSheet wkbk, sht, wkbk.Sheets(wkbk.Sheets.Count).Name
    Set wksht = wkbk.Sheets(sht)

    With wksht
        ClearTestSheetAndNames wksht

        .Range("A1:C1").Value2 = Split("Category,SubCategory,Amount", ",")

        .Range("A2:A7").Value2 = WorksheetFunction.Transpose(Split("A,A,A,B,B,B", ","))
        .Range("B2:B7").Value2 = WorksheetFunction.Transpose(Split("X,Y,X,X,Y,Y", ","))
        .Range("C2:C7").Value2 = WorksheetFunction.Transpose(Split("10,20,5,7,3,2", ","))
    End With
End Sub

'-----------------------------------------------------------------------------------------------
' Populate OTB-like rows/columns data with two column variables and three analytes
' JDL 3/30/26
'
Sub PopulatePivotTableOTBLike(wkbk, sht)
    Dim wksht As Worksheet

    If Not SheetExists(wkbk, sht) Then AddSheet wkbk, sht, wkbk.Sheets(wkbk.Sheets.Count).Name
    Set wksht = wkbk.Sheets(sht)

    With wksht
        ClearTestSheetAndNames wksht

        .Range("A1:G1").Value2 = Split("Store,Prodtype,Week,Year,Discounts,Markdowns,COGS", ",")

        .Range("A2:A13").Value2 = WorksheetFunction.Transpose(Split("Store1,Store1,Store1,Store1,Store1,Store1,Store2,Store2,Store2,Store2,Store2,Store2", ","))
        .Range("B2:B13").Value2 = WorksheetFunction.Transpose(Split("X,Y,Z,X,Y,Z,X,Y,Z,X,Y,Z", ","))
        .Range("C2:C13").Value2 = WorksheetFunction.Transpose(Split("1,1,1,2,2,2,1,1,1,2,2,2", ","))
        .Range("D2:D13").Value2 = WorksheetFunction.Transpose(Split("2025,2025,2025,2025,2025,2025,2025,2025,2025,2025,2025,2025", ","))
        .Range("E2:E13").Value2 = WorksheetFunction.Transpose(Split("10,20,30,11,21,31,40,50,60,41,51,61", ","))
        .Range("F2:F13").Value2 = WorksheetFunction.Transpose(Split("100,200,300,101,201,301,400,500,600,401,501,601", ","))
        .Range("G2:G13").Value2 = WorksheetFunction.Transpose(Split("1000,2000,3000,1001,2001,3001,4000,5000,6000,4001,5001,6001", ","))
    End With
End Sub


