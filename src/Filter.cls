VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Filter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Version 1/29/26
'11/18/21 add list filtering capability
Option Explicit
Public wkbk As Workbook
Public wksht As Worksheet
Public sht As String
Public IsShtFilter As Boolean
Public rngTable As Range
Public rngFilter As Range

'Use comma-separated lists because arrays problematic in VBA Classes
Public lstColIsFiltered As String
Public lstColFilterOperator As String
Public lstColFilCriteria1 As String
Public lstColFilCriteria2 As String

'Flag if any column filter is not AND/OR operator
Public IsListCriteria As Boolean
'-----------------------------------------------------------------------------------
' Captureexisting filtering
'
Sub CaptureExisting(cls, rngTable, rngHeader)
    With cls
        Set .rngTable = rngTable
        Set .wkbk = rngTable.Parent.Parent
        Set .wksht = rngTable.Parent
        .sht = rngTable.Parent.Name
        .IsShtFilter = .wksht.AutoFilterMode
        If .IsShtFilter Then Set .rngFilter = .wksht.AutoFilter.Range
        
        'Capture by-column parameters in comma-delimited lists
        RecordParams cls
    End With
End Sub
'-----------------------------------------------------------------------------------
' Populate Lists capturing filtering parameters for each table column
'
' Class uses matched-length lists to store filter parameters: .On, .Operator,
' .Criteria1 and .Criteria2
'
' Currently enables single criterion filter such as "=xxx" or ">1" and
' Two-criteria filters involving AND and OR.  Complex, list-based filters not
' supported, and .IsListCriteria attribute is set to True to flag such cases
'
' 11/10/21 JDL Modified: 11/18/21 - add list filtering capability
'
Sub RecordParams(cls)
    Dim i As Integer, nCols As Integer
    Dim aryIsFiltered As Variant, aryCrit1 As Variant
    Dim aryOper As Variant, aryCrit2 As Variant
    
    With cls
        If Not .IsShtFilter Then Exit Sub
        nCols = .rngFilter.Columns.Count
        
        'Default: no list filters (.Operator=7); All Operators are AND or OR
        .IsListCriteria = False
        
        'initialize local arrays and loop through columns to record params
        ReDim aryIsFiltered(1 To nCols)
        ReDim aryCrit1(1 To nCols)
        ReDim aryOper(1 To nCols)
        ReDim aryCrit2(1 To nCols)
        For i = 1 To nCols
        
            'Set default values (space characters enable conversion to lists)
            aryOper(i) = 0
            aryCrit1(i) = " "
            aryCrit2(i) = " "
            aryIsFiltered(i) = False
            
            'Is Column(i) filtered?
            If .wksht.AutoFilter.filters(i).On Then
                aryIsFiltered(i) = True
                
                '0=Single criterion; 1=Multi AND; 2=Multi OR; 7=List/complex
                aryOper(i) = .wksht.AutoFilter.filters(i).Operator
                
                'Set IsListCriteria flag if List-based filtering encountered
                If aryOper(i) > 2 Then .IsListCriteria = True
                
                'Capture criteria strings (example string: "=aa")
                If Not .IsListCriteria Then
                    aryCrit1(i) = .wksht.AutoFilter.filters(i).Criteria1
                    If aryOper(i) > 0 Then aryCrit2(i) = _
                        .wksht.AutoFilter.filters(i).Criteria2
                
                'Capture array of selected items as semicolon delimited list
                Else
                    aryCrit1(i) = _
                        ListFromArray(.wksht.AutoFilter.filters(i).Criteria1, sDelim:=";")
                End If
            End If
        Next i
        
        'for Class/VBA compatibility, convert arrays to comma-delimited lists
        .lstColIsFiltered = ListFromArray(aryIsFiltered)
        .lstColFilterOperator = ListFromArray(aryOper)
        .lstColFilCriteria1 = ListFromArray(aryCrit1)
        .lstColFilCriteria2 = ListFromArray(aryCrit2)
    End With
End Sub
'-----------------------------------------------------------------------------------
' Add saved filtering back to a table
' 11/11/21 JDL Modified: 11/18/21 - add list filtering capability
'
Sub AddToTable(cls)
    Dim i As Integer, nCols As Integer, aryIsFiltered As Variant
    Dim aryCrit1 As Variant, aryOper As Variant, aryCrit2 As Variant
    
    With cls

        'Convert parameter lists to arrays
        aryIsFiltered = Split(.lstColIsFiltered, ",")
        aryOper = Split(.lstColFilterOperator, ",")
        aryCrit1 = Split(.lstColFilCriteria1, ",")
        aryCrit2 = Split(.lstColFilCriteria2, ",")
        nCols = .rngFilter.Columns.Count
        
        'Loop through table columns and add specified filtering
        For i = 0 To nCols - 1
            If aryIsFiltered(i) Then
                
                '1=Multi AND; 2=Multi OR
                If aryOper(i) = 1 Or aryOper(i) = 2 Then
                    .rngFilter.AutoFilter field:=(i + 1), Criteria1:=aryCrit1(i), _
                        Operator:=val(aryOper(i)), Criteria2:=aryCrit2(i)
                        
                '0=Single criterion;
                ElseIf aryOper(i) = 0 Then
                    .rngFilter.AutoFilter field:=(i + 1), Criteria1:=aryCrit1(i)
                    
                '7=List of filtered values;
                ElseIf aryOper(i) = 7 Then
                    .rngFilter.AutoFilter field:=(i + 1), _
                    Criteria1:=Split(aryCrit1(i), ";"), Operator:=xlFilterValues
                End If
            End If
        Next i
    End With
End Sub

