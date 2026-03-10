Attribute VB_Name = "Populate_ImportMdl"
Option Explicit
Sub PopulateDashAndMdlImportSht(Test)
    Dim mdlDash As Object, tblImp As Object
    
    'Populate model on sMdl and tbl on mdlImport sheet
    PopulateDashMdl Test, mdlDash
    PopulateList_Plants Test
    PopulateMdlImportType1AndType2 Test, tblImp
End Sub
'-----------------------------------------------------------------------------------------------
'Populate mdlImport sheet with Type1 and Type2 models
'JDL 7/17/23
'
Sub PopulateMdlImportType1AndType2(Test, tblImp)
    
    With Test
        .wkbkTest.Sheets(shtTblImp).Cells.Clear
        
        'Provision and format the tblImp sheet
        Set tblImp = Excelsteps.New_tbl
        .Assert Test, tblImp.Provision(tblImp, .wkbkTest, False, shtTblImp, nCols:=10, IsSetColRngs:=True)
        .Assert Test, tblImp.FormatMdlImport(tblImp)
        
        'Populate models onto mdlImport sheet
        PopulateMdlType1 tblImp
        PopulateMdlType2 tblImp
    End With

End Sub
' Populate Type1 model onto MdlImport
'
' Created: 1/3/22 JDL    Modified: 7/17/23 to use for more complex models
'
Sub PopulateMdlType1(tblImport)
    Dim aryRows As Variant
        
    'Specify each row as an array
    aryRows = Array( _
        "SMdlType1,Setup,,Configuration Name (used by program),mdl_name,,,Input,Calculator,SMdlType1", _
        "SMdlType1,Setup,,,<blank>,,,,,,,", _
        "SMdlType1,Batch Plant Configuration,,Batch Size,batch_size,kg,0,Input,Calculator,10000", _
        "SMdlType1,Batch Plant Configuration,,Use Premix,use_premix,kg,,Input,Calculator,True")
        AddMdlToMdlImport2 tblImport, aryRows
End Sub
' Populate Type2 model onto MdlImport
'
' Created: 1/3/22 JDL    Modified: 7/17/23 to use for more complex models
'                                  7/27/23 fix "SMdlType2"
'
Sub PopulateMdlType2(tblImport)
    Dim aryRows As Variant
        
    'Specify each row as an array
    aryRows = Array( _
        "SMdlType2,Setup,,Configuration Name (used by program),mdl_name,,,Input,Calculator,SMdlType2", _
        "SMdlType2,Setup,,,<blank>,,,,,,,", _
        "SMdlType2,Other Plant Configuration,,No. Sections,n_sections,,,Input,Calculator,4", _
        "SMdlType2,Other Plant Configuration,,Start Temperature (Celsius),T_start,C,,Input,Calculator,40", _
        "SMdlType2,Other Plant Configuration,,Start Temperature (Fahrenheit),T_start_f,F,0.0,=(T_start * 9/5) + 32,Calculator,")
    AddMdlToMdlImport2 tblImport, aryRows
End Sub
'-----------------------------------------------------------------------------------------------
'Purpose: Populate "Dashboard" Scenario model into sMdl sheet and Provision mdlDash
'
'Created:   12/15/21 JDL      Modified: 7/25/23
'
Sub PopulateDashMdl(Test, mdlDash)
    
    Dim ary As Variant, LstVals() As Variant, i As Integer, nRows As Integer
    
    ClearTestSheetAndNames Test.wkbkTest.Sheets("SMdl")
            
    'Populate lists of column values
    ReDim LstVals(1 To 6)
    LstVals(1) = "Dashboard,,"
    LstVals(2) = ",Plant,Dashboard 2"
    LstVals(3) = ",plant_name,dash_2"
    LstVals(4) = ",mm,mm"
    LstVals(5) = ",,,"
    LstVals(6) = ",Batch Plant,xxx"
    
    'Number of vals in each comma-separated LstVals string
    nRows = 3

    'Initialize mdlDash and Populate by columns
    Set mdlDash = Excelsteps.New_mdl
    With mdlDash
        Test.Assert Test, .Init(mdlDash, Test.wkbkTest, defn:=defn_dash)
    
        With .cellHome
            For i = 1 To UBound(LstVals)
                ary = WorksheetFunction.Transpose(Split(LstVals(i), ","))
                Range(.Offset(0, i - 1), .Offset(nRows - 1, i - 1)) = ary
            Next i
        End With
        
        PopulateList_Plants Test
        PopulateDashMdlStepsSht Test.wkbkTest
        
        'Provision and Refresh mdlDash
        Test.Assert Test, .Provision(mdlDash, .wkbk, defn:=defn_dash)
        Test.Assert Test, .Refresh(mdlDash)
    
        'Cosmetics
        .ApplyBorderAroundModel mdlDash, IsBufferCol:=True
        .wksht.Activate
        .cellHome.Select
    End With
End Sub
'-----------------------------------------------------------------------------------------------
'Create a named list of plants to use for list validation
'
'Created:   6/30/22 JDL     Modified 7/25/23
'
Sub PopulateList_Plants(Test)
    Dim ary As Variant, sNameString As String

    With Test.wkbkTest.Sheets("SMdl")
                 
        'Write list name and values to SMdl
        ary = Split("list_plants,Batch Plant,Other Plant", ",")
        Range(.Cells(1, 20), .Cells(3, 20)) = Application.Transpose(ary)
        
        'Name the list
        sNameString = MakeRefNameString("SMdl", 2, 3, 20, 20)
        MakeXLName Test.wkbkTest, xlName(.Cells(1, 20).Value), sNameString
        
        ShadeYellow Range(.Cells(2, 20), .Cells(3, 20))

    End With
End Sub
Sub PopulateDashMdlStepsSht(wkbk)
    Dim tbls As Object, refr As Object, lst As String

    'Recreate and populate ExcelSteps
    PrepBlankStepsForTesting wkbk, refr, tbls
    
    'Populate an instruction for plant_name dropdown list
    With tbls.wksht
        lst = "SMdlDash,plant_name,Col_Dropdown,list_plants"
        Range(.Cells(2, 1), .Cells(2, 4)) = Split(lst, ",")
    End With
End Sub
'-----------------------------------------------------------------------------
'This section populates models onto the MdlImport (rows/cols) sheet
'-----------------------------------------------------------------------------
' Utility to add a model to MdlImport in rows/columns format
'
' Created: 1/4/22 JDL   Modified: 7/17/23 to use for more complex models
'                                   2/13/25 for Mac compatibility CDbl instead of CDec
'
Sub AddMdlToMdlImport2(tblImport, aryRows)
    Dim cellCur As Range, sRow As Variant, val As Variant
    With tblImport.wksht
            
        'Find first blank row on mdlImport sheet
        Set cellCur = rngLastPopCell(.Cells(1, 1), xlDown).Offset(1, 0)
        
        'Write the rows and convert numerics
        For Each sRow In aryRows
            Range(cellCur, cellCur.Offset(0, 9)) = Split(sRow, ",")
            val = cellCur.Offset(0, 9)
            If IsNumeric(val) And Len(val) > 0 Then cellCur.Offset(0, 9) = CDbl(val)
            
            Set cellCur = cellCur.Offset(1, 0)
        Next sRow
    End With
End Sub
'Purpose: Utility to add a model to MdlImport in rows/columns format
'
'Created:   1/4/22 JDL      Modified:
'
Sub AddMdlToMdlImport(wkbk, aryMdlStrings)
    Dim c As Range, s As Variant
    With wkbk.Sheets(shtTblImp)

        'Refresh MdlImport sheet header cell values
        Range(.Cells(1, 1), .Cells(1, 9)) = Split(sHeaderMdlImport, ",")
        
        'Find last populated row and populate down from there
        Set c = rngLastPopCell(.Cells(1, 4), xlDown).Offset(1, 0)
        For Each s In aryMdlStrings
            Range(.Cells(c.Row, 1), .Cells(c.Row, 10)) = Split(s, ",")
            Set c = c.Offset(1, 0)
        Next s
    End With
End Sub
