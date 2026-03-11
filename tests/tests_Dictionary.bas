Attribute VB_Name = "tests_Dictionary"
Option Explicit
'Version 3/10/26
'--------------------------------------------------------------------------------------
' Dictionary Class Testing
Sub TestDriver_Dictionary()
    Dim procs As New Procedures, AllEnabled As Boolean
    
    With procs
        .Init procs, ThisWorkbook, "Tests_Dictionary", "Tests_Dictionary"
        SetApplEnvir False, False, xlCalculationManual
        
        'Enable testing of all or individual procedures
        AllEnabled = True
        .dictionary.Enabled = True
    End With
    
    With procs.dictionary
        If .Enabled Or AllEnabled Then
            procs.curProcedure = .Name
            test_Add procs
            test_Item procs
            test_Exists procs
            test_Remove procs
            test_Size procs
            test_GetKeys procs
            test_Clear procs
            test_UpdateValue procs
            test_ObjectValues procs
            test_ExpandCapacity procs
            test_NestedHelperAccess procs
            test_ValidateAndStripBraces procs
            test_SplitIntoPairs procs
            test_ParsePair procs
            test_DetectValueType procs
            test_AddParsedValue procs
            test_ParseStringToDictProcedure procs
        End If
    End With
    
    procs.EvalOverall procs
    SetApplEnvir True, True, xlCalculationAutomatic
End Sub
'--------------------------------------------------------------------------------------
'--------------------------------------------------------------------------------------
' procs.dictionary
'--------------------------------------------------------------------------------------
'--------------------------------------------------------------------------------------
'--------------------------------------------------------------------------------------
' Test adding new key-value pairs
' JDL 1/30/26
'
Sub test_Add(procs)
    Dim tst As New Test: tst.Init tst, "test_Add"
    Set ExcelSteps.errs = Nothing
    Dim dict As Object
    
    Set dict = ExcelSteps.New_Dictionary
    
    With tst
        'Add first item
        dict.Add "key1", "value1"
        .Assert tst, dict.Size = 1
        .Assert tst, dict.Item("key1") = "value1"
        
        'Add second item
        dict.Add "key2", "value2"
        .Assert tst, dict.Size = 2
        .Assert tst, dict.Item("key2") = "value2"
        
        'Add numeric value
        dict.Add "key3", 123
        .Assert tst, dict.Item("key3") = 123
        
        .Update tst, procs
    End With
End Sub
'--------------------------------------------------------------------------------------
' Test retrieving values by key
' JDL 1/30/26
'
Sub test_Item(procs)
    Dim tst As New Test: tst.Init tst, "test_Item"
    Set ExcelSteps.errs = Nothing
    Dim dict As Object
    
    Set dict = ExcelSteps.New_Dictionary
    
    With tst
        dict.Add "name", "John"
        dict.Add "age", 30
        dict.Add "city", "Boston"
        
        .Assert tst, dict.Item("name") = "John"
        .Assert tst, dict.Item("age") = 30
        .Assert tst, dict.Item("city") = "Boston"
        
        'Non-existent key returns Empty
        .Assert tst, IsEmpty(dict.Item("nonexistent"))
        
        .Update tst, procs
    End With
End Sub
'--------------------------------------------------------------------------------------
' Test checking if keys exist
' JDL 1/30/26
'
Sub test_Exists(procs)
    Dim tst As New Test: tst.Init tst, "test_Exists"
    Set ExcelSteps.errs = Nothing
    Dim dict As Object
    
    Set dict = ExcelSteps.New_Dictionary
    
    With tst
        .Assert tst, Not dict.Exists("key1")
        
        dict.Add "key1", "value1"
        .Assert tst, dict.Exists("key1")
        
        dict.Add "key2", "value2"
        .Assert tst, dict.Exists("key2")
        .Assert tst, dict.Exists("key1")
        
        .Assert tst, Not dict.Exists("key3")
        
        .Update tst, procs
    End With
End Sub
'--------------------------------------------------------------------------------------
' Test removing key-value pairs
' JDL 1/30/26
'
Sub test_Remove(procs)
    Dim tst As New Test: tst.Init tst, "test_Remove"
    Set ExcelSteps.errs = Nothing
    Dim dict As Object
    
    Set dict = ExcelSteps.New_Dictionary
    
    With tst
        dict.Add "key1", "value1"
        dict.Add "key2", "value2"
        dict.Add "key3", "value3"
        .Assert tst, dict.Size = 3
        
        'Remove middle item
        dict.Remove "key2"
        .Assert tst, dict.Size = 2
        .Assert tst, Not dict.Exists("key2")
        .Assert tst, dict.Exists("key1")
        .Assert tst, dict.Exists("key3")
        
        'Remove first item
        dict.Remove "key1"
        .Assert tst, dict.Size = 1
        .Assert tst, dict.Exists("key3")
        
        'Remove last item
        dict.Remove "key3"
        .Assert tst, dict.Size = 0
        
        .Update tst, procs
    End With
End Sub
'--------------------------------------------------------------------------------------
' Test item count tracking
' JDL 1/30/26
'
Sub test_Size(procs)
    Dim tst As New Test: tst.Init tst, "test_Size"
    Set ExcelSteps.errs = Nothing
    Dim dict As Object
    
    Set dict = ExcelSteps.New_Dictionary
    
    With tst
        .Assert tst, dict.Size = 0
        
        dict.Add "item1", 1
        .Assert tst, dict.Size = 1
        
        dict.Add "item2", 2
        dict.Add "item3", 3
        .Assert tst, dict.Size = 3
        
        dict.Remove "item2"
        .Assert tst, dict.Size = 2
        
        dict.Clear
        .Assert tst, dict.Size = 0
        
        .Update tst, procs
    End With
End Sub
'--------------------------------------------------------------------------------------
' Test retrieving all keys as array
' JDL 1/30/26
'
Sub test_GetKeys(procs)
    Dim tst As New Test: tst.Init tst, "test_GetKeys"
    Set ExcelSteps.errs = Nothing
    Dim dict As Object, keys() As String, i As Long
    
    Set dict = ExcelSteps.New_Dictionary
    
    With tst
        dict.Add "key1", "value1"
        dict.Add "key2", "value2"
        dict.Add "key3", "value3"
        
        keys = dict.GetKeys
        .Assert tst, UBound(keys) = 2
        .Assert tst, keys(0) = "key1"
        .Assert tst, keys(1) = "key2"
        .Assert tst, keys(2) = "key3"
        
        .Update tst, procs
    End With
End Sub
'--------------------------------------------------------------------------------------
' Test clearing all items
' JDL 1/30/26
'
Sub test_Clear(procs)
    Dim tst As New Test: tst.Init tst, "test_Clear"
    Set ExcelSteps.errs = Nothing
    Dim dict As Object
    
    Set dict = ExcelSteps.New_Dictionary
    
    With tst
        dict.Add "key1", "value1"
        dict.Add "key2", "value2"
        dict.Add "key3", "value3"
        .Assert tst, dict.Size = 3
        
        dict.Clear
        .Assert tst, dict.Size = 0
        .Assert tst, Not dict.Exists("key1")
        .Assert tst, Not dict.Exists("key2")
        .Assert tst, Not dict.Exists("key3")
        
        'Can add items after clear
        dict.Add "new", "item"
        .Assert tst, dict.Size = 1
        .Assert tst, dict.Item("new") = "item"
        
        .Update tst, procs
    End With
End Sub
'--------------------------------------------------------------------------------------
' Test updating existing key with new value
' JDL 1/30/26
'
Sub test_UpdateValue(procs)
    Dim tst As New Test: tst.Init tst, "test_UpdateValue"
    Set ExcelSteps.errs = Nothing
    Dim dict As Object
    
    Set dict = ExcelSteps.New_Dictionary
    
    With tst
        dict.Add "key1", "value1"
        .Assert tst, dict.Item("key1") = "value1"
        .Assert tst, dict.Size = 1
        
        'Update with new value
        dict.Add "key1", "updated_value"
        .Assert tst, dict.Item("key1") = "updated_value"
        .Assert tst, dict.Size = 1
        
        'Update with numeric value
        dict.Add "key1", 999
        .Assert tst, dict.Item("key1") = 999
        .Assert tst, dict.Size = 1
        
        .Update tst, procs
    End With
End Sub
'--------------------------------------------------------------------------------------
' Test storing and retrieving object references
' JDL 1/30/26
'
Sub test_ObjectValues(procs)
    Dim tst As New Test: tst.Init tst, "test_ObjectValues"
    Set ExcelSteps.errs = Nothing
    Dim dict As Object, rng As Range, wksht As Worksheet
    
    Set dict = ExcelSteps.New_Dictionary
    Set rng = tst.wkbkTest.Sheets(1).Range("A1")
    Set wksht = tst.wkbkTest.Sheets(1)
    
    With tst
        'Store range object
        dict.Add "range", rng
        .Assert tst, dict.Item("range").Address = "$A$1"
        
        'Store worksheet object
        dict.Add "sheet", wksht
        .Assert tst, dict.Item("sheet").Name = wksht.Name
        
        'Mix objects and values
        dict.Add "text", "hello"
        .Assert tst, dict.Item("text") = "hello"
        .Assert tst, dict.Size = 3
        
        .Update tst, procs
    End With
End Sub
'--------------------------------------------------------------------------------------
' Test capacity expansion when adding more than 16 items
' JDL 1/30/26
'
Sub test_ExpandCapacity(procs)
    Dim tst As New Test: tst.Init tst, "test_ExpandCapacity"
    Set ExcelSteps.errs = Nothing
    Dim dict As Object, i As Long, sKey As String
    
    Set dict = ExcelSteps.New_Dictionary
    
    With tst
        'Add 20 items to force expansion beyond initial capacity of 16
        For i = 1 To 20
            sKey = "key" & i
            dict.Add sKey, "value" & i
        Next i
        
        .Assert tst, dict.Size = 20
        
        'Verify all items still accessible
        .Assert tst, dict.Item("key1") = "value1"
        .Assert tst, dict.Item("key10") = "value10"
        .Assert tst, dict.Item("key20") = "value20"
        
        'Verify all keys exist
        .Assert tst, dict.Exists("key1")
        .Assert tst, dict.Exists("key16")
        .Assert tst, dict.Exists("key17")
        .Assert tst, dict.Exists("key20")
        
        .Update tst, procs
    End With
End Sub
'--------------------------------------------------------------------------------------
' Test accessing dictionary methods from nested helper function (See vba_class_arrays.md
' Skill for background on working with arrays as class attributes)
' JDL 2/6/26
'
Sub test_NestedHelperAccess(procs)
    Dim tst As New Test: tst.Init tst, "test_NestedHelperAccess"
    Set ExcelSteps.errs = Nothing
    Dim dict As Object
    
    Set dict = ExcelSteps.New_Dictionary
    
    With tst
        dict.Add "name", "Alice"
        dict.Add "age", 25
        dict.Add "city", "Boston"
        dict.Add "state", "MA"
        
        'Call helper function that accesses dict methods
        HelperAccessDictionary tst, dict
        
        .Update tst, procs
    End With
End Sub
'--------------------------------------------------------------------------------------
' Helper function to test accessing dictionary from nested function
' JDL 2/6/26
'
Sub HelperAccessDictionary(tst As Test, dict As Object)
    Dim keys() As String, i As Long, tempVal As Variant
    
    With tst
        'Test Item method from helper
        tempVal = dict.Item("name")
        .Assert tst, tempVal = "Alice"
        
        tempVal = dict.Item("age")
        .Assert tst, tempVal = 25
        
        tempVal = dict.Item("city")
        .Assert tst, tempVal = "Boston"
        
        'Test GetKeys method from helper
        keys = dict.GetKeys
        .Assert tst, UBound(keys) = 3
        .Assert tst, keys(0) = "name"
        .Assert tst, keys(1) = "age"
        .Assert tst, keys(2) = "city"
        .Assert tst, keys(3) = "state"
    End With
End Sub

'--------------------------------------------------------------------------------------
'--------------------------------------------------------------------------------------
' ParseStringToDictProcedure and related method tests
'--------------------------------------------------------------------------------------
'--------------------------------------------------------------------------------------
'--------------------------------------------------------------------------------------
' Test full ParseStringToDictProcedure integration behavior
' JDL 3/10/26
'
Sub test_ParseStringToDictProcedure(procs)
    Dim tst As New Test: tst.Init tst, "test_ParseStringToDictProcedure"
    Set ExcelSteps.errs = Nothing

    Dim dict As Object, jsonStr As String

    Set dict = ExcelSteps.New_Dictionary

    With tst
        dict.Add "keep", "orig"
        jsonStr = "{""name"":""Alice"", ""age"":25, ""active"":True, ""keep"":""updated""}"

        .Assert tst, dict.ParseStringToDictProcedure(jsonStr)
        .Assert tst, dict.Item("name") = "Alice"
        .Assert tst, dict.Item("age") = 25
        .Assert tst, dict.Item("active") = True
        .Assert tst, dict.Item("keep") = "updated"

        .Assert tst, Not dict.ParseStringToDictProcedure("{""bad"":oops}")

        .Update tst, procs
    End With
End Sub
'--------------------------------------------------------------------------------------
' Test ValidateAndStripBraces method
' JDL 3/10/26
'
Sub test_ValidateAndStripBraces(procs)
    Dim tst As New Test: tst.Init tst, "test_ValidateAndStripBraces"
    Set ExcelSteps.errs = Nothing
    Dim dict As Object, innerStr As String

    Set dict = ExcelSteps.New_Dictionary

    With tst
        .Assert tst, dict.ValidateAndStripBraces(" {""a"":1} ", innerStr)
        .Assert tst, innerStr = """a"":1"

        .Assert tst, dict.ValidateAndStripBraces("{}", innerStr)
        .Assert tst, innerStr = ""

        .Assert tst, Not dict.ValidateAndStripBraces("""a"":1", innerStr)
        .Assert tst, Not dict.ValidateAndStripBraces("{", innerStr)

        .Update tst, procs
    End With
End Sub
'--------------------------------------------------------------------------------------
' Test SplitIntoPairs method
' JDL 3/10/26
'
Sub test_SplitIntoPairs(procs)
    Dim tst As New Test: tst.Init tst, "test_SplitIntoPairs"
    Set ExcelSteps.errs = Nothing
    Dim dict As Object, pairs As Variant

    Set dict = ExcelSteps.New_Dictionary

    With tst
        .Assert tst, dict.SplitIntoPairs("""name"":""Last, First"",""age"":25", pairs)
        .Assert tst, UBound(pairs) = 1
        .Assert tst, pairs(0) = """name"":""Last, First"""
        .Assert tst, pairs(1) = """age"":25"

        .Assert tst, dict.SplitIntoPairs("", pairs)
        .Assert tst, IsArray(pairs)

        .Assert tst, Not dict.SplitIntoPairs("""a"":1,", pairs)

        .Update tst, procs
    End With
End Sub
'--------------------------------------------------------------------------------------
' Test ParsePair method
' JDL 3/10/26
'
Sub test_ParsePair(procs)
    Dim tst As New Test: tst.Init tst, "test_ParsePair"
    Set ExcelSteps.errs = Nothing
    Dim dict As Object, key As String, valParsed As Variant

    Set dict = ExcelSteps.New_Dictionary

    With tst
        .Assert tst, dict.ParsePair("""age"":25", key, valParsed)
        .Assert tst, key = "age"
        .Assert tst, valParsed = 25

        .Assert tst, dict.ParsePair("name:""Alice""", key, valParsed)
        .Assert tst, key = "name"
        .Assert tst, valParsed = "Alice"

        .Assert tst, Not dict.ParsePair("invalidpair", key, valParsed)

        .Update tst, procs
    End With
End Sub
'--------------------------------------------------------------------------------------
' Test DetectValueType method
' JDL 3/10/26
'
Sub test_DetectValueType(procs)
    Dim tst As New Test: tst.Init tst, "test_DetectValueType"
    Set ExcelSteps.errs = Nothing
    Dim dict As Object, valParsed As Variant

    Set dict = ExcelSteps.New_Dictionary

    With tst
        .Assert tst, dict.DetectValueType("""Alice""", valParsed)
        .Assert tst, valParsed = "Alice"

        .Assert tst, dict.DetectValueType("tRuE", valParsed)
        .Assert tst, valParsed = True

        .Assert tst, dict.DetectValueType("23.5", valParsed)
        .Assert tst, valParsed = 23.5

        .Assert tst, dict.DetectValueType("", valParsed)
        .Assert tst, IsEmpty(valParsed)

        .Assert tst, Not dict.DetectValueType("oops", valParsed)

        .Update tst, procs
    End With
End Sub
'--------------------------------------------------------------------------------------
' Test AddParsedValue method
' JDL 3/10/26
'
Sub test_AddParsedValue(procs)
    Dim tst As New Test: tst.Init tst, "test_AddParsedValue"
    Set ExcelSteps.errs = Nothing
    Dim dict As Object

    Set dict = ExcelSteps.New_Dictionary

    With tst
        .Assert tst, dict.AddParsedValue("k1", "v1")
        .Assert tst, dict.Item("k1") = "v1"

        .Assert tst, dict.AddParsedValue("k1", 99)
        .Assert tst, dict.Item("k1") = 99

        .Update tst, procs
    End With
End Sub

