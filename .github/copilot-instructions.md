# Dashboard VBA Project - AI Coding Instructions
updated 11/11/25
## Project Architecture

Projects built in VBA have cross-platform compatibility (Windows/Mac Excel). A typical project consists of:
- **ProjectName.xlsm** - Main project workbook (VBA Project: `VBAProject_ProjectName`)
- **XLSteps.xlam** - ExcelSteps add-in (VBA Project: `ExcelSteps`)
- **Tests_ProjectName.xlsm** - Unit test suite workbook (VBA Project: `VBAProject_Tests`)
`VBAProject_ProjectName` has `ExcelSteps` as a reference.  `VBAProject_Tests` has `VBAProject_ProjectName` as a reference. We assume comprehensive unit testing in the test suite which contains one or more test modules grouped by topic. Within a test module, we use the VBAProject_Tests.Procedures class instance, procs to manage test groups and reporting.

## Core Data Management Pattern

**Use structured data objects instead of ad hoc Excel ranges, arrays etc.:**
Projects utilize a structured approach to data management through use of ExcelSteps addin classes for data objects. The project organizes data to be managed as tblRowsCols (rows x columns tables) and mdlScenario (alternate columns x rows format) objects.
```vb
' Initialize global data objects
Dim tbls As Object, mdls As Object
If Not InitAllTbls(tbls) Then GoTo ErrorExit  ' Tables collection
If Not InitAllMdls(mdls) Then GoTo ErrorExit  ' Models collection
```
- **`tbls`** - Collection of `tblRowsCols` objects (row×column tables)
- **`mdls`** - Collection of `mdlScenario` objects (column×row scenario models)
- Use `tbls.Raw.rngHeader` instead of hardcoded ranges like `A1:Z1`
- Use `mdls.params.ScenModelLoc(mdls.params, "variable_name")` for model lookups

By default, the above example will initialize all tables and models defined in the project. By use case, you can also initialize specific tables/models by passing Boolean parameters to `InitAllTbls` and `InitAllMdls`. This example initializes just the project's params and Weekly Scenario models

```vb
InitMdls(mdls, IsAll:=False, IsParams:=True, IsWeekly:=True)
```

## Table and Model Types

**tblRowsCols Types:**
- **Default Tables**: Header in row 1, data starts row 2
- **Custom Tables**: Flexible positioning via definition string

**mdlScenario Types:**
- **Calculator Scenario Model** (`.IsCalc=True`): Single scenario column model
- **Lite Scenario Model** (`.IsLiteModel=True`): Minimal template columns; formatting instructions on ExcelSteps recipe sheet

Project workbooks typically contain a params Scenario model. It is a Calculator (single-column) on the sheet, params, and is often hidden from the user. It is used for storing single-valued parameters such as filenames, configuration inputs and directory paths

## Critical Function Architecture

**Every function must follow this architecture and error-handling pattern:**

```vb
'--------------------------------------------------------------------------------------
' Short description of what function does (never repeat function name)
' JDL MM/DD/YY
'
Public Function MyFunction(arg1, arg2) As Boolean
    SetErrs MyFunction: If errs.IsHandle Then On Error GoTo ErrorExit
    Dim var1 As String, var2 As Object  ' All Dims immediately after SetErrs
    
    ' Function logic here using "If Not" pattern
    If Not SomeSubFunction() Then GoTo ErrorExit

    ' Example error check - use integer error codes 1, 2, etc. within each function
    If errs.IsFail(var2 Is Nothing, 1) Then GoTo ErrorExit
    Exit Function
    
ErrorExit:
    errs.RecordErr "MyFunction", MyFunction
End Function
```

**Key requirements:**
- `SetErrs` initializes function to True and handles error setup including setting `errs.IsHandle`
- All `Dim` statements immediately follow `SetErrs` line in project code and should follow Dim tst line in tests. No `Dim` statements later in function
- Use `If Not FunctionCall() Then GoTo ErrorExit` pattern for chaining and redirection if errors
- Blank line after Exit Function (but no blank immediately preceding)
- `errs.RecordErr` sets function to False and logs error
- No need to manually set function True/False

**Docstring requirements:**
- 3-line format: hyphens line, description, author/date (e.g., "JDL MM/DD/YY")
- Description never repeats function name
- Hyphens line indicates maximum code width

**VBA Quirks reminders:**
When setting a class attribute by calling a function that sets the attribute with a ByRef argument, you must use a local variable as the argument and then set the attribute equal to the local variable. You cannot pass the attribute directly as the argument

```vb
' Incorrect - does not work to set obj.attr directly
If Not SomeFunctionSetsAttr(obj.attr) Then GoTo ErrorExit
' Correct - use local variable
Dim tempAttr As Object
If Not SomeFunctionSetsAttr(tempAttr) Then GoTo ErrorExit
Set obj.attr = tempAttr
```

## Key Driver Patterns
Driver subs (e.g. user-initiated) in VBAProject_ProjectName follow this structure. 
* SetErrs "driver" argument informs SetErrs that this is a driver subroutine not a function.
* SetApplEnvir sets application environment (screen updating, events, calculation mode) for performance
* errs.RecordErr single argument is subroutine name as String

```vb
'--------------------------------------------------------------------------------------
' Short description of what sub/use case does (never repeat sub name)
' JDL MM/DD/YY
'
'Sub ImportSalesDataDriver()
    SetErrs "driver": If errs.IsHandle Then On Error GoTo ErrorExit
    Dim wHist As Object, mdls As Object
    
    SetApplEnvir False, False, xlCalculationManual
    
    ' Initialize only needed objects to avoid overhead
    If Not InitAllMdls(mdls, IsWeekly:=False, IsWklySalesHist:=True) Then GoTo ErrorExit
    
    ' Main procedure calls
    Set wHist = New_WklyHist
    If Not wHist.ImportSalesDataProcedure(wHist, mdls) Then GoTo ErrorExit

    SetApplEnvir True, True, xlCalculationAutomatic
    Exit Sub
    
ErrorExit:
    errs.RecordErr "ImportSalesDataDriver"
    SetApplEnvir True, True, xlCalculationAutomatic
End Sub
```
**Driver requirements:**
- `SetErrs "driver"` initializes subroutine as driver
- `SetApplEnvir` called at start and before Exit Sub to reset application environment
- `errs.RecordErr` single argument logs error with sub name; causes reporting of nested errors from functions called within driver

## Data Object Initialization
tblRowsCols and mdlScenario objects have `.Init()` and `.Provision()` methods. Init sets basic wayfinding attributes. Provision is more extensive and sets all relevant intra-object locations as ranges. Init is also called by Provision but can be called independently. It locates the object by setting its `.wkbk` (Workbook object), `.sht` (String sheet name) and .wksht (Worksheet object). tblRowsCols and mdlScenario have differing "Refresh" procedures that refresh/propagate specified formulas for calculated variables. mdlScenario `.Refresh` names variable (row-oriented) ranges and scenario column ranges. `tblRowsCols.Provision` optionally names column ranges for variables based on variable names in `.rngHeader`

Default objects can be initialized by just `wkbk` and `sht` argument such as
```vb
If Not tbls.Raw.Init(tbls.Raw, wkbk, "raw_data") Then GoTo ErrorExit
```

Custom objects require either a definition string or explicit arguments to locate the object in the workbook and set its parameters. All definition string sub-parts have standalone argument counterparts, and arguments override specification by the defn string.

mdlScenario Example with definition string (Refresh names ranges and propagates formulas). Docstrings in the classes give details about defn string format.
```vb
With mdlWHist
    defnWklyHist As String = "Weekly:103,2:0:F:T:T:T:T:WklyHist"
    If Not .Provision(mdlWHist, wkbk, defn:=defnWklyHist) Then GoTo ErrorExit
    If Not .Refresh(mdlWHist) Then GoTo ErrorExit    
End With
```

tblRowsCols Example with custom arguments (False argument specifies whether to also reformat the table)
```vb
With tbls.ExampleTbl
   If Not .Provision(tbls.ExampleTbl, ThisWorkbook, False, sht:="home_sheet", _
      IsSetColNames:=False) Then GoTo ErrorExit
End With
```

## ExcelSteps Integration Patterns
**Use ExcelSteps utilities instead of native VBA:**

```vb
' Preferred find: works with hidden cells
Set rng = ExcelSteps.FindInRange(searchRange, "value")

'Avoid: Native Range Find that doesn't work with hidden cells
Set rng = searchRange.Find("value")
```

**Data wayfinding in tblRowsCols:**
```vb
' Key utility functions for tblRowsCols
Function TableLoc(rngCell, rngCol, Optional ishift = 0) As Variant
Sub SetTableLoc(rngCell, rngCol, val, Optional ishift = 0)
'rngCell is a cell within the table's data rows. The intersect of its entire row and rngCol 
'locates the cell returned or whose value is set

'Searches tbl.rngHeader for column header sVal
Function rngTblHeaderVal(tbl, sVal) As Range 

' Key utility functions for mdlScenario
Function ScenModelLoc(mdl, sVar, Optional rngCol) As Range
Sub SetScenModelLoc(mdl, sVar, val, Optional rngCol)
'sVar is the name of a variable in the Scenario Model's .colrngVarNames column
'rngCol is an optional scenario column range in the model (not specified for .IsCalc models where .colrngModel is single column)
```

General (utility functions and subs in ExcelSteps; call by ExcelSteps.utility_name())
```vb
'Open file at fullpath (sets wkbkOpened if successful)
Public Function OpenFile(ByVal fullpath As String, wkbkOpened As Workbook) As Boolean

'Close and/or SaveAs wkbk to filepath (overwrites if exists)
Public Function SaveAsCloseOverwrite(ByVal wkbk As Workbook, _
            ByVal filepath As String, Optional IsSave As Boolean = True, _
            Optional IsClose As Boolean = True) As Boolean

'Return rng of contiguous, populated cells from cell range, rng1; 
'Searches rng1 column (xlDown) if IsRows=True; rng1 row otherwise (xlToRight)
'Returns multicell Range from rng1 to populated extent
Function rngToExtent(rng1, IsRows) As Range
```

- **Data Object Iteration**
- Use `.rowCur` and `colCur` as temporary variables to track iteration within `tblRowsCols` and `mdlScenario` instances.
- Use `.colrngModel` and `rngRows` attributes as overall Scenario Model column and row iteration range but column iteration begins at `.colrngFirstScenario` and needs to check inclusion in `.colrngPopCols` to skip blanks/unused columns. Similarly check inclusion in `.rngPopRows` for Scenario Model row iteration
- tblRowsCols.rngHeader and .rngRows contains contiguous column and row blocks, so iteration can be continuous

```vb
' Use .colCur for Scenario columns
Set mdl.colCur = mdl.rngPopCols.Columns(i)
cellValue = mdl.ScenModelLoc(mdl, "variable_name", mdl.colCur)

' Key utility functions - note .colrngModel may be multirange
Function ScenModelLoc(mdl, sVar, Optional rngCol) As Range
Sub SetScenModelLoc(mdl, sVar, val, Optional rngCol)

' Handle multirange columns when iterating
For Each colArea In mdl.colrngModel.Areas
    For Each col In colArea.Columns
        ' Process each column
    Next col
Next colArea
```

**Preferred Array/Range writing patterns:**
Preferred for writing multiple header values
```vb
rng.Value = Split("Header1,Header2,Header3", ",")
```

Preferred for transferring range of values - set source and equal-sized destination ranges. Also illustrages use of .ScenModelLoc and .colrngFirstScenario for wayfinding
```vb
    With ProjCls
        ' Set source and destination ranges
        Set rngSrc = Intersect(.wkshtPivot.Rows(1), .colRngSrc)
        Set rngDest = Intersect(.mdl.ScenModelLoc(.mdl, "date_wkstart").EntireRow, .mdl.colrngFirstScenario)
        Set rngDest = .mdl.wksht.Range(rngDest, rngDest.Offset(0, .colRngSrc.Count - 1))

        ' Transfer the values from source to destination
        rngDest.Value = rngSrc.Value
    End With
```

** General Code Syntax and Style Guidelines**
Use single-line `If` statements for simple conditions (that do not exceed one line).
```vb
If condition Then action
```

In project code, strictly use continuation `_` if line length exceeds the length of the docstring hyphens line (typically 95-100 characters). In test code, use continuation `_` more liberally for readability. Its ok to exceed hyphen line length by 10-20 characters if it improves readability

## Testing Framework
**Unit tests use custom Test and Procedures classes:**
We declare tbls (or mdls) as Object and then call the projects .New_Tbls or .New_Mdls function and tbls or mdls .Init() to instance tbls and the individual data objects whose hard-coded instantiation is in .Init. Note that, if multiple tests will utilize the same pattern of setting and initializing tbls, mdls or project classes, the initialization code should be placed in a helper subroutine that has `tst` and other objects as arguments

```vb
'Example test; do not include explanatory comments in actual tests - just action description like "Check sht"
'-------------------------------------------------------------------------------------
' Verbatim copy of docstring from function being tested
' JDL MM/DD/YY
'
Sub test_FunctionName(procs)
    Dim tst As New Test: tst.Init tst, "test_FunctionName"
    Dim tbls As Object, expected as String
    Set tbls = VBAProject_ProjectName.New_Tbls
    tst.Assert tst, tbls.Init(tbls)

    With tst
        'Check sht set as way of checking initialization
        .Assert tst, VBAProject_ProjectName.InitAllTbls(tbls)  ' Use .Assert for all function calls
        expected = "raw_data"
        .Assert tst, tbls.raw.sht = expected
        .Update tst, procs  ' Always last line in With block
    End With
End Sub
```
**Pattern for opening a test data file**
Place in a helper sub if repeated in multiple tests
```vb
sep = Application.PathSeparator
pathFile = .wkbkTest.Path & sep & "test_data_import" & sep & "ML - WeeklySales - Mockup.xlsx"
.Assert tst, Len(Dir$(pathFile)) > 0
.Assert tst, ExcelSteps.OpenFile(pathFile, wkbkData)
.Assert tst, Not wkbkData Is Nothing
```

**Test driver organization:**
Driver sub is at top of module and used to run tests in Procedure (e.g. procs) groups. 
- Within the driver, new test calls (and new procs groups of tests) should be inserted at the end of the driver
- All newly-written tests should be inserted immediately below the driver sub for navigation ease
- procs Init second argument is test suite sheet name for writing test results; 3rd argument is test suite name used for MsgBox reporting of results to user.
- Comments included here are documentation only. Do not include them in generated code
```vb
Sub TestingDriver_Dashboard()
    Dim procs As New Procedures, AllEnabled As Boolean

    ' Select which procs test groups to run
    With procs
        .Init procs, ThisWorkbook, "Project", "Tests_ProjectName"
        .TblsAndMdls.Enabled = True  ' Toggle procedure groups
        AllEnabled = False
    End With
    
    ' Test individual procedure groups
    With procs.TblsAndMdls
        If .Enabled Or AllEnabled Then
            procs.curProcedure = .name
            test_InitAllTbls procs
            test_InitAllMdls procs
        End If
    End With

    ' Report results to user
    procs.EvalOverall procs
End Sub
```

**Test/Production mode toggle:**
Set .IsTest global constant = False for tests that import/export files or other use cases involving user interaction when running in production mode
```vb
VBAProject_ProjectName.IsTest = True  ' Set test mode
Dim pathname As String
pathname = tst.wkbktest.Path & mdls.params.ScenModelLoc(mdls.params, "path_testing").Value
```

## Cross-Platform Dictionary Usage
**Use project's custom dictionary class instead of VBA Dictionary:**
```vb
Dim dict As Object
Set dict = New dictionary_cls
dict.Add "key", "value"  ' Cross-platform compatible
```

## File Organization (using ProjectName "Dashboard" as example)
- **Interface modules** (`VBAProject_projectname_Interface.vb`) - Main driver subs and initialization functions
- **Validation modules** - Data validation and business logic
- **Class modules** - Custom classes like `WklyHist`, `CurSnap`
- **Test modules** (`VBAProject_Tests_*`) - Organized by functional area
- **ExcelSteps classes** - Structured data management (`tblRowsCols`, `mdlScenario`)

## Naming Conventions
- **Functions**: PascalCase returning Boolean (e.g., `InitAllTbls`, `RefreshCalendar`)
- **Variables**: camelCase or lowercase with descriptive prefixes (`wHist`, `mdls`, `tbls`)
- **Constants**: PascalCase with descriptive names (e.g., `defnWeekly`, `shtThisMonth`)
- **Range objects**: Prefix with `rng` or `cell` (e.g., `rngHeader`, `rngRows`, `cellSrc` etc.)

## Common Integration Points
- **File I/O**: Always use `ExcelSteps.OpenFile` and `SaveAsCloseOverwrite`
- **Pivot tables**: Use class attributes pattern (`wHist.pivotTable`, `wHist.pivotCache`)
- **Parameter blocks**: Adjacent to scenario models with `InitParamBlock` pattern

## ExcelSteps and Test Suite Code Modules
- **tblRowsCols_cls.vb**: ExcelSteps Class module for tblRowsCols object
- **mdlScenario_cls.vb**: ExcelSteps Class module for mdlScenario object
- **dictionary_cls.vb**: ExcelSteps Class module for cross-platform dictionary class
- **procedures_cls.vb**: Test Suite Class module for Procedures object