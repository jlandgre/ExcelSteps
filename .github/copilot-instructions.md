# Dashboard VBA Project - AI Coding Instructions
updated 4/14/26
## Project Architecture

Three-workbook structure for cross-platform (Windows/Mac) compatibility:
- **Dashboard.xlsm** — Main workbook (VBA Project: `VBAProject_Dashboard`)
- **XLSteps.xlam** — ExcelSteps add-in (VBA Project: `ExcelSteps`)
- **tests_Dashboard.xlsm** — Unit test suite (VBA Project: `Tests`)

`VBAProject_Dashboard` references `ExcelSteps`; `Tests` references `VBAProject_Dashboard`.

Code files in `src/` and `tests/` use extensions `.bas` (standard modules), `.cls` (class modules), `.frm` (userforms). Before making code changes, confirm a VS Code Terminal is running with `xlwings vba edit` active for the target workbook. Stop and address this if not enabled.

## Core Data Management Pattern

All data is managed through ExcelSteps structured objects — never ad hoc ranges or arrays:
- **`tbls`** — instance of `Tables.cls`; collection of `tblRowsCols` objects (row×column tables)
- **`mdls`** — instance of `Models.cls`; collection of `mdlScenario` objects (column×row scenario models)

```vb
Dim tbls As Object, mdls As Object
If Not InitAllTbls(tbls) Then GoTo ErrorExit
If Not InitAllMdls(mdls) Then GoTo ErrorExit
```

Use named object attributes instead of hardcoded ranges: `tbls.Raw.rngHeader`, `mdls.params.ScenModelLoc(mdls.params, "variable_name")`.

`InitAllTbls`/`InitAllMdls` initialize (instance if `Is Nothing`), provision (set `.wkbk`, `.sht`, `.wksht`, ranges), and optionally refresh (update named ranges and formulas). Use named parameters to initialize a subset:

```vb
InitAllMdls(mdls, IsAll:=False, IsParams:=True, IsWeekly:=True, IsRefresh:=True)
```

## Table and Model Types

**tblRowsCols Types:**
- **Default Tables**: Header in row 1, data starts row 2
- **Custom Tables**: Flexible positioning via definition string

**mdlScenario Types:**
- **Calculator Scenario Model** (`.IsCalc=True`): Single scenario column model
- **Lite Scenario Model** (`.IsLiteModel=True`): Minimal template columns; formatting instructions on ExcelSteps recipe sheet

Project workbooks typically contain a `params` Scenario model on the `params` sheet. It is a Calculator (single-column) model and is generally hidden from the user. It is used for storing single-valued parameters such as configuration inputs and directory paths

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

**ByRef alias pattern (VBA quirk):** A class attribute cannot be passed directly as ByRef — VBA silently ignores the assignment. Set the attribute first; create an alias only for the ByRef call. Do not create a proxy local, do work on it, then assign back to the attribute.

```vb
' Preferred
Set obj.attr = ExcelSteps.New_tbl
Dim tblAlias As Object: Set tblAlias = obj.attr
If Not tblAlias.Provision(tblAlias, wkbk, False) Then GoTo ErrorExit
```

## Key Driver Patterns

```vb
'--------------------------------------------------------------------------------------
' Short description of what sub/use case does (never repeat sub name)
' JDL MM/DD/YY
'
Sub ImportSalesDataDriver()
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

Default objects: `wkbk` and `sht` are sufficient:
```vb
If Not tbls.Raw.Init(tbls.Raw, wkbk, "raw_data") Then GoTo ErrorExit
```

Custom objects use a definition string (arguments override defn string). Docstrings in classes give format details:
```vb
With mdlWHist
    If Not .Provision(mdlWHist, wkbk, defn:="Weekly:103,2:0:F:T:T:T:T:WklyHist") Then GoTo ErrorExit
    If Not .Refresh(mdlWHist) Then GoTo ErrorExit
End With
```

```vb
With tbls.ExampleTbl
    If Not .Provision(tbls.ExampleTbl, ThisWorkbook, False, sht:="home_sheet", _
       IsSetColNames:=False) Then GoTo ErrorExit
End With
```

## ExcelSteps Integration Patterns

**Prefer ExcelSteps utilities over native VBA:**
```vb
Set rng = ExcelSteps.FindInRange(searchRange, "value")  ' works with hidden cells
' Not: searchRange.Find("value")
```

**Wayfinding — tblRowsCols (call via `tbl.FunctionName`):**
```vb
Function TableLoc(rngCell, rngCol, Optional ishift = 0) As Variant  ' get value at row/col intersection
Sub SetTableLoc(rngCell, rngCol, val, Optional ishift = 0)           ' set value at row/col intersection
Function rngTblHeaderVal(tbl, sVal) As Range                         ' find column range by header name
```

**Wayfinding — mdlScenario (call via `mdl.FunctionName`):**
```vb
Function ScenModelLoc(mdl, sVar, Optional rngCol) As Range  ' get variable range
Sub SetScenModelLoc(mdl, sVar, val, Optional rngCol)         ' set variable value
```

**General utilities (call as `ExcelSteps.FunctionName`):**
```vb
Public Function OpenFile(ByVal fullpath As String, wkbkOpened As Workbook) As Boolean
Public Function SaveAsCloseOverwrite(ByVal wkbk As Workbook, ByVal filepath As String, _
    Optional IsSave As Boolean = True, Optional IsClose As Boolean = True) As Boolean
Function rngToExtent(rng1, IsRows) As Range  ' contiguous populated range from rng1
```

**Iteration:** Use `.rowCur`/`.colCur` as temporary iteration variables. For `mdlScenario`, iterate columns via `.colrngPopCols` (skips blanks) starting at `.colrngFirstScenario`; iterate rows via `.rngPopRows`. Handle multirange `.colrngModel`:
```vb
For Each colArea In mdl.colrngModel.Areas
    For Each col In colArea.Columns
    Next col
Next colArea
```
`tblRowsCols.rngHeader` and `.rngRows` are contiguous, so iteration can be continuous.

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
Unit tests use custom Test and Procedures classes. Test class manages individual test cases and assertions. Procedures class manages groups of tests and reporting. Tests are organized by functional area in separate test modules. Each test module has a driver subroutine that calls the individual tests in the module and reports results to the user. A driver subroutine may contain multiple test groups by procedure (e.g. procs.TblsAndMdls) that can be toggled on/off for selective testing.

**Cross-workbook class instantiation in Test Workbook:**
Because test code resides in a separate workbook from the classes being tested, you cannot directly instantiate classes with `New` or `Dim As New`. Instead, use factory functions in the source workbook's Validation or Interface module:

```vb
' INCORRECT - Cannot access classes across workbooks
Dim dict As New Dictionary
Set dict = New Dictionary

' CORRECT - Use factory function from source workbook
Dim dict As Object
Set dict = ExcelSteps.New_Dictionary  ' For ExcelSteps classes
Set tbls = VBAProject_ProjectName.New_Tbls  ' For project classes
```

**Data object initialization pattern:**
Declare tbls (or mdls) as Object, call the project's `New_Tbls` or `New_Mdls` factory function, then call `.Init()` to instance the collection and individual data objects whose hard-coded instantiation is in `.Init`. If multiple tests utilize the same initialization pattern for tbls, mdls, or project classes, place the initialization code in a helper subroutine that has `tst` and other objects as arguments

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
**Test driver structure:**
- Driver sub at top of module; new tests inserted immediately below it
- New test calls added at end of their procs group in the driver
- `procs.Init` args: 2nd = sheet name for results, 3rd = name for MsgBox reporting
- Do not include explanatory comments shown here in generated code
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

## Cross-Platform Dictionary

Always use the custom dictionary class — not `Scripting.Dictionary` (Windows-only):
```vb
Dim dict As Object
Set dict = New dictionary_cls
dict.Add "key", "value"
```

## File Organization (using ProjectName "Dashboard" as example)
- **Interface modules** (`VBAProject_projectname_Interface.vb`) - Main driver subs and initialization functions
- **Validation modules** - Data validation and business logic
- **Class modules** - Custom classes like `WklyHist`, `CurSnap`
- **Test modules** (`Tests_*`) - Organized by functional area
- **ExcelSteps classes** - Structured data management (`tblRowsCols`, `mdlScenario`)

## Naming Conventions
- **Functions**: PascalCase returning Boolean (e.g., `InitAllTbls`, `RefreshCalendar`)
- **Variables**: camelCase or lowercase with descriptive prefixes (`wHist`, `mdls`, `tbls`)
- **Constants**: PascalCase with descriptive names (e.g., `defnWeekly`, `shtThisMonth`)
- **Range objects**: Prefix with `rng` or `cell` (e.g., `rngHeader`, `rngRows`, `cellSrc` etc.)

## Common Integration Points
- **File I/O**: Always use `ExcelSteps.OpenFile` and `ExcelSteps.SaveAsCloseOverwrite`
- **Pivot tables**: Use class attributes pattern (`wHist.pivotTable`, `wHist.pivotCache`)
- **VBA classes**: Use direct public attributes — do not use `Property Get/Let`