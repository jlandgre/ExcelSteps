---
name: vba-excelsteps-projfiles-class-as-files
description: Use ExcelSteps ProjFiles class to manage project file paths. Use when initializing file paths for a project workbook, when passing `files` to ColInfo.Init or ImportParseNorm.Init, when the code uses New_Dictionary for a `files` variable, or when adding project-specific path/file attributes for raw data imports.
---

# ExcelSteps ProjFiles Class (`files`)

## Quick start

Always instance `files` via `ExcelSteps.New_ProjFiles` — never as a Dictionary.

```vb
Dim files As Object
Set files = ExcelSteps.New_ProjFiles
If Not files.Init(files, ThisWorkbook) Then GoTo ErrorExit
```

`files` is the standard argument name throughout ExcelSteps — use it consistently.

## Init signature

```vb
files.Init(files, wkbk As Workbook, [subdir_tests As String]) As Boolean
```

- `wkbk` — the project workbook (typically `ThisWorkbook`)
- `subdir_tests` — subfolder inside `tests/` for test data (e.g., `"BR_Import"`)

## Generic attributes set by Init

| Attribute | Description |
|---|---|
| `wkbkProj` | Project workbook (Set from `wkbk` arg) |
| `fWkbkProj` | Project workbook filename |
| `IsDevelopment` | True when workbook is in a `src/` folder |
| `pathSrc` | Path to workbook directory (trailing sep) |
| `pathRoot` | Root directory (IsDevelopment only) |
| `pathTests` | `pathRoot & "tests/"` (IsDevelopment only) |
| `pathData` | Production: `pathSrc`; Test: `pathTests[/subdir_tests/]` |
| `pathColInfo` | Same as `pathData` |
| `fColInfo` | `"ColInfo.xlsx"` (fixed) |
| `pfColInfo` | `pathColInfo & fColInfo` |
| `pfImportFile` | `""` — **must be set manually after Init** |

## Setting pfImportFile

`pfImportFile` is not set by `Init`. Set it explicitly before calling ImportParseNorm:

```vb
files.pfImportFile = files.pathData & "BR_Raw_Data.xlsx"
```

## Adding project-specific attributes

ProjFiles has `SetProjSpecificPaths` (private stub). For project-specific paths, add
public attributes directly to the class and populate them in `SetProjSpecificPaths`:

```vb
' In ProjFiles.cls — add public attrs
Public pathBRRawData As String
Public fBRRawData As String
Public pfBRRawData As String

' In SetProjSpecificPaths — populate from pathData
With files
    .fBRRawData = "BR_Raw_Data.xlsx"
    .pfBRRawData = .pathData & .fBRRawData
End With
```

## Critical: initialize files BEFORE passing to other functions

`files` must be fully initialized before being passed to `colinfo.Init` or
`imptbl.Init`. Initialize it in the driver sub (production) or test helper (tests).

**Wrong — initializing files inside a class method:**
```vb
' WRONG: inside InitImportContext
Set files = ExcelSteps.New_Dictionary       ' wrong class
files.Init(files, ThisWorkbook)             ' colinfo can't use this
```

**Correct — initialize in driver, pass as argument:**
```vb
' In driver sub or test helper
Set files = ExcelSteps.New_ProjFiles
If Not files.Init(files, ThisWorkbook) Then GoTo ErrorExit
files.pfImportFile = files.pathData & "BR_Raw_Data.xlsx"
If Not import.InitImportContext(import, files) Then GoTo ErrorExit
```

## Production vs test paths

`pathData` resolves differently based on `ExcelSteps.IsTest`:

| IsTest | pathData |
|---|---|
| False | `pathSrc` (same directory as workbook) |
| True, no subdir_tests | `pathTests` |
| True, subdir_tests set | `pathTests & subdir_tests & sep` |

In tests always set `ExcelSteps.IsTest = True` before calling `files.Init`.

## Test pattern

```vb
ExcelSteps.IsTest = True
Set files = ExcelSteps.New_ProjFiles
tst.Assert tst, files.Init(files, ThisWorkbook, "BR_Import")
files.pfImportFile = files.pathData & "BR_Raw_Mockup.xlsx"
```

JDL 5/12/26