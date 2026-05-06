---
name: vba-driver-sub-error-handling-structure
description: Correct SetErrs and RecordErr call patterns for VBA driver subroutines in the ExcelSteps error handling framework. Use when writing or reviewing Sub procedures that are user-initiated entry points, when seeing SetErrs used in a Sub, or when wiring up application environment setup around a main workflow.
---

# Driver Sub Error Handling Structure

## Correct Pattern

```vb
'--------------------------------------------------------------------------------------
' Short description of what sub/use case does (never repeat sub name)
' JDL MM/DD/YY
'
Sub ImportSalesDataDriver()
    SetErrs "driver": If errs.IsHandle Then On Error GoTo ErrorExit
    Dim wHist As Object, mdls As Object

    SetApplEnvir False, False, xlCalculationManual

    ' Main procedure calls
    If Not InitAllMdls(mdls, IsWeekly:=False, IsWklySalesHist:=True) Then GoTo ErrorExit
    Set wHist = New_WklyHist
    If Not wHist.ImportSalesDataProcedure(wHist, mdls) Then GoTo ErrorExit

    SetApplEnvir True, True, xlCalculationAutomatic
    Exit Sub

ErrorExit:
    errs.RecordErr "ImportSalesDataDriver"
    SetApplEnvir True, True, xlCalculationAutomatic
End Sub
```

## Key Rules

- `SetErrs "driver"` — literal string `"driver"`, never the sub name
- `errs.RecordErr "SubName"` — **single argument**, the sub name as a string literal
- `SetApplEnvir` called at start (disable) and **both** before `Exit Sub` and in `ErrorExit` (restore)
- All `Dim` statements immediately follow `SetErrs` line — no Dims later in the sub

## Distinction from Other Patterns

| Context | `SetErrs` call | `RecordErr` call |
|---|---|---|
| Driver `Sub` | `SetErrs "driver"` | `errs.RecordErr "SubName"` |
| Boolean `Function` | `SetErrs MyFunction` | `errs.RecordErr "Locn", MyFunction` |
| Non-Boolean `Function` | `SetErrs "non-bool"` | `errs.RecordErr "Locn"` |

## Why `errs.RecordErr` Has One Argument Here

`errs.RecordErr "SubName"` with a single arg records the error location and triggers reporting of all nested errors from functions called within the driver. There is no return value to set to False.
