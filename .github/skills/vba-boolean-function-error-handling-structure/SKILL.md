---
name: vba-boolean-function-error-handling-structure
description: Correct SetErrs and RecordErr call patterns for Boolean-returning VBA functions in the ExcelSteps error handling framework. Use when writing or reviewing functions that return Boolean, when wiring up If Not FunctionCall() Then GoTo ErrorExit chains, or when a function must signal success/failure to its caller.
---

# Boolean Function Error Handling Structure

## Correct Pattern

```vb
'--------------------------------------------------------------------------------------
' Short description of what function does (never repeat function name)
' JDL MM/DD/YY
'
Public Function MyFunction(arg1, arg2) As Boolean
    SetErrs MyFunction: If errs.IsHandle Then On Error GoTo ErrorExit
    Dim var1 As String, var2 As Object

    ' Chain sub-calls with If Not pattern
    If Not SomeSubFunction() Then GoTo ErrorExit

    ' Integer error codes 1, 2, etc. within each function
    If errs.IsFail(var2 Is Nothing, 1) Then GoTo ErrorExit

    Exit Function

ErrorExit:
    errs.RecordErr "MyFunction", MyFunction
End Function
```

## Key Rules

- `SetErrs MyFunction` — pass the **function name itself** (not a string), so SetErrs initializes it to `True`
- `errs.RecordErr "Locn", MyFunction` — **two arguments**: location string + function name; sets function to `False` on error
- All `Dim` statements immediately follow the `SetErrs` line — no Dims later
- Blank line after `Exit Function`; no blank line immediately before it
- Use `If Not FunctionCall() Then GoTo ErrorExit` to chain dependent calls

## Distinction from Other Patterns

| Context | `SetErrs` call | `RecordErr` call |
|---|---|---|
| Boolean `Function` | `SetErrs MyFunction` | `errs.RecordErr "Locn", MyFunction` |
| Non-Boolean `Function` | `SetErrs "non-bool"` | `errs.RecordErr "Locn"` |
| Driver `Sub` | `SetErrs "driver"` | `errs.RecordErr "SubName"` |

## Why Two Arguments in `RecordErr`

`errs.RecordErr "Locn", MyFunction` sets `MyFunction = False` to signal failure to the caller. This enables the `If Not MyFunction() Then GoTo ErrorExit` chain to propagate errors up the call stack automatically.
