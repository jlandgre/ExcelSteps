---
name: vba-non-boolean-function-error-handling-structure
description: Correct SetErrs and RecordErr call patterns for non-Boolean VBA functions (Variant, Object, String, Long, etc.) in the ExcelSteps error handling framework. Use when writing or reviewing functions that return a non-Boolean type, when seeing "SetErrs FunctionName" in a non-Boolean function, or when RecordErr is called with two arguments in a non-Boolean context.
---

# Non-Boolean Function Error Handling Structure

## The Core Distinction

`SetErrs` and `errs.RecordErr` have **two distinct call signatures** depending on return type:

| Return type | `SetErrs` call | `RecordErr` call |
|---|---|---|
| `Boolean` | `SetErrs MyFunction` | `errs.RecordErr "Locn", MyFunction` |
| Everything else | `SetErrs "non-bool"` | `errs.RecordErr "Locn"` |

**"Everything else"** = `Variant`, `Object`, `String`, `Long`, `Integer`, `Double`, any class type.

## Correct Pattern

```vb
Public Function YieldSomething(obj) As Variant
    SetErrs "non-bool": If errs.IsHandle Then On Error GoTo ErrorExit
    Dim ary() As Variant

    ' ... function logic ...

    YieldSomething = ary
    Exit Function

ErrorExit:
    errs.RecordErr "ColInfo.YieldSomething"
End Function
```

## Why

- `SetErrs BooleanFunction` initializes the function variable to `True` — only valid for `Boolean` return types.
- `errs.RecordErr "Locn", BooleanFunction` sets the function to `False` to signal failure — meaningless for non-Boolean returns.
- For non-Boolean functions, `SetErrs "non-bool"` skips the function-init step, and single-arg `errs.RecordErr "Locn"` records the error location without touching the return value.

## Common Mistake to Catch

```vb
' WRONG - function returns Variant, not Boolean
Public Function YieldAryIndices(colinfo) As Variant
    SetErrs YieldAryIndices: ...          ' ← should be "non-bool"
    ...
ErrorExit:
    errs.RecordErr "...", YieldAryIndices  ' ← remove second argument
End Function
```
