---
name: vba-function-architecture
description: Enforce compact VBA function architecture for code: single-purpose scope, compact Dim declarations,~50-line size target, attribute-first coding, and correct error-handling pattern by return type. Use when writing/refactoring VBA functions or reviewing long/complex procedures.
---

# VBA Function Architecture

## Pocket rules

1. Single purpose only.
2. All Dim statements at top of function with multiple declarations per line preferred.
3. Target ~50 lines max (excluding docstring lines).
4. If too complex, split helpers or make a thin `...Procedure` orchestrator.
5. Prefer object attributes directly (`tbl.wksht`, `mdl.rngRows`, `.nRows`); avoid gratuitous local aliases.
6. Use `If Not Helper() Then GoTo ErrorExit` for dependent call chains.
7. Match error structure to return type:
   - `Boolean`: `SetErrs FnName` / `errs.RecordErr "Locn", FnName`
   - non-Boolean: `SetErrs "non-bool"` / `errs.RecordErr "Locn"`
8. Use files instance of ProjFiles and colinfo instance of ColInfo to manage file paths and names and to manage metadata about variables. Pass ByRef to helpers as needed.

Locals are justified only for ByRef pass-through constraints, short-scope loop/index temps, or meaningful readability gains.

## Related skills

- `vba-boolean-function-error-handling-structure`
- `vba-non-boolean-function-error-handling-structure`
- `vba-driver-sub-error-handling-structure`
- `vba-excelsteps-projfiles-class-as-files`
- `vba-excelsteps-colinfo-class`

Use these for exact signatures and edge cases.

## Review checklist

- [ ] Function is single-purpose.
- [ ] All `Dim` statements at the top of function, with multiple declarations per line preferred.
- [ ] Function body is near 50 lines, or intentionally an orchestrator.
- [ ] No gratuitous local aliases of object attributes.
- [ ] Error-handling pattern matches return type.
- [ ] `If Not ... Then GoTo ErrorExit` chain is used for dependent calls.
- [ ] `errs.IsFail` codes are local and deterministic.


Updated 5/12/26