---
name: vba-function-architecture
description: Enforce compact VBA function architecture for ExcelSteps code: single-purpose scope, ~50-line size target, attribute-first coding, and correct error-handling pattern by return type. Use when writing/refactoring VBA functions or reviewing long/complex procedures.
---

# VBA Function Architecture

## Pocket rules

1. Single purpose only.
2. Target ~50 lines max (excluding docstring/blank lines).
3. If too complex, split helpers or make a thin `...Procedure` orchestrator.
4. Prefer object attributes directly (`tbl.wksht`, `mdl.rngRows`, `.nRows`); avoid gratuitous local aliases.
5. Use `If Not Helper() Then GoTo ErrorExit` for dependent call chains.
6. Match error structure to return type:
   - `Boolean`: `SetErrs FnName` / `errs.RecordErr "Locn", FnName`
   - non-Boolean: `SetErrs "non-bool"` / `errs.RecordErr "Locn"`

Locals are justified only for ByRef pass-through constraints, short-scope loop/index temps, or meaningful readability gains.

## Related skills

- `vba-boolean-function-error-handling-structure`
- `vba-non-boolean-function-error-handling-structure`
- `vba-driver-sub-error-handling-structure`

Use these for exact signatures and edge cases.

## Review checklist

- [ ] Function is single-purpose.
- [ ] Function body is near 50 lines, or intentionally an orchestrator.
- [ ] No gratuitous local aliases of object attributes.
- [ ] Error-handling pattern matches return type.
- [ ] `If Not ... Then GoTo ErrorExit` chain is used for dependent calls.
- [ ] `errs.IsFail` codes are local and deterministic.
