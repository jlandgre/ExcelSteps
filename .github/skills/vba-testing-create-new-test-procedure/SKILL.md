---
name: vba-testing-create-new-test-procedure
description: Add a new Procedure group to a VBA test module and wire it into the Procedures class and driver sub. Use when adding a test procedure group, enabling a new procs block, or wiring up a new test category in a test driver.
---

# Add New Test Procedure

## Quick start

Add `MyProc` to `tests_Module`:
1. Declare + init in `Procedures.cls`
2. Enable in driver `With procs` block
3. Add `With procs.MyProc` test block before `EvalOverall`
4. Add docstring fence after driver sub

See [REFERENCE.md](REFERENCE.md) for full code templates.

## Workflow

### 1. Update Procedures.cls
- **Declarations:** `Public <new_procedure_name> As Object`
- **Procedures.Init**, under `' TestDriver_<module_name>` comment:
  `Set .<name> = New Procedure` then `.<name>.Name = "<name>"`
- Property name must match programmatic name used in test modules

### 2. Enable in driver init block
Add `.<name>.Enabled = True` as **last item** in the `With procs` init block.

### 3. Add With block to driver
Insert `With procs.<name>` block as **last block** before `procs.EvalOverall`.
Run tests inside `If .Enabled Or AllEnabled Then`.

### 4. Add docstring fence to tests module
Immediately after driver sub, before first test fn — two full hyphens lines,
`' procs.<name>`, two more hyphens lines.

See [REFERENCE.md](REFERENCE.md) for complete code templates per step.

version 4/29/26 JDL