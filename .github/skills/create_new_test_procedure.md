# Skill: Adding a New Test Procedure

## Overview
As we add tests, we group them under `Procedure` instances (instanced as `procs` in test modules such as `tests_ParseModel`). The `Procedures` class manages these by declaring each as an `Object` and instancing them as `New Procedure` in the `Procedures.Init` method. The Procedure's .Enabled flag is used to turn that group of tests on or off while debugging.

## Architecture
In a test driver subroutine (e.g., `TestDriver_ParseModel`):
1. The driver at the beginning of the module instances `procs` and calls `.Init`
2. The driver uses `With procs.<procedure_name>` blocks to run test groups based on the `.Enabled` flag
3. Each procedure gets its own section with a docstring fence separator

## Steps to Add a New Procedure

When you need to add a new procedure `<new_procedure_name>` to `tests_<module_name>`:

### 1. Update Procedures.cls
Add the declaration and initialization for the new procedure:

**Declarations section:**
```vb
Public <new_procedure_name> As Object
```

**Procedures.Init method:**
Add the initialization under the appropriate driver group comment. To keep the Init method clear and readable, initializations are grouped by driver subroutine:

```vb
' TestDriver_<module_name>
Set .<new_procedure_name> = New Procedure
.<new_procedure_name>.Name = "<new_procedure_name>"
```

**Note:** The property name in `Procedures.cls` must match the programmatic name used in test modules.

### 2. Enable the New Procedure in Driver Initialization
In the `TestDriver_<module_name>` subroutine, add the `.Enabled` flag as the **last item** in the initialization block:

```vb
Sub TestDriver_<module_name>()
    Dim procs As New Procedures, AllEnabled As Boolean
    
    With procs
        .Init procs, ThisWorkbook, "<module_name>", "Tests_<module_name>"
        SetApplEnvir False, False, xlCalculationManual
        
        'Enable testing of all or individual procedures
        AllEnabled = False
        .ExistingProcedure1.Enabled = True
        .ExistingProcedure2.Enabled = True
        .<new_procedure_name>.Enabled = True  ' Add as last item
    End With
```

### 3. Add Test Block to Driver Subroutine
Add the `With procs.<new_procedure_name>` block as the **last block** at the bottom of the driver subroutine, just before `procs.EvalOverall`. This positions it closest to where the actual test functions will be added:

```vb
    ' ... existing procedure blocks ...
    
    'Setup procedure group
    With procs.<new_procedure_name>
        If .Enabled Or AllEnabled Then
            procs.curProcedure = .Name
            test_YourTest1 procs
            test_YourTest2 procs
            ' ... additional tests
        End If
    End With
    
    procs.EvalOverall procs
End Sub
```

### 4. Add Docstring Fence
Add a separator immediately after the driver subroutine and before the first test:

```vb
'--------------------------------------------------------------------------------------
'--------------------------------------------------------------------------------------
' procs.<new_procedure_name>
'--------------------------------------------------------------------------------------
'--------------------------------------------------------------------------------------
```

## Example
See `tests_ParseModel.bas` for a complete reference implementation.

