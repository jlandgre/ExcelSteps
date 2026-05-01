# Add New Test Procedure — Code Templates

## Step 1: Procedures.cls

**Declarations section:**
```vb
Public <new_procedure_name> As Object
```

**Procedures.Init method** (grouped under driver comment):
```vb
' TestDriver_<module_name>
Set .<new_procedure_name> = New Procedure
.<new_procedure_name>.Name = "<new_procedure_name>"
```

## Step 2: Enable in driver init block

```vb
Sub TestDriver_<module_name>()
    Dim procs As New Procedures, AllEnabled As Boolean

    With procs
        .Init procs, ThisWorkbook, "<module_name>", "Tests_<module_name>"
        SetApplEnvir False, False, xlCalculationManual

        AllEnabled = False
        .ExistingProcedure1.Enabled = True
        .ExistingProcedure2.Enabled = True
        .<new_procedure_name>.Enabled = True  ' last item
    End With
```

## Step 3: With block in driver

```vb
    With procs.<new_procedure_name>
        If .Enabled Or AllEnabled Then
            procs.curProcedure = .Name
            test_YourTest1 procs
            test_YourTest2 procs
        End If
    End With

    procs.EvalOverall procs
End Sub
```

## Step 4: Docstring fence in tests module below driver sub

```vb
'--------------------------------------------------------------------------------------
'--------------------------------------------------------------------------------------
' procs.<new_procedure_name>
'--------------------------------------------------------------------------------------
'--------------------------------------------------------------------------------------
```
version 4/29/26 JDL