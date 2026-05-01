---
name: excelsteps-dictionary-usage
description: Use ExcelSteps cross-platform Dictionary class instead of Scripting.Dictionary. Use when creating, populating, or iterating dictionaries in VBA
---

# ExcelSteps Dictionary Usage

## Quick start

Always use `ExcelSteps.New_Dictionary` — never `Scripting.Dictionary` (Windows-only).

```vb
Dim dict As Object
Set dict = ExcelSteps.New_Dictionary
dict.Add "key", "value"
```

In test workbooks, use the factory function:
```vb
Set dict = ExcelSteps.New_Dictionary
```

## Key API

| Method / Property | Purpose |
|---|---|
| `dict.Add key, value` | Add or update a key-value pair |
| `dict.Item(key)` | Get value by key |
| `dict.Exists(key)` | Check if key present |
| `dict.GetKeys()` | Returns `String()` array of all keys |
| `dict.GetValues()` | Returns `Variant()` array of all values |
| `dict.Size` | Number of items |
| `dict.Remove key` | Remove a key |
| `dict.Clear` | Remove all items |
| `dict.ParseStringToDictProcedure(str)` | Parse JSON-like string into dict |

## ParseStringToDictProcedure

Populates the dictionary from a `{key:value, ...}` string. Keys do **not** need to be quoted — the parser infers unquoted tokens as keys. Quoted keys are also accepted.

Supported value types: string (quoted), numeric, boolean (`True`/`False`), empty.

```vb
' Keys without quotes (preferred for simple keys)
If Not dict.ParseStringToDictProcedure("{x:1, y:2, z:3}") Then GoTo ErrorExit

' Keys with quotes also valid
If Not dict.ParseStringToDictProcedure("{""label"":""sales"", ""active"":True}") Then GoTo ErrorExit
```

## Iterating keys

`GetKeys()` returns a `String` array — use `For Each` per preferred loop pattern:

```vb
Dim key As Variant

' Debug print each key and value
For Each key In dict.GetKeys()
    Debug.Print key, dict.Item(key)
Next key
```
version 4/29/26 JDL