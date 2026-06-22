---
name: vba-excelsteps-colinfo-class
description: Use ExcelSteps ColInfo class to read column metadata from a colinfo_ sheet. Use when initializing colinfo for import pipelines, when extracting index/metric variable names from colinfo metadata, when getting VarNameNorm-to-VarNameRaw mappings, or when replacing deprecated hardcoded constant arrays (sBRMetricRawNames, keyFieldsBR, etc.) with Yield function calls.
---

# ExcelSteps ColInfo Class

## What it is

ColInfo manages a column-metadata table (`colinfo_` sheet). It provides structured
access to per-variable metadata: normalized names, raw names, index/metric flags,
fill rules, and filter rules. One ColInfo instance can serve multiple import tables
by switching `curTbl` via `SetCurTbl`.

## Quick start

```vb
Dim colinfo As Object
Set colinfo = ExcelSteps.New_ColInfo
If Not colinfo.Init(colinfo, files, curTbl:="BRRaw_prodn") Then GoTo ErrorExit
```

`files` must already be initialized (see `vba-excelsteps-projfiles-class-as-files`).

## Init signature

```vb
colinfo.Init(colinfo, files As Object, [curTbl As String], [wkbkColInfo As Workbook]) As Boolean
```

- Default: opens `files.pfColInfo` (ColInfo.xlsx on disk)
- `wkbkColInfo` — pass `ThisWorkbook` to use an embedded `colinfo_` sheet instead

**Embedded sheet pattern (Dashboard project):**
```vb
If Not colinfo.Init(colinfo, files, curTbl:="BRRaw_prodn", wkbkColInfo:=ThisWorkbook) _
    Then GoTo ErrorExit
```

After `Init`, `colinfo.tbl` is provisioned. `curTbl` and `rngRowsCurTbl` are set
only if the optional `curTbl` arg was passed.

## SetCurTbl

Switches the active table, sorts by its column, and sets `rngRowsCurTbl`:

```vb
If Not colinfo.SetCurTbl(colinfo, "BRRaw_tests") Then GoTo ErrorExit
```

Re-entrant — safe to call multiple times to switch between tables.

## Key attributes

| Attribute | Type | Description |
|---|---|---|
| `tbl` | tblRowsCols | Provisioned table for the colinfo_ sheet |
| `curTbl` | String | Active table column name |
| `rngRowsCurTbl` | Range | Non-blank rows for curTbl (set by SetCurTbl) |

## Yield functions

All Yield functions require `SetCurTbl` to have been called. They read from
`rngRowsCurTbl` and return data for the active table only.

### YieldAryIndices — index variable names

```vb
Dim aryIdx As Variant
aryIdx = colinfo.YieldAryIndices(colinfo)
' Returns: Array of VarNameNorm where IsIndex = True, in colinfo sort order
' e.g., Array("Location", "ProdType", "Year", "SerialWeek")
```

### YieldAryMetrics — metric variable names

```vb
Dim aryMet As Variant
aryMet = colinfo.YieldAryMetrics(colinfo)
' Returns: Array of VarNameNorm where IsIndex <> True, in colinfo sort order
' e.g., Array("Net_Sales", "Discounts", "Markdowns", "COGS")
```

### YieldDNormalize — VarNameNorm → VarNameRaw mapping

```vb
Dim dNorm As Object
Set dNorm = colinfo.YieldDNormalize(colinfo)
' Returns: Dictionary of VarNameNorm -> VarNameRaw for all rows in rngRowsCurTbl
' e.g., dNorm.Item("Location") = "Locn_Raw"
```

## colinfo_ sheet structure

Required columns: `VarNameNorm`, `VarNameRaw`, `IsIndex`, `Description`,
`units`, `data_type_VBA`, `FillVals`, `FilterVals`, then one column per
import table (e.g., `BRRaw_prodn`, `BRRaw_tests`).

Table columns contain the sort-order integer for each variable in that table.
Blank = variable not included in that table.

`FillVals` and `FilterVals` use JSON-like dict strings convertible to Dictionary instance by `ParseStringToDictProcedure`: `{"BLANK":"Unknown"}`,
`{"KeepOnly":"Online"}`.

## Full init pattern (production)

```vb
Set files = ExcelSteps.New_ProjFiles
If Not files.Init(files, ThisWorkbook) Then GoTo ErrorExit

Set colinfo = ExcelSteps.New_ColInfo
If Not colinfo.Init(colinfo, files, curTbl:="BRRaw_prodn", _
    wkbkColInfo:=ThisWorkbook) Then GoTo ErrorExit
```

## Full init pattern (tests)

```vb
ExcelSteps.IsTest = True
Set files = ExcelSteps.New_ProjFiles
tst.Assert tst, files.Init(files, ThisWorkbook, "BR_Import")

Set colinfo = ExcelSteps.New_ColInfo
tst.Assert tst, colinfo.Init(colinfo, files, curTbl:="BRRaw_tests", _
    wkbkColInfo:=ThisWorkbook)
```

## Using colinfo_.tbl colrng attributes directly

After `colinfo.Init`, `colinfo.tbl` is a provisioned `tblRowsCols`. Because `Provision`
calls `SetColRanges`, the standard colinfo column ranges are already set as named
attributes — no need to look them up again with `rngTblHeaderVal`:

```vb
' Preferred — use pre-set colrng attribute
Intersect(colinfo.rngRowsCurTbl, colinfo.tbl.colrngFilterVals).ClearContents
Intersect(rngLocnRow.EntireRow, colinfo.tbl.colrngFilterVals).Value2 = sFilter

' Avoid — redundant lookup
Set rngFilterValsCol = colinfo.tbl.rngTblHeaderVal(colinfo.tbl, "FilterVals").EntireColumn
Intersect(colinfo.rngRowsCurTbl, rngFilterValsCol).ClearContents
```

Available colrng attributes (set in `tblRowsCols.SetColRanges` for `shtColInfo`):
`colrngVarNorm`, `colrngVarRaw`, `colrngIsIndex`, `colrngVarDesc`, `colrngVarUnits`,
`colrngFillVals`, `colrngFilterVals`.

JDL 5/12/26