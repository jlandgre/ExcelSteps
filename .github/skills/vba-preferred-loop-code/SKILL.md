---
name: vba-preferred-looping
description: Prefer For Each loops for arrays, dictionary keys, and comma-separated lists split into arrays. Use a separate Long variable named idx inside the loop when position tracking is needed.
---

# VBA Preferred Looping

Use these looping defaults unless there is a clear reason to do otherwise.

## Primary preference

Prefer `For Each` loops when iterating over:

- arrays
- `Collection` objects
- `Dictionary` keys or items
- comma-separated lists that can be split into an array

`For Each` is usually clearer, less error-prone, and easier to maintain than index-based looping when the main goal is to process each item.

## Position tracking and Explanatory comments

- When the loop also needs to know the position of the current item, still prefer `For Each` and maintain a separate `Long` variable named `idx` inside the loop scope. See exceptions below when index-based loops are appropriate.
- Always include a blank line and comment above each loop
- Strictly avoid using Goto within loops except for Goto ErrorExit for error handling.

Use:
```vb
    Dim idx As Long, item As Variant
    
    ' Debug print items contents
    idx = 0
    For Each item In items
        idx = idx + 1
        Debug.Print idx, item
    Next item
```

Do not switch to a `For idx = LBound(...) To UBound(...)` loop

## Loop variable naming

- Use Variant `val` as generic item unless the iterated object has a plural name then use the singular form of that name

Preferred:
```vb
    Dim address As Variant, val as Variant
    For Each address In addresses
        Debug.Print address
    Next address

    For Each val In Split(lst, ",")
        Debug.Print (val)
    Next val
```

## Dictionaries

When iterating a dictionary, prefer `For Each` over keys.

Preferred:
```vb
    Dim key As Variant

    ' Explanatory comment
    For Each key In dict.Keys
        Debug.Print key, dict(key)
    Next key
```

Preferred when position is also needed:
```vb
    Dim idx As Long, key As Variant

    ' Explanatory comment
    idx = 0
    For Each key In dict.Keys
        idx = idx + 1
        Debug.Print idx, key, dict(key)
    Next key
```
## Comma-separated lists

When handling comma-separated input, Split in `For Each` line. Do not split into a local variable in a separate line

Preferred:
```vb
    Dim val As Variant

    ' Explanatory comment
    For Each val In Split(csvText, ",")
        Debug.Print Trim$(val)
    Next val
```

## When index-based loops are appropriate

Use an index-based `For ... Next` loop only when one of these is true:

- you need reverse iteration
- you are iterating multiple arrays in parallel by shared index
- you need direct neighbor access such as `arr(idx - 1)` or `arr(idx + 1)`

In those cases, use `Long` for the loop counter.

Preferred:
```vb
    Dim idx As Long

    ' Explanatory comment
    For idx = LBound(arr) To UBound(arr)
        If idx > LBound(arr) Then
            Debug.Print arr(idx - 1), arr(idx)
        End If
    Next idx
```
version 4/29/26 JDL
