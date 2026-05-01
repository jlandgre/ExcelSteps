---
name: vba-class-arrays
description: Avoid VBA class array property errors by using methods that return arrays instead of public array attributes. Use when defining array attributes on VBA class modules, accessing class arrays from nested functions or external modules, or when encountering "Property let procedure not defined" errors.
---

# VBA Class Arrays

## The Problem

Public array attributes on VBA classes cause errors when accessed from nested functions or external modules:

```vb
' PROBLEMATIC
Public aryColRngs As Variant

' Later in nested function - ERROR:
For i = LBound(obj.aryColRngs) To UBound(obj.aryColRngs)
    Debug.Print obj.aryColRngs(i).Address  ' "Property let procedure not defined"
Next i
```

VBA treats public attributes as properties requiring Get/Let/Set procedures. Direct array element access (`obj.attr(i)`) fails without them.

## Preferred Pattern: Method Returns Array

```vb
' Class Module
Private pKeys() As String
Private pCount As Long

Public Function GetKeys() As String()
    Dim result() As String, i As Long
    ReDim result(0 To pCount - 1)
    For i = 0 To pCount - 1
        result(i) = pKeys(i)
    Next i
    GetKeys = result
End Function
```

```vb
' Caller - assign to local variable first
Dim keys() As String, i As Long
keys = dict.GetKeys          ' Works: method returns array
For i = LBound(keys) To UBound(keys)
    Debug.Print keys(i)      ' Works: accessing local variable
Next i
```

## Alternative: Property Get/Let

If you must expose an array as a property, use explicit procedures — but the local-variable rule still applies:

```vb
' Class Module
Private pAryColRngs As Variant

Public Property Get aryColRngs() As Variant
    aryColRngs = pAryColRngs
End Property
Public Property Let aryColRngs(vArray As Variant)
    pAryColRngs = vArray
End Property
```

```vb
' Caller - must still assign to local variable
Dim myArray As Variant
myArray = tbl.aryColRngs             ' Works
For i = LBound(myArray) To UBound(myArray)
    Debug.Print myArray(i).Address   ' Works
Next i

' STILL FAILS even with Property Get:
For i = LBound(tbl.aryColRngs) To UBound(tbl.aryColRngs)
    Debug.Print tbl.aryColRngs(i).Address  ' Error
Next i
```

## Rules

1. Keep arrays `Private`; expose via methods (`GetKeys()`) or Property Get/Let
2. Always assign to a local variable before iterating
3. Never directly index into a class attribute (`obj.attr(i)`)
4. Return individual values through single-item methods (e.g., `Item(key)`)

See `Dictionary.cls` in ExcelSteps for a working implementation.
