Attribute VB_Name = "xlFuncs"
Option Explicit
'-----------------------------------------------------------------------------------------------------
'User defined Excel function returns value offset from calling cell
'
'Created: 10/5/22 JDL
'
Public Function ROWSHIFTVAL(iShiftRow, Optional sCol) As Variant
    Application.Volatile
    Dim iRow As Long, iCol As Long, wkbk As Workbook, sht As String, rngCol As Range
    
    ROWSHIFTVAL = "#NOT FOUND"
    
    'Locate the calling cell and its location
    With Application.Caller
        iRow = .Row
        iCol = .Column
        Set wkbk = .Parent.Parent
        sht = .Parent
    End With
    
    'Exit if iShiftRow non-numeric or leads to invalid row
    If Not VarType(iShiftRow) = vbInteger Or Not VarType(iShiftRow) = vbLong Then Exit Function
    If iRow + iShiftRow < 1 Then Exit Function
    
    With wkbk.Sheets(sht)
    
        'If col string is specified, locate the column and reset iCol
        If Not IsMissing(sCol) Then
            Set rngCol = .Rows(1).Find(sCol, lookat:=xlWhole)
            If rngCol Is Nothing Then Exit Function
            iCol = rngCol.Column
        End If
        
        ROWSHIFTVAL = .Cells(iRow + iShiftRow, iCol)
    End With
End Function
Sub tester()
    MsgBox ROWSHIFTVAL(-1)
End Sub
