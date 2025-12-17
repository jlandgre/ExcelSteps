'ExcelSteps_ThisWorkbook.vb
'Version 12/16/25
Option Explicit
#If VBA7 And Win64 Then
    ' 64-bit Windows
    #Const Win64 = True
#ElseIf VBA7 And Win32 Then
    ' 32-bit Windows
    #Const Win32 = True
#ElseIf Mac Then
    ' Mac
    #Const Mac = True
#End If
