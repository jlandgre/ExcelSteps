VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ProjFiles"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'Version 4/29/26

Public wkbkProj As Workbook ' Project workbook
Public IsDevelopment As Boolean ' True if working in development, src/tests folder structure
Public subdir_tests As String ' tests subdirectory containing test data
Public pathSrc As String ' path to src directory (or directory containing project workbook)
Public fWkbkProj As String ' Project workbook filename
Public pathRoot As String ' path to root directory
Public pathTests As String ' path to tests directory
Public pathData As String ' path to data directory (tests/subdir_tests if specified)
Public pathColInfo As String ' path to ColInfo file
Public fColInfo As String ' ColInfo filename (typically ColInfo.xlsx)
Public pfColInfo As String ' ColInfo path + filename
'-----------------------------------------------------------------------------------------
' Set path/filename attributes for project workbook and test suite locations
' JDL 4/29/26; updated 5/1/26
'
Public Function Init(files, ByVal wkbk As Workbook, Optional subdir_tests As String) As Boolean
    SetErrs Init: If errs.IsHandle Then On Error GoTo ErrorExit
    With files
        Set .wkbkProj = wkbk
        .fWkbkProj = .wkbkProj.Name
        .subdir_tests = subdir_tests
    End With
    
    'Set paths and filenames
    If Not SetGenericPaths(files) Then GoTo ErrorExit
    'If Not SetProjSpecificPaths(files) Then GoTo ErrorExit
    Exit Function

ErrorExit:
    errs.RecordErr "Init", Init
End Function
'-----------------------------------------------------------------------------------------
' Set generic paths and filenames
' JDL 4/29/26; updated 5/1/26
'
Private Function SetGenericPaths(files) As Boolean
    SetErrs SetGenericPaths: If errs.IsHandle Then On Error GoTo ErrorExit
    Dim sep As String, aryPath As Variant
    sep = Application.PathSeparator
    
    With files
        .pathSrc = .wkbkProj.Path & sep

        ' Set flag for Development folder structure (root, src, tests)
        aryPath = Split(.pathSrc, sep)
        .IsDevelopment = aryPath(UBound(aryPath) - 1) = "src"

        If .IsDevelopment Then
            ReDim Preserve aryPath(0 To UBound(aryPath) - 2)
            .pathRoot = ListFromArray(aryPath, sep) & sep
            .pathTests = .pathRoot & "tests" & sep
        End If
    
        If Not IsTest Then
            .pathData = .pathSrc
            .pathColInfo = .pathSrc
        ElseIf .subdir_tests = "" Then
            .pathData = .pathTests
            .pathColInfo = .pathTests
        Else
            .pathData = .pathTests & .subdir_tests & sep
            .pathColInfo = .pathData
        End If
        .fColInfo = "ColInfo.xlsx"
        .pfColInfo = .pathColInfo & .fColInfo
    End With
    Exit Function

ErrorExit:
    errs.RecordErr "SetGenericPaths", SetGenericPaths
End Function
'-----------------------------------------------------------------------------------------
' Set project-specific paths and filenames
' JDL 5/1/26
'
Private Function SetProjSpecificPaths(files) As Boolean
    SetErrs SetProjSpecificPaths: If errs.IsHandle Then On Error GoTo ErrorExit
    
    With files
    
    End With
    Exit Function
    
ErrorExit:
    errs.RecordErr "SetProjSpecificPaths", SetProjSpecificPaths
End Function


