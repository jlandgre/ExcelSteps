VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ParamBlock"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Version 1/29/26
' This class enables the feature of adding a parameter block in left-most columns of
' a scenario model's .colrngModel in columns that do not contain named scenarios.
' wayfinding for the block can either be by naming the entire column(s)
' or just the block's rows based on the IsNameWholeCol argument. Blocks can be single
' column or multicolumn based on the nCols attribute
Option Explicit

Public mdl As Object           ' mdl object where the ParamBlock will reside
Public heading As String       ' Param Block locator string in mdl.rngRows
Public iColMdl As Long         ' relative column in mdl.colRngModel
Public nRows As Long           ' number rows in the block
Public nCols As Long           ' number columns in the block
Public iRowOffsetHome As Long  ' number rows cellHome offset from header string
Public IsNameWholeCol As Boolean ' flag name the whole column or just block rows
Public TblNameBlock As String  ' name for the parameter block
'-----------------------------------------------------------------------------------------
' Initialize ParamBlock
' JDL 7/30/25
'
Public Function Init(pblock As Object, mdlObj As Object, heading As String, _
    iColMdl As Long, nRows As Long, nCols As Long, iRowOffsetHome As Long, _
    IsNameWholeCol As Boolean, TblNameBlock As String) As Boolean
    SetErrs Init: If errs.IsHandle Then On Error GoTo ErrorExit
    
    With pblock
        Set .mdl = mdlObj
        .heading = heading
        .iColMdl = iColMdl
        .nRows = nRows
        .nCols = nCols
        .iRowOffsetHome = iRowOffsetHome
        .IsNameWholeCol = IsNameWholeCol
        .TblNameBlock = TblNameBlock
    End With
    Exit Function
    
ErrorExit:
    errs.RecordErr "ParamBlock.Init", Init
End Function
'-----------------------------------------------------------------------------------------
'Provision parameter block to name its range
' JDL 7/30/25
'
Public Function ProvisionBlock(pblock As Object, tbl As Object) As Boolean
    SetErrs ProvisionBlock: If errs.IsHandle Then On Error GoTo ErrorExit
    
    With pblock
        ' Delete previous range names if any exist
        If Not ExcelSteps.DeleteRngNamesWithPrefix(.mdl.wkbk, .TblNameBlock) Then GoTo ErrorExit
        
        ' Set definition string for the parameter block
        If Not .SetDefn(pblock, tbl) Then GoTo ErrorExit
        
        ' Set tbl.IsNameColRows to name whole column or just block rows
        tbl.IsNameColRows = Not .IsNameWholeCol
        
        ' Provision the table based on defn
        If Not tbl.Provision(tbl, .mdl.wkbk, IsFormat:=True, _
            TblName:=.TblNameBlock, Defn:=tbl.Defn) Then GoTo ErrorExit
    End With
    Exit Function
    
ErrorExit:
    errs.RecordErr "ParamBlock.ProvisionBlock", ProvisionBlock
End Function
'-----------------------------------------------------------------------------------------
'Set custom definition string for parameter block adjacent to Scenario Model
' JDL 7/30/25
'
Public Function SetDefn(pblock As Object, tbl As Object) As Boolean
    SetErrs SetDefn: If errs.IsHandle Then On Error GoTo ErrorExit
    Dim col As Range, c As Range, headingValue As String
    Dim xOffsetRow As Long, flagsStr As String, iCol As Long
    
    With pblock
        ' Find the heading and set the cellHome row offset
        Set col = .mdl.colrngModel.Columns(.iColMdl)
        Set c = ExcelSteps.FindInRange(col, .heading)
        If c Is Nothing Then GoTo ErrorExit
        xOffsetRow = c.Row + .iRowOffsetHome
        
        ' Flags True for IsSetColNames and IsNamePrefix
        flagsStr = "F:T:T:F:F:F"
        
        ' iColMdl gives relative location within .colRngModel
        iCol = .mdl.colrngModel.Columns(.iColMdl).Column

        ' Set defn (format: "Sheet:x,iColMdl:flags:0:-1:nRows:nCols")
        tbl.Defn = .mdl.sht & ":" & xOffsetRow & "," & iCol & ":" & _
                      flagsStr & ":0:-1:" & .nRows & ":" & .nCols
    End With
    Exit Function
    
ErrorExit:
    errs.RecordErr "ParamBlock.SetDefn", SetDefn
End Function
