Attribute VB_Name = "Interface"
'Version 1/29/26
Option Explicit
'---------------------------------------------------------------------------------------------
'---------------------------------------------------------------------------------------------
' ParseModel
'---------------------------------------------------------------------------------------------
'---------------------------------------------------------------------------------------------
' Driver sub to parse Scenario Model on active sheet
' JDL 11/19/25
'
Sub ParseModelDriver()
    SetErrs "driver", ThisWorkbook: If errs.IsHandle Then On Error GoTo ErrorExit
    Dim mdl As New mdlScenario, IsCalcMdl As Boolean
    
    SetApplEnvir False, False, xlCalculationManual
    
    'Validate and detect multicolum or calculator default model
    With ActiveSheet
        If errs.IsFail(.Cells(1, 1) <> "Grp", 1) Then GoTo ErrorExit
        If .Cells(2, 9) = "Calculator" Then
            IsCalcMdl = True
        Else
            IsCalcMdl = False
            If errs.IsFail(.Cells(2, 4) <> "Scenario", 2) Then GoTo ErrorExit
        End If
    End With
        
    If Not mdl.Provision(mdl, ActiveWorkbook, sht:=ActiveSheet.Name, IsCalc:=IsCalcMdl) _
        Then GoTo ErrorExit
    If Not ParseMdl(mdl) Then GoTo ErrorExit
    
    SetApplEnvir True, True, xlCalculationAutomatic
    Exit Sub
    
ErrorExit:
    errs.RecordErr "ParseModelDriver"
End Sub
'---------------------------------------------------------------------------------------------
'Create menu when the add-vin opens
Private Sub Auto_Open()
    Dim HelpMenu As CommandBarControl, SubMenuItem As CommandBarButton
    Dim ExcelSteps As CommandBarPopup, MenuItem As CommandBarControl

    ' Exit if running on a Mac
    If InStr(Application.OperatingSystem, "Macintosh") > 0 Then Exit Sub
    
    'Create menu if none exists
    Set ExcelSteps = CommandBars(1).FindControl(Tag:="ExcelSteps")
    If (ExcelSteps Is Nothing) Then
    
        'add ExcelSteps To the ribbon; if there is help menu, add before Help
        Set HelpMenu = CommandBars(1).FindControl(ID:=30010)
        If (HelpMenu Is Nothing) Then
            Set ExcelSteps = CommandBars(1).Controls.Add(Type:=msoControlPopup, Temporary:=True)
        Else
            Set ExcelSteps = CommandBars(1).Controls.Add(Type:=msoControlPopup, _
                before:=HelpMenu.index, Temporary:=True)
        End If
        ExcelSteps.Caption = "&ExcelSteps"
        ExcelSteps.Tag = "ExcelSteps"
    
        'Add menu item: Refresh Worksheets
        Set MenuItem = ExcelSteps.Controls.Add(Type:=msoControlButton)
        MenuItem.Caption = sRefresh
        MenuItem.OnAction = "RefreshDriver"
    
        'Add menu item: Parse Scenario Model
        Set MenuItem = ExcelSteps.Controls.Add(Type:=msoControlButton)
        MenuItem.Caption = sParseSM
        MenuItem.OnAction = "ParseModelDriver"
    End If
End Sub

'---------------------------------------------------------------------------------------------
'Show Refresh dialog to allow user selections; OK Button calls RefreshRC
'
'Created:   7/29/21 JDL
'
'Inputs: OK Button or as API
'
' Sub RefreshDriver
'
' Called by ExcelSteps refresh menu item (modInterface). RefreshDriver instances and shows the
' dialog box to receive user inputs. The form's OK button calls RefreshRC function.  This is
' the user-initiated way to refresh. RefreshRC can also be called as an API by a driver program
'
Public Sub RefreshDriver()
    Dim frm As New frmRefresh, i As Integer
    If ActiveWorkbook Is Nothing Then End
    i = frm.RefreshInit(ActiveWorkbook)
    frm.CmdButOK.SetFocus
    frm.Show vbModal
    frm.Hide
    Set frm = Nothing
End Sub
'-----------------------------------------------------------------------------------------------------
' API for refreshing tables
' Either sht or (TblName or Defn) must be specified for default or custom table
' JDL 10/24/24 validated in ExcelSteps_Validation.xlsm modTests_tbl module
'              12/2/25 Modify RecordErr call to not be driver
'
Function RefreshTblAPI(wkbkR, IsReplace, IsTblFormat, _
    Optional sht, Optional TblName, Optional Defn, Optional rcHome, Optional nRows, _
    Optional nCols, Optional iOffsetKeyCol, Optional iOffsetHeader, Optional IsSetAryCols, _
    Optional IsSetColRngs, Optional IsSetTblNames, Optional IsSetColNames, _
    Optional IsNamePrefix, Optional IsPrefixSht, Optional NamePrefix) As Boolean
    
    SetErrs RefreshTblAPI: If errs.IsHandle Then On Error GoTo ErrorExit
    Dim refr As New Refresh, tblSteps As New tblRowsCols, mdl As New mdlScenario
    
    With refr

        'Initialize Refrsesh (aka tblRowsCols) attributes
        If Not .InitTbl(refr, wkbkR, IsReplace:=IsReplace, IsTblFormat:=IsTblFormat) Then GoTo ErrorExit

        'Refresh the table and clean up names
        If Not .RefreshRC(refr, tblSteps, sht:=sht, TblName:=TblName, Defn:=Defn, rcHome:=rcHome, nRows:=nRows, nCols:=nCols, _
            iOffsetKeyCol:=iOffsetKeyCol, iOffsetHeader:=iOffsetHeader, IsSetAryCols:=IsSetAryCols, _
            IsSetColRngs:=IsSetColRngs, IsSetTblNames:=IsSetTblNames, IsSetColNames:=IsSetColNames, _
            IsNamePrefix:=IsNamePrefix, IsPrefixSht:=IsPrefixSht, NamePrefix:=NamePrefix) Then GoTo ErrorExit
        DeleteUnusedNames .wkbk
    End With
    Exit Function
    
ErrorExit:
    errs.RecordErr "RefreshTblAPI", RefreshTblAPI
End Function
'---------------------------------------------------------------------------------------------
' Set application status to optimize performance and remember initial, active sheet
' JDL updated 12/19/22; 12/13/24 add check of whether shtCurrent is initialized before setting
'
Sub wkbkResetStatus(IsRefreshInit, wkbk, xCalculation, shtCurrent)
    
    If Len(shtCurrent) < 1 Then shtCurrent = wkbk.ActiveSheet.Name
    
    With Application
        .ScreenUpdating = Not IsRefreshInit
        .EnableEvents = Not IsRefreshInit
    End With
    
    If IsRefreshInit Then
        xCalculation = Application.Calculation
        Application.Calculation = xlCalculationManual
    Else
        Application.Calculation = xCalculation
        If SheetExists(wkbk, shtCurrent) Then wkbk.Sheets(shtCurrent).Activate
    End If
End Sub
