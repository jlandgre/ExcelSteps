'ExcelSteps_frmRefresh.vb
'Version 10/24/24
Option Explicit
Private Const iShtsMax As Integer = 8 'limit of dialog box
'-----------------------------------------------------------------------------------------------------
' Initialize the form
' Modified 9/23/21 JDL - Add Scenario Model refresh capability
'
Function RefreshInit(wkbk) As Integer
    Dim iShts As Integer, w As Variant, aryExcludeShts As Variant
    Dim SheetLbls As Variant, RefreshBxes As Variant, ReplaceBxes As Variant
    Dim SMBxes As Variant, AppendShtNmBxes As Variant, IsCalcBxes As Variant
    
    'Populate arrays to simplify iterating And initialize controls
    PopulateArrays SheetLbls, RefreshBxes, ReplaceBxes, SMBxes, AppendShtNmBxes, IsCalcBxes
    For iShts = 0 To iShtsMax - 1
        EnableDialogItems iShts, SheetLbls, RefreshBxes, ReplaceBxes, SMBxes, "", False
    Next iShts
    aryExcludeShts = Array(shtSteps, shtSettings, shtLists, shtTblImp)
    
    'Add sheet names to form if not in exclude list and not xlVeryHidden
    iShts = 0
    For Each w In wkbk.Sheets
        If Not IsInAry(aryExcludeShts, w.Name) And w.Visible <> xlVeryHidden Then
            If iShts <= iShtsMax - 1 Then
                EnableDialogItems iShts, SheetLbls, RefreshBxes, ReplaceBxes, SMBxes, w.Name, True
                iShts = iShts + 1
            End If
        End If
    Next w
    
    'Reload previous settings or add a Settings sheet if there is none
    If Not SheetExists(wkbk, shtSettings) Then
        AddSettingsSht wkbk
    ElseIf Not ApplySettings(wkbk) Then
        GoTo ErrorExit
    End If
    Exit Function
    
ErrorExit:
    MsgBox "Error in frmRefresh.RefreshInit"
End Function
'-----------------------------------------------------------------------------------------------------
'Purpose:   OK Button - refreshes selected rows/cols table and/or Scenario Models
'
'Created:   January 2020 JDL      Modified: 2/13/25 update comments
'
Private Sub CmdButOK_Click()
    Dim refr As New Refresh, mdl As Object, tblSteps As New tblRowsCols
    
    Dim IsTblRefresh As Boolean, IsMdlRefresh As Boolean, w As Variant, i As Integer
    Dim SheetLbls As Variant, RefreshBxes As Variant, ReplaceBxes As Variant
    Dim SMBxes As Variant, AppendShtNmBxes As Variant, IsCalcBxes As Variant
    Dim xCalculation As Integer, shtCurrent As String

    SetErrs "driver": If errs.IsHandle Then On Error GoTo ErrorExit
    
    'Hide frmRefresh dialog
    Hide

    'Populate clsRefresh instance and write dialog box state to Settings_ sheet
    With refr
    
        'xxx why here? JDL 10/23/24 --loop below performs refreshes on individual tables and models
        If Not .InitMdl(refr, ActiveWorkbook) Then GoTo ErrorExit
        
        ' Save the dialog box selections to Settings to use for repopulating in future refreshes
        If Not WriteSettings(refr) Then GoTo ErrorExit
        
        'Arrays to aid iteration over form's CheckBoxes
        PopulateArrays SheetLbls, RefreshBxes, ReplaceBxes, SMBxes, AppendShtNmBxes, IsCalcBxes
        
        'Exit if no refresh selected - xxx don't need local boolean variables?
        IsTblRefresh = IsAnyRefresh(RefreshBxes)
        IsMdlRefresh = IsAnyRefresh(SMBxes)
        If Not (IsTblRefresh Or IsMdlRefresh) Then Exit Sub
        
        'Set status for performance during execution
        wkbkResetStatus True, .wkbk, xCalculation, shtCurrent
        
        'Always reformat tables when refreshed from ExcelSteps menu
        .IsTblFormat = True
        
        'ExcelSteps Prep - needed if any tblRowsCols is being refreshed
        'xxx this is taken care of by .RefreshRC so not needed here
        If IsTblRefresh Then If Not .PrepExcelStepsSht(refr, tblSteps) Then GoTo ErrorExit
            
        'Iterate through sheet labels and refresh as specified
        For i = 0 To iShtsMax - 1
        
            '10/21/24 brought into this function because of unexplained VBA Internal Error if call SetRefreshAttributes
            If False Then
                If Not SetRefreshAttributes(refr, i, SheetLbls, AppendShtNmBxes, SMBxes, IsCalcBxes, ReplaceBxes, RefreshBxes) Then GoTo ErrorExit
                '.sht = SheetLbls(i)
            Else
                With refr
                    .IsTbl = False
                    .IsMdl = False
                    
                    .sht = SheetLbls(i).Caption
                    If Len(.sht) > 0 Then
                        If RefreshBxes(i) Then
                            .IsTbl = True
                            .IsReplace = ReplaceBxes(i)
                        ElseIf SMBxes(i) Then
                            .IsMdl = True
                            .IsCalc = IsCalcBxes(i)
                        End If
                        .IsNamePrefix = AppendShtNmBxes(i)
                    End If
                End With
            End If
            
            'Rows/Cols refresh (frmRefresh only handles default tbl refresh so just sht specified)
            If .IsTbl Then
                    If Not .RefreshRC(refr, tblSteps, sht:=.sht, IsNamePrefix:=.IsNamePrefix) Then GoTo ErrorExit
                        
            'Scenario model refresh
            ElseIf .IsMdl Then
                Set mdl = New mdlScenario
                mdl.Provision mdl, .wkbk, .sht, IsMdlNmPrefix:=.IsNamePrefix, IsCalc:=IsCalcBxes(i)
                If Not mdl.Refresh(mdl) Then GoTo ErrorExit
            End If
        Next i
        DeleteUnusedNames .wkbk
        wkbkResetStatus False, .wkbk, xCalculation, shtCurrent
    End With
    Exit Sub
    
ErrorExit:
    Hide
    If refr.wkbk Is Nothing Then Set refr.wkbk = ActiveWorkbook
    wkbkResetStatus False, refr.wkbk, xCalculation, shtCurrent
    errs.RecordErr "CmdButOK_Click"
End Sub
Function IsAnyRefresh(aryBxes)
    Dim w As Variant
    For Each w In aryBxes
        If w Then IsAnyRefresh = True
    Next w
End Function
'-----------------------------------------------------------------------------------------------------
' Set refresh attributes based on dialog box row i checkbox settings
' JDL 12/15/22
'
Public Function SetRefreshAttributes(refr, i, SheetLbls, AppendShtNmBxes, SMBxes, IsCalcBxes, ReplaceBxes, RefreshBxes) As Boolean
    If errs.IsHandle Then On Error GoTo ErrorExit
    Const Locn = "frmRefresh.SetRefreshAttributes": SetRefreshAttributes = True
           
    With refr
        .IsTbl = False
        .IsMdl = False
        
        .sht = SheetLbls(i).Caption
        If Len(.sht) > 0 Then
            If RefreshBxes(i) Then
                .IsTbl = True
                .IsReplace = ReplaceBxes(i)
            ElseIf SMBxes(i) Then
                .IsMdl = True
                .IsCalc = IsCalcBxes(i)
            End If
            .IsAppendName = AppendShtNmBxes(i)
        End If
    End With
    Exit Function
ErrorExit:
    errs.RecordErr Locn
    SetRefreshAttributes = False
End Function
'-----------------------------------------------------------------------------------------------------
' Sub CmdButCancel_Click - Cancel the dialog
'
Private Sub CmdButCancel_Click()
    Hide
End Sub
'-----------------------------------------------------------------------------------------------------
' Sub EnableDialogItems - Enable or disable item i in arrays of dialog items
' Call with .sCaption blank when disabling
'
Private Sub EnableDialogItems(i, ShtLbls, RefBxes, ReplBxes, SMBxes, sCaption, bEnable)
    ShtLbls(i).Enabled = bEnable
    RefBxes(i).Enabled = bEnable
    ReplBxes(i).Enabled = bEnable
    SMBxes(i).Enabled = bEnable
    ShtLbls(i).Caption = sCaption
End Sub
'
' Sub Sub PopulateArrays - Populate arrays of dialog items
'
Private Sub PopulateArrays(SheetLbls, RefreshBxes, ReplaceBxes, SMBxes, AppendShtNmBxes, IsCalcBxes)
    SheetLbls = Array(lblSht_1, lblSht_2, lblSht_3, lblSht_4, lblSht_5, lblSht_6, _
        lblSht_7, lblSht_8)
    RefreshBxes = Array(CkBoxSht_1, CkBoxSht_2, CkBoxSht_3, CkBoxSht_4, CkBoxSht_5, _
        CkBoxSht_6, CkBoxSht_7, CkBoxSht_8)
    ReplaceBxes = Array(CkBoxRepl_1, CkBoxRepl_2, CkBoxRepl_3, CkBoxRepl_4, CkBoxRepl_5, _
        CkBoxRepl_6, CkBoxRepl_7, CkBoxRepl_8)
    SMBxes = Array(CkBoxSM_1, CkBoxSM_2, CkBoxSM_3, CkBoxSM_4, CkBoxSM_5, _
        CkBoxSM_6, CkBoxSM_7, CkBoxSM_8)
    AppendShtNmBxes = Array(CkBoxShtNm_1, CkBoxShtNm_2, CkBoxShtNm_3, CkBoxShtNm_4, CkBoxShtNm_5, _
        CkBoxShtNm_6, CkBoxShtNm_7, CkBoxShtNm_8)
    IsCalcBxes = Array(CkBoxSMCalc_1, CkBoxSMCalc_2, CkBoxSMCalc_3, CkBoxSMCalc_4, CkBoxSMCalc_5, _
        CkBoxSMCalc_6, CkBoxSMCalc_7, CkBoxSMCalc_8)
End Sub
'-----------------------------------------------------------------------------------
' Write form settings to Settings_ sheet
' Modified JDL 7/10/24 fix bug where while loop omitted writing last sheet settings
'
Function WriteSettings(refr) As Boolean
    Dim i As Integer, SheetLbls As Variant, RefreshBxes As Variant, ReplaceBxes As Variant
    Dim SMBxes As Variant, AppendShtNmBxes As Variant, IsCalcBxes As Variant
    If errs.IsHandle Then On Error GoTo ErrorExit
    Const Locn = "WriteSettings": WriteSettings = True
    
    'Populate arrays with current states of checkboxes on form
    PopulateArrays SheetLbls, RefreshBxes, ReplaceBxes, SMBxes, AppendShtNmBxes, IsCalcBxes
    
    'Loop through worksheet labels and update their status to Settings
    For i = 0 To UBound(SheetLbls)
        If Len(SheetLbls(i).Caption > 0) Then
            UpdateSetting refr.wkbk, setting_shtFrm, SheetLbls(i).Caption, RefreshBxes(i).value, _
                ReplaceBxes(i).value, SMBxes(i).value, IsCalcBxes(i).value, AppendShtNmBxes(i).value
        End If
    Next i
    Exit Function
    
ErrorExit:
    errs.RecordErr Locn
    WriteSettings = False
End Function
Sub UpdateSetting(wkbk, sType, txtVal, bool1, bool2, bool3, bool4, bool5)
    Dim curCell As Range
    
    With wkbk.Sheets(shtSettings)
        
        'Update Sheet refresh settings
        If sType = setting_shtFrm Then
            
            'Find the existing setting row or add a new setting if it is not found
            Set curCell = .Columns(2).Find(txtVal, lookat:=xlWhole)
            If curCell Is Nothing Then
                Set curCell = .Cells(.Rows.Count, 1).End(xlUp).Offset(1, 0)
            Else
                Set curCell = curCell.Offset(0, -1)
            End If
            
            'Update the setting's values
            curCell = sType
            curCell.Offset(0, 1) = txtVal
            curCell.Offset(0, 2) = bool1
            curCell.Offset(0, 3) = bool2
            curCell.Offset(0, 4) = bool3
            curCell.Offset(0, 5) = bool4
            curCell.Offset(0, 6) = bool5
        
        'Update Form checkbox settings
        ElseIf sType = setting_cBoxFrm1 Then
        
            'Find the existing setting row or add a new setting if it is not found
            Set curCell = .Columns(1).Find(setting_cBoxFrm1, lookat:=xlWhole)
            If curCell Is Nothing Then Set curCell = .Cells(.Rows.Count, 1).End(xlUp).Offset(1, 0)
            
            'Update the setting's values
            curCell = sType
            curCell.Offset(0, 2) = bool1
        End If
    End With
End Sub
Sub AddSettingsSht(wkbk)
    AddSheet wkbk, shtSettings, wkbk.Sheets(wkbk.Sheets.Count).Name
    With wkbk.Sheets(shtSettings)
        Range(.Cells(1, 1), .Cells(1, 7)) = Array("Type", "txtval", "bool1", "bool2", "bool3", "bool4", "bool5")
        .Visible = xlSheetVeryHidden
    End With
End Sub
Function ApplySettings(wkbk) As Boolean
    Dim curCell As Range, delRow As Range
    Dim sName As String, sType As String, sTxtVal As String
    Dim bool1 As Boolean, bool2 As Boolean, bool3 As Boolean, bool4 As Boolean, bool5 As Boolean
    On Error GoTo ErrorExit
    Const Locn = "ApplySettings": ApplySettings = True
        
    'Read and apply existing settings
    With wkbk.Sheets(shtSettings)
        Set curCell = .Cells(2, 1)
        While Len(curCell) > 0
            ReadSetting_frmRefresh curCell, sType, sTxtVal, bool1, bool2, bool3, bool4, bool5
            If sType = setting_shtFrm Then
            
                'Either update form based on setting values or delete setting if sheet doesn't exist
                If Not UpdateShtOnFrm(sTxtVal, bool1, bool2, bool3, bool4, bool5) Then
                    Set delRow = curCell.EntireRow
                    Set curCell = curCell.Offset(-1, 0)
                    delRow.Delete
                End If
            End If
            Set curCell = curCell.Offset(1, 0)
        Wend
    End With
    ApplySettings = 1
    Exit Function
    
ErrorExit:
    errs.RecordErr Locn
    ApplySettings = False
End Function
Sub ReadSetting_frmRefresh(rngCell, sType, sTxtVal, bool1, bool2, bool3, bool4, bool5)
    With rngCell.EntireRow
        sType = .Cells(1)
        sTxtVal = .Cells(2)
        bool1 = .Cells(3)
        bool2 = .Cells(4)
        bool3 = .Cells(5)
        bool4 = .Cells(6)
        bool5 = .Cells(7)
    End With
End Sub
Function UpdateShtOnFrm(sht, bRefresh, bReplace, bScenModel, bSMIsCalc, bAppendShtNm) As Boolean
    Dim i As Integer
    Dim SheetLbls As Variant, RefreshBxes As Variant, ReplaceBxes As Variant
    Dim SMBxes As Variant, AppendShtNmBxes As Variant, IsCalcBxes As Variant
    
    'Populate arrays with current states of checkboxes on form
    PopulateArrays SheetLbls, RefreshBxes, ReplaceBxes, SMBxes, AppendShtNmBxes, IsCalcBxes

    UpdateShtOnFrm = False
    
    'If sheet name is listed in form's labels, update form settings for the sheet
    If IsInAry(SheetLbls, sht) Then
        i = 0
        
        'Loop through form labels to find the index, i, to update
        While Not UpdateShtOnFrm
            If SheetLbls(i) = sht Then
                RefreshBxes(i).value = bRefresh
                ReplaceBxes(i).value = bReplace
                SMBxes(i).value = bScenModel
                IsCalcBxes(i).value = bSMIsCalc
                AppendShtNmBxes(i).value = bAppendShtNm
                UpdateShtOnFrm = True
            End If
            i = i + 1
        Wend
    End If
End Function
Private Sub TextBox1_AfterUpdate()
    Me.CmdButOK.SetFocus
End Sub
'-----------------------------------------------------------------------------------------------------
'Purpose:   Toggle Refresh/Replace button status if Scenario Model enabled and vice versa
'
'Created:   9/23/21
'

Private Sub CkBoxSM_1_Click()
    If CkBoxSM_1 Then
        CkBoxSht_1 = False
        CkBoxRepl_1 = False
        CkBoxSht_1.Enabled = False
        CkBoxRepl_1.Enabled = False
        CkBoxSMCalc_1.Enabled = True
    Else
        CkBoxSMCalc_1 = False
        CkBoxSMCalc_1.Enabled = False
        CkBoxSht_1.Enabled = True
        CkBoxRepl_1.Enabled = True

    End If
End Sub
'-----------------------------------------------------------------------------------------------------
' Handle CheckBox Click events for rows/cols refresh sheets 1 to 8
' JDL 7/10/24
'-----------------------------------------------------------------------------------------------------
Private Sub CkBoxSht_1_Click()
    HandleCkBoxShtClick CkBoxSht_1, CkBoxSM_1, CkBoxSMCalc_1, CkBoxRepl_1, CkBoxShtNm_1
End Sub
Private Sub CkBoxSht_2_Click()
    HandleCkBoxShtClick CkBoxSht_2, CkBoxSM_2, CkBoxSMCalc_2, CkBoxRepl_2, CkBoxShtNm_2
End Sub
Private Sub CkBoxSht_3_Click()
    HandleCkBoxShtClick CkBoxSht_3, CkBoxSM_3, CkBoxSMCalc_3, CkBoxRepl_3, CkBoxShtNm_3
End Sub
Private Sub CkBoxSht_4_Click()
    HandleCkBoxShtClick CkBoxSht_4, CkBoxSM_4, CkBoxSMCalc_4, CkBoxRepl_4, CkBoxShtNm_4
End Sub
Private Sub CkBoxSht_5_Click()
    HandleCkBoxShtClick CkBoxSht_5, CkBoxSM_5, CkBoxSMCalc_5, CkBoxRepl_5, CkBoxShtNm_5
End Sub
Private Sub CkBoxSht_6_Click()
    HandleCkBoxShtClick CkBoxSht_6, CkBoxSM_6, CkBoxSMCalc_6, CkBoxRepl_6, CkBoxShtNm_6
End Sub
Private Sub CkBoxSht_7_Click()
    HandleCkBoxShtClick CkBoxSht_7, CkBoxSM_7, CkBoxSMCalc_7, CkBoxRepl_7, CkBoxShtNm_7
End Sub
Private Sub CkBoxSht_8_Click()
    HandleCkBoxShtClick CkBoxSht_8, CkBoxSM_8, CkBoxSMCalc_8, CkBoxRepl_8, CkBoxShtNm_8
End Sub
'-----------------------------------------------------------------------------------------------------
' Handle CkBoxSht Click
' JDL 7/10/24
'
Sub HandleCkBoxShtClick(CkBoxSht, CkBoxSM, CkBoxSMCalc, CkBoxRepl, CkBoxShtNm)
    If CkBoxSht Then SetButtonStates_RCRefresh CkBoxSM, CkBoxSMCalc
    If Not CkBoxSht Then SetButtonStates_NotRCRefresh CkBoxSM, CkBoxSMCalc, CkBoxRepl, CkBoxShtNm
End Sub
'-----------------------------------------------------------------------------------------------------
' Set Checkbox values and .Enabled states if user selects sheet's rows/cols refresh
' (Turn off and disable Scenario Model checkboxes)
' JDL 7/10/24
'
Sub SetButtonStates_RCRefresh(CkBoxSM, CkBoxSMCalc)
    CkBoxSM.value = False
    CkBoxSM.Enabled = False
    CkBoxSMCalc.value = False
    CkBoxSMCalc.Enabled = False
End Sub
'-----------------------------------------------------------------------------------------------------
' Set Scenario Model Checkbox values and .Enabled states if user unselects sheet's rows/cols refresh
' JDL 7/10/24
'
Sub SetButtonStates_NotRCRefresh(CkBoxSM, CkBoxSMCalc, ByRef CkBoxSMRepl, ByRef CkboxSMShtNm)
    CkBoxSM.Enabled = True
    CkBoxSMCalc.Enabled = True
    CkBoxSMRepl.value = False
    CkboxSMShtNm.value = False
End Sub
'-----------------------------------------------------------------------------------------------------
' Handle other button click events - xxx refactor modeled off of CkBoxSht code above
'-----------------------------------------------------------------------------------------------------
Private Sub CkBoxRepl_1_Click()
    If CkBoxRepl_1 Then
        CkBoxSM_1 = False
        CkBoxSht_1 = True
    End If
End Sub
Private Sub CkBoxSM_2_Click()
    If CkBoxSM_2 Then
        CkBoxSht_2 = False
        CkBoxRepl_2 = False
        CkBoxSht_2.Enabled = False
        CkBoxRepl_2.Enabled = False
        CkBoxSMCalc_2.Enabled = True
    Else
        CkBoxSMCalc_2 = False
        CkBoxSMCalc_2.Enabled = False
        CkBoxSht_2.Enabled = True
        CkBoxRepl_2.Enabled = True
    End If
End Sub
Private Sub CkBoxRepl_2_Click()
    If CkBoxRepl_2 Then
        CkBoxSM_2 = False
        CkBoxSht_2 = True
    End If
End Sub

Private Sub CkBoxSM_3_Click()
    If CkBoxSM_3 Then
        CkBoxSht_3 = False
        CkBoxRepl_3 = False
        CkBoxSht_3.Enabled = False
        CkBoxRepl_3.Enabled = False
        CkBoxSMCalc_3.Enabled = True
    Else
        CkBoxSMCalc_3 = False
        CkBoxSMCalc_3.Enabled = False
        CkBoxSht_3.Enabled = True
        CkBoxRepl_3.Enabled = True
    End If
End Sub
Private Sub CkBoxRepl_3_Click()
    If CkBoxRepl_3 Then
        CkBoxSM_3 = False
        CkBoxSht_3 = True
    End If
End Sub
Private Sub CkBoxSM_4_Click()
    If CkBoxSM_4 Then
        CkBoxSht_4 = False
        CkBoxRepl_4 = False
        CkBoxSht_4.Enabled = False
        CkBoxRepl_4.Enabled = False
        CkBoxSMCalc_4.Enabled = True
    Else
        CkBoxSMCalc_4 = False
        CkBoxSMCalc_4.Enabled = False
        CkBoxSht_4.Enabled = True
        CkBoxRepl_4.Enabled = True

    End If
End Sub
Private Sub CkBoxRepl_4_Click()
    If CkBoxRepl_4 Then
        CkBoxSM_4 = False
        CkBoxSht_4 = True
    End If
End Sub
Private Sub CkBoxSM_5_Click()
    If CkBoxSM_5 Then
        CkBoxSht_5 = False
        CkBoxRepl_5 = False
        CkBoxSht_5.Enabled = False
        CkBoxRepl_5.Enabled = False
        CkBoxSMCalc_5.Enabled = True
    Else
        CkBoxSMCalc_5 = False
        CkBoxSMCalc_5.Enabled = False
        CkBoxSht_5.Enabled = True
        CkBoxRepl_5.Enabled = True

    End If
End Sub
Private Sub CkBoxRepl_5_Click()
    If CkBoxRepl_5 Then
        CkBoxSM_5 = False
        CkBoxSht_5 = True
    End If
End Sub
Private Sub CkBoxSM_6_Click()
    If CkBoxSM_6 Then
        CkBoxSht_6 = False
        CkBoxRepl_6 = False
        CkBoxSht_6.Enabled = False
        CkBoxRepl_6.Enabled = False
        CkBoxSMCalc_6.Enabled = True
    Else
        CkBoxSMCalc_6 = False
        CkBoxSMCalc_6.Enabled = False
        CkBoxSht_6.Enabled = True
        CkBoxRepl_6.Enabled = True

    End If
End Sub
Private Sub CkBoxRepl_6_Click()
    If CkBoxRepl_6 Then
        CkBoxSM_6 = False
        CkBoxSht_6 = True
    End If
End Sub
Private Sub CkBoxSM_7_Click()
    If CkBoxSM_7 Then
        CkBoxSht_7 = False
        CkBoxRepl_7 = False
        CkBoxSht_7.Enabled = False
        CkBoxRepl_7.Enabled = False
        CkBoxSMCalc_7.Enabled = True
    Else
        CkBoxSMCalc_7 = False
        CkBoxSMCalc_7.Enabled = False
        CkBoxSht_7.Enabled = True
        CkBoxRepl_7.Enabled = True

    End If
End Sub
Private Sub CkBoxRepl_7_Click()
    If CkBoxRepl_7 Then
        CkBoxSM_7 = False
        CkBoxSht_7 = True
    End If
End Sub
Private Sub CkBoxSM_8_Click()
    If CkBoxSM_8 Then
        CkBoxSht_8 = False
        CkBoxRepl_8 = False
        CkBoxSht_8.Enabled = False
        CkBoxRepl_8.Enabled = False
        CkBoxSMCalc_8.Enabled = True
    Else
        CkBoxSMCalc_8 = False
        CkBoxSMCalc_8.Enabled = False
        CkBoxSht_8.Enabled = True
        CkBoxRepl_8.Enabled = True

    End If
End Sub
Private Sub CkBoxRepl_8_Click()
    If CkBoxRepl_8 Then
        CkBoxSM_8 = False
        CkBoxSht_8 = True
    End If
End Sub

Private Sub UserForm_Click()

End Sub