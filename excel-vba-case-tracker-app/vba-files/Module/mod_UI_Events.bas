Attribute VB_Name = "mod_UI_Events"
Option Explicit

'*****************************************************************************
' Module: mod_UI_Events
' Purpose: This module contains public functions for handling user interface events and interactions.
'*****************************************************************************

Public Sub HandleLoginPanelEvents(buttonName As String, parentFrame As MSForms.Frame)
    If buttonName = "btn_ToOSM" Then
        Dim hasOpen As Boolean, hasClosed As Boolean
        hasOpen = mod_Services.HasAnyOpenCase()
        hasClosed = mod_Services.HasAnyClosedCase()
        
        If Not hasOpen And Not hasClosed Then
            MsgBox MSG_NO_SHEETS_TO_PROCESS, vbInformation, MSG_TITLE_INFO
        ElseIf Not hasOpen And hasClosed Then
            mod_UI_Manager.SwitchTo "Closed"
        Else
            mod_UI_Manager.SwitchTo "OSM"
        End If
    ElseIf buttonName = "btn_ToPreviousMonth" Then
        Dim results As Collection
        Set results = mod_Services.ScanPreviousMonthSheets()
        
        If results.Count = 0 Then
            MsgBox MSG_PREVIOUS_MONTH_SHEETS_NOT_FOUND, vbExclamation, MSG_TITLE_ERROR
            Exit Sub
        End If
        
        mod_UI_Manager.SwitchTo "PreviousMonth"
    End If
End Sub

Public Sub HandleOSMPanelEvents(buttonName As String, parentFrame As MSForms.Frame)
    Select Case buttonName
        Case "btn_BackToLogin"
            mod_UI_Manager.SwitchTo "Login"
        Case "btn_CheckSheets"
            mod_UI_View.RenderOSMSheetStatus parentFrame, True
        Case "btn_NextStep"
            If mod_Services.HasAnyOpenCase() Then
                mod_UI_Manager.SwitchTo "Open"
            ElseIf mod_Services.HasAnyClosedCase() Then
                mod_UI_Manager.SwitchTo "Closed"
            Else
                MsgBox MSG_NO_SHEETS_TO_PROCESS, vbInformation, MSG_TITLE_INFO
            End If
    End Select
End Sub

Public Sub HandleOpenPanelEvents(buttonName As String, parentFrame As MSForms.Frame)
    Select Case buttonName
        Case "btn_BackToLogin"
            mod_UI_Manager.SwitchTo "Login"
        Case "updateOpen"
            HandleUpdateOpen
        Case "nextOpenButton"
            If mod_Services.HasAnyClosedCase() Then
                mod_UI_Manager.SwitchTo "Closed"
            Else
                MsgBox MSG_NO_MORE_CLOSED_CASES, vbInformation, MSG_TITLE_COMPLETED
                mod_UI_Manager.SwitchTo "Login"
            End If
    End Select
End Sub

Public Sub HandleClosedPanelEvents(buttonName As String, parentFrame As MSForms.Frame)
    Select Case buttonName
        Case "btn_BackToLogin"
            mod_UI_Manager.SwitchTo "Login"
        Case "updateClosed"
            HandleUpdateClosed
        Case "nextClosedButton"
            MsgBox MSG_ALL_OPERATIONS_COMPLETED, vbInformation, MSG_TITLE_COMPLETED
            mod_UI_Manager.SwitchTo "Login"
    End Select
End Sub

Public Sub HandlePreviousMonthPanelEvents(buttonName As String, parentFrame As MSForms.Frame)
    Select Case buttonName
        Case "btn_BackToLogin"
            mod_UI_Manager.SwitchTo "Login"
            
        Case "btn_CheckPreviousMonthSheets"
            mod_UI_View.RenderPreviousMonthSheets parentFrame, True
            
        Case "btn_NextStepPreviousMonth"
            mod_UI_Manager.SwitchTo "SelectPreviousMonth"
    End Select
End Sub

Public Sub HandleSelectPreviousMonthEvents(buttonName As String, parentFrame As MSForms.Frame)
    Dim txtYear As MSForms.TextBox, cmbMonth As MSForms.ComboBox
    
    If buttonName = "btn_BackToPreviousMonth" Then
        mod_UI_Manager.SwitchTo "PreviousMonth"
        
    ElseIf buttonName = "btn_NextFromSelectMonth" Then
        Set txtYear = parentFrame.Controls("txt_YearInput")
        Set cmbMonth = parentFrame.Controls("cmb_MonthSelect")
        
        If cmbMonth.ListIndex <> -1 And IsNumeric(txtYear.Text) And Len(Trim(txtYear.Text)) = 4 Then
            SelectedPreviousYear = CInt(txtYear.Text)
            SelectedPreviousMonth = cmbMonth.ListIndex + 1
            
            On Error GoTo CleanUp
            Application.ScreenUpdating = False
            Application.Calculation = xlCalculationManual
            
            mod_UI_Manager.SwitchTo "PrevMonthClosed"
            
CleanUp:
            Application.Calculation = xlCalculationAutomatic
            Application.ScreenUpdating = True
            On Error GoTo 0
        Else
            MsgBox MSG_INVALID_YEAR_INPUT, vbExclamation, MSG_TITLE_ERROR
        End If
    End If
End Sub

Public Sub HandlePrevMonthClosedEvents(buttonName As String, parentFrame As MSForms.Frame)
    Select Case buttonName
        Case "btn_BackToSelectMonth"
            mod_UI_Manager.SwitchTo "SelectPreviousMonth"
        Case "btn_UpdatePrevMonthClosed"
            If mod_Services.ExecutePreviousMonthWrite(parentFrame.Controls("lst_PrevMonthClosedStatus"), cellColIdx:=3, valColIdx:=2) Then
                SetUpdateButtonSuccess parentFrame.Controls("btn_UpdatePrevMonthClosed"), parentFrame.Controls("btn_NextFromPrevMonthClosed")
                MsgBox MSG_PREV_MONTH_SAVE_SUCCESS, vbInformation, MSG_TITLE_SUCCESS
            End If
        Case "btn_NextFromPrevMonthClosed"
            mod_UI_Manager.SwitchTo "PrevMonthOpen"
    End Select
End Sub

Public Sub HandlePrevMonthOpenEvents(buttonName As String, parentFrame As MSForms.Frame)
    Select Case buttonName
        Case "btn_BackToPrevClosed"
            mod_UI_Manager.SwitchTo "PrevMonthClosed"
        Case "btn_UpdatePrevMonthOpen"
            If mod_Services.ExecutePreviousMonthWrite(parentFrame.Controls("lst_PrevMonthOpenStatus"), cellColIdx:=4, valColIdx:=3) Then
                SetUpdateButtonSuccess parentFrame.Controls("btn_UpdatePrevMonthOpen"), parentFrame.Controls("btn_NextFromPrevMonthOpen")
                MsgBox MSG_PREV_MONTH_OPEN_SAVE_SUCCESS, vbInformation, MSG_TITLE_SUCCESS
            End If
        Case "btn_NextFromPrevMonthOpen"
            MsgBox MSG_ALL_OPERATIONS_COMPLETED, vbInformation, MSG_TITLE_COMPLETED
            mod_UI_Manager.SwitchTo "Login"
    End Select
End Sub

Private Sub HandleUpdateOpen()
    Dim ws As Worksheet, realFrame As MSForms.Frame
    Set ws = ThisWorkbook.Sheets(MAIN_SHEET_NAME)
    Set realFrame = ws.OLEObjects("pan_Open").Object
    
    If mod_Services.ExecuteDatabaseWrite(realFrame.Controls("lst_OpenStatus"), cellColIdx:=4, valColIdx:=3, sheetNameColIdx:=0) Then
        SetUpdateButtonSuccess realFrame.Controls("updateOpen"), realFrame.Controls("nextOpenButton")
        MsgBox MSG_OPEN_SAVE_SUCCESS, vbInformation, MSG_TITLE_SUCCESS
    End If
End Sub

Private Sub HandleUpdateClosed()
    Dim ws As Worksheet, realFrame As MSForms.Frame
    Set ws = ThisWorkbook.Sheets(MAIN_SHEET_NAME)
    Set realFrame = ws.OLEObjects("pan_Closed").Object
    
    If mod_Services.ExecuteDatabaseWrite(realFrame.Controls("lst_ClosedStatus"), cellColIdx:=5, valColIdx:=4, sheetNameColIdx:=0) Then
        SetUpdateButtonSuccess realFrame.Controls("updateClosed"), realFrame.Controls("nextClosedButton")
        MsgBox MSG_CLOSED_SAVE_SUCCESS, vbInformation, MSG_TITLE_SUCCESS
    End If
End Sub

Private Sub SetUpdateButtonSuccess(btnUpdate As MSForms.CommandButton, btnNext As MSForms.CommandButton)
    btnUpdate.BackColor = RGB(46, 204, 113)
    btnUpdate.Caption = BTN_CAPTION_UPDATED
    btnUpdate.Enabled = False
    btnNext.Enabled = True
    btnNext.BackColor = COLOR_GREEN
    btnNext.ForeColor = COLOR_WHITE
End Sub