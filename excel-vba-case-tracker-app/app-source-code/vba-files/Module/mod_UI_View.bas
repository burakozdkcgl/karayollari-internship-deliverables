Attribute VB_Name = "mod_UI_View"
Option Explicit

'*****************************************************************************
' Module: mod_UI_View
' Purpose: This module contains public functions for creating and rendering user interface panels and controls.
'*****************************************************************************

Public Sub CreateLoginPanel(ws As Worksheet)
    Dim oleFrame As OLEObject, realFrame As MSForms.Frame
    Set oleFrame = ws.OLEObjects.Add(ClassType:="Forms.Frame.1", Left:=PANEL_LEFT, Top:=PANEL_TOP, Width:=PANEL_WIDTH, Height:=PANEL_HEIGHT)
    oleFrame.Name = "pan_Login"
    Set realFrame = oleFrame.Object
    realFrame.Caption = ""
    realFrame.BackColor = COLOR_GRAY_BG
    realFrame.SpecialEffect = 0
    
    AddControlButton realFrame, "btn_ToOSM", BTN_CAPTION_OSM_ANALYSIS, 120, (PANEL_HEIGHT - 50) / 2, 210, 50, COLOR_GREEN
    AddControlButton realFrame, "btn_ToPreviousMonth", BTN_CAPTION_PREVIOUS_MONTH_UPDATE, 370, (PANEL_HEIGHT - 50) / 2, 210, 50, COLOR_WHITE, &H333333
    
    oleFrame.Visible = True
End Sub

Public Sub CreateOSMPanel(ws As Worksheet)
    Dim oleFrame As OLEObject, realFrame As MSForms.Frame, btnNext As MSForms.CommandButton
    Set oleFrame = ws.OLEObjects.Add(ClassType:="Forms.Frame.1", Left:=PANEL_LEFT, Top:=PANEL_TOP, Width:=PANEL_WIDTH, Height:=PANEL_HEIGHT)
    oleFrame.Name = "pan_OSM"
    Set realFrame = oleFrame.Object
    realFrame.Caption = ""
    realFrame.BackColor = &HF0F0F0
    realFrame.SpecialEffect = 0
    
    AddControlLabel realFrame, "lbl_Title", LBL_OSM_TITLE, 40, 20
    AddControlListBox realFrame, "lst_Status", 5, "60;110;100;100;250"
    
    AddControlButton realFrame, "btn_BackToLogin", BTN_CAPTION_BACK, 40, 260, 120, 35, COLOR_RED
    AddControlButton realFrame, "btn_CheckSheets", BTN_CAPTION_CHECK_SHEETS, 275, 260, 150, 35, COLOR_WHITE, &H333333
    Set btnNext = AddControlButton(realFrame, "btn_NextStep", BTN_CAPTION_NEXT, 540, 260, 120, 35, &HDCDCDC)
    btnNext.Enabled = False
    
    oleFrame.Visible = False
End Sub

Public Sub CreateOpenPanel(ws As Worksheet)
    Dim oleFrame As OLEObject, realFrame As MSForms.Frame, btnNext As MSForms.CommandButton
    Set oleFrame = ws.OLEObjects.Add(ClassType:="Forms.Frame.1", Left:=PANEL_LEFT, Top:=PANEL_TOP, Width:=PANEL_WIDTH, Height:=PANEL_HEIGHT)
    oleFrame.Name = "pan_Open"
    Set realFrame = oleFrame.Object
    realFrame.Caption = ""
    realFrame.BackColor = COLOR_GRAY_BG
    realFrame.SpecialEffect = 0
    
    AddControlLabel realFrame, "lbl_OpenTitle", LBL_OPEN_TITLE, 40, 20
    AddControlListBox realFrame, "lst_OpenStatus", 5, "200;90;100;80;70"
    
    AddControlButton realFrame, "btn_BackToLogin", BTN_CAPTION_BACK, 40, 260, 120, 35, COLOR_RED
    AddControlButton realFrame, "updateOpen", BTN_CAPTION_UPDATE, 380, 260, 120, 35, COLOR_WHITE, &H333333
    Set btnNext = AddControlButton(realFrame, "nextOpenButton", BTN_CAPTION_NEXT, 540, 260, 120, 35, RGB(220, 220, 220))
    btnNext.Enabled = False
    
    oleFrame.Visible = False
End Sub

Public Sub CreateClosedPanel(ws As Worksheet)
    Dim oleFrame As OLEObject, realFrame As MSForms.Frame, btnNext As MSForms.CommandButton
    Set oleFrame = ws.OLEObjects.Add(ClassType:="Forms.Frame.1", Left:=PANEL_LEFT, Top:=PANEL_TOP, Width:=PANEL_WIDTH, Height:=PANEL_HEIGHT)
    oleFrame.Name = "pan_Closed"
    Set realFrame = oleFrame.Object
    realFrame.Caption = ""
    realFrame.BackColor = COLOR_GRAY_BG
    realFrame.SpecialEffect = 0
    
    AddControlLabel realFrame, "lbl_ClosedTitle", LBL_CLOSED_TITLE, 40, 20
    AddControlListBox realFrame, "lst_ClosedStatus", 6, "170;75;85;130;50;80"
    
    AddControlButton realFrame, "btn_BackToLogin", BTN_CAPTION_BACK, 40, 260, 120, 35, COLOR_RED
    AddControlButton realFrame, "updateClosed", BTN_CAPTION_UPDATE, 380, 260, 120, 35, COLOR_WHITE, &H333333
    Set btnNext = AddControlButton(realFrame, "nextClosedButton", BTN_CAPTION_NEXT, 540, 260, 120, 35, RGB(220, 220, 220))
    btnNext.Enabled = False
    
    oleFrame.Visible = False
End Sub

Public Sub CreatePreviousMonthPanel(ws As Worksheet)
    Dim oleFrame As OLEObject, realFrame As MSForms.Frame, btnNext As MSForms.CommandButton
    Set oleFrame = ws.OLEObjects.Add(ClassType:="Forms.Frame.1", Left:=PANEL_LEFT, Top:=PANEL_TOP, Width:=PANEL_WIDTH, Height:=PANEL_HEIGHT)
    oleFrame.Name = "pan_PreviousMonth"
    Set realFrame = oleFrame.Object
    realFrame.Caption = ""
    realFrame.BackColor = COLOR_GRAY_BG
    realFrame.SpecialEffect = 0
    
    AddControlLabel realFrame, "lbl_PreviousMonthTitle", LBL_PREVIOUS_MONTH_TITLE, 40, 20
    AddControlListBox realFrame, "lst_PreviousMonthStatus", 3, "70;140;370"
    
    AddControlButton realFrame, "btn_BackToLogin", BTN_CAPTION_BACK, 40, 260, 120, 35, COLOR_RED
    AddControlButton realFrame, "btn_CheckPreviousMonthSheets", BTN_CAPTION_CHECK_SHEETS, 275, 260, 150, 35, COLOR_WHITE, &H333333
    
    Set btnNext = AddControlButton(realFrame, "btn_NextStepPreviousMonth", BTN_CAPTION_NEXT, 540, 260, 120, 35, &HDCDCDC)
    btnNext.Enabled = False
    
    oleFrame.Visible = False
End Sub

Public Sub CreateSelectPreviousMonthPanel(ws As Worksheet)
    Dim oleFrame As OLEObject, realFrame As MSForms.Frame
    Dim cmbMonth As MSForms.ComboBox, txtYear As MSForms.TextBox
    Dim lblMonthPrompt As MSForms.Label, lblYearPrompt As MSForms.Label
    
    Dim totalGroupWidth As Single, totalGroupHeight As Single
    Dim startX As Single, startY As Single
    Dim lblWidth As Single, inputWidth As Single, rowHeight As Single, gapY As Single
    
    Set oleFrame = ws.OLEObjects.Add(ClassType:="Forms.Frame.1", Left:=PANEL_LEFT, Top:=PANEL_TOP, Width:=PANEL_WIDTH, Height:=PANEL_HEIGHT)
    oleFrame.Name = "pan_SelectPreviousMonth"
    Set realFrame = oleFrame.Object
    realFrame.Caption = ""
    realFrame.BackColor = COLOR_GRAY_BG
    realFrame.SpecialEffect = 0
    
    AddControlLabel realFrame, "lbl_SelectTitle", LBL_SELECT_PREV_MONTH_TITLE, 40, 20
    
    lblWidth = 60
    inputWidth = 160
    rowHeight = 28
    gapY = 20
    
    totalGroupWidth = lblWidth + inputWidth
    totalGroupHeight = (rowHeight * 2) + gapY
    
    startX = (PANEL_WIDTH - totalGroupWidth) / 2
    startY = (PANEL_HEIGHT - totalGroupHeight - 60) / 2
    
    Set lblMonthPrompt = realFrame.Controls.Add("Forms.Label.1", "lbl_MonthPrompt", True)
    With lblMonthPrompt
        .Caption = LBL_SELECT_MONTH_PROMPT
        .Font.Name = "Segoe UI"
        .Font.Bold = True
        .Font.Size = 10
        .Left = startX
        .Top = startY + 4
        .Width = lblWidth
        .Height = rowHeight
    End With
    
    Set cmbMonth = realFrame.Controls.Add("Forms.ComboBox.1", "cmb_MonthSelect", True)
    With cmbMonth
        .Font.Name = "Segoe UI"
        .Font.Bold = True
        .Font.Size = 10
        .Left = startX + lblWidth
        .Top = startY
        .Width = inputWidth
        .Height = rowHeight
        .Style = fmStyleDropDownList
    End With
    
    Set lblYearPrompt = realFrame.Controls.Add("Forms.Label.1", "lbl_YearPrompt", True)
    With lblYearPrompt
        .Caption = LBL_SELECT_YEAR_PROMPT
        .Font.Name = "Segoe UI"
        .Font.Bold = True
        .Font.Size = 10
        .Left = startX
        .Top = startY + rowHeight + gapY + 4
        .Width = lblWidth
        .Height = rowHeight
    End With
    
    Set txtYear = realFrame.Controls.Add("Forms.TextBox.1", "txt_YearInput", True)
    With txtYear
        .Font.Name = "Segoe UI"
        .Font.Bold = True
        .Font.Size = 10
        .Left = startX + lblWidth
        .Top = startY + rowHeight + gapY
        .Width = inputWidth
        .Height = rowHeight
        .BorderStyle = 1
        .SpecialEffect = 0
        .MaxLength = 4
    End With
    
    AddControlButton realFrame, "btn_BackToPreviousMonth", BTN_CAPTION_BACK, 40, 260, 120, 35, COLOR_RED
    AddControlButton realFrame, "btn_NextFromSelectMonth", BTN_CAPTION_NEXT, 540, 260, 120, 35, COLOR_GREEN
    
    oleFrame.Visible = False
End Sub

Public Sub CreatePrevMonthClosedPanel(ws As Worksheet)
    Dim oleFrame As OLEObject, realFrame As MSForms.Frame, btnNext As MSForms.CommandButton
    Set oleFrame = ws.OLEObjects.Add(ClassType:="Forms.Frame.1", Left:=PANEL_LEFT, Top:=PANEL_TOP, Width:=PANEL_WIDTH, Height:=PANEL_HEIGHT)
    oleFrame.Name = "pan_PrevMonthClosed"
    Set realFrame = oleFrame.Object
    realFrame.Caption = ""
    realFrame.BackColor = COLOR_GRAY_BG
    realFrame.SpecialEffect = 0
    
    AddControlLabel realFrame, "lbl_PrevMonthClosedTitle", LBL_PREV_MONTH_CLOSED_TITLE, 40, 20
    AddControlListBox realFrame, "lst_PrevMonthClosedStatus", 4, "180;100;120;120"
    
    AddControlButton realFrame, "btn_BackToSelectMonth", BTN_CAPTION_BACK, 40, 260, 120, 35, COLOR_RED
    AddControlButton realFrame, "btn_UpdatePrevMonthClosed", BTN_CAPTION_UPDATE, 380, 260, 120, 35, COLOR_WHITE, &H333333
    Set btnNext = AddControlButton(realFrame, "btn_NextFromPrevMonthClosed", BTN_CAPTION_NEXT, 540, 260, 120, 35, RGB(220, 220, 220))
    btnNext.Enabled = False
    
    oleFrame.Visible = False
End Sub

Public Sub CreatePrevMonthOpenPanel(ws As Worksheet)
    Dim oleFrame As OLEObject, realFrame As MSForms.Frame, btnNext As MSForms.CommandButton
    Set oleFrame = ws.OLEObjects.Add(ClassType:="Forms.Frame.1", Left:=PANEL_LEFT, Top:=PANEL_TOP, Width:=PANEL_WIDTH, Height:=PANEL_HEIGHT)
    oleFrame.Name = "pan_PrevMonthOpen"
    Set realFrame = oleFrame.Object
    realFrame.Caption = ""
    realFrame.BackColor = COLOR_GRAY_BG
    realFrame.SpecialEffect = 0
    
    AddControlLabel realFrame, "lbl_PrevMonthOpenTitle", LBL_PREV_MONTH_OPEN_TITLE, 40, 20
    AddControlListBox realFrame, "lst_PrevMonthOpenStatus", 5, "140;80;150;90;120"
    
    AddControlButton realFrame, "btn_BackToPrevClosed", BTN_CAPTION_BACK, 40, 260, 120, 35, COLOR_RED
    AddControlButton realFrame, "btn_UpdatePrevMonthOpen", BTN_CAPTION_UPDATE, 380, 260, 120, 35, COLOR_WHITE, &H333333
    Set btnNext = AddControlButton(realFrame, "btn_NextFromPrevMonthOpen", BTN_CAPTION_NEXT, 540, 260, 120, 35, RGB(220, 220, 220))
    btnNext.Enabled = False
    
    oleFrame.Visible = False
End Sub

' =========================================================================
' UI RENDERER METHODS
' =========================================================================

Public Sub RenderOSMSheetStatus(realFrame As MSForms.Frame, Optional isManualCheck As Boolean = False)
    Dim lstStatus As MSForms.ListBox, btnNext As MSForms.CommandButton
    Dim scanResults As Collection, sortedResults As Collection, item As Variant, rIdx As Long
    
    Set lstStatus = realFrame.Controls("lst_Status")
    Set btnNext = realFrame.Controls("btn_NextStep")
    
    Dim isFirstOpen As Boolean: isFirstOpen = (lstStatus.ListCount = 0)
    lstStatus.Clear
    
    lstStatus.AddItem
    lstStatus.List(0, 0) = LST_HDR_STATUS
    lstStatus.List(0, 1) = LST_HDR_REPORT_TYPE
    lstStatus.List(0, 2) = LST_HDR_REPORT_DATE
    lstStatus.List(0, 3) = LST_HDR_AFFECTED_DATE
    lstStatus.List(0, 4) = LST_HDR_SHEET_NAME
    
    Set scanResults = mod_Services.ScanWorkbookSheets()
    If scanResults.Count = 0 Then
        lstStatus.AddItem
        lstStatus.List(1, 0) = STATUS_ERROR
        lstStatus.List(1, 1) = STATUS_NOT_FOUND
        lstStatus.List(1, 2) = STATUS_NOT_FOUND
        lstStatus.List(1, 3) = STATUS_NOT_FOUND
        lstStatus.List(1, 4) = LST_VAL_MISSING_SHEET
        btnNext.Enabled = False
        btnNext.BackColor = &HDCDCDC
    Else
        Set sortedResults = mod_Services.SortCollectionByDateIndex(scanResults, 3)
        rIdx = 1
        For Each item In sortedResults
            lstStatus.AddItem
            lstStatus.List(rIdx, 0) = IIf(item(2) = STATUS_NOT_FOUND Or Trim(item(2)) = "", STATUS_ERROR, STATUS_OK)
            lstStatus.List(rIdx, 1) = mod_Services.GetTypeStringByID(CStr(item(1)))
            lstStatus.List(rIdx, 2) = item(2)
            lstStatus.List(rIdx, 3) = item(3)
            lstStatus.List(rIdx, 4) = item(4)
            rIdx = rIdx + 1
        Next item
        
        If Not isFirstOpen Or isManualCheck Then
            btnNext.Enabled = True
            btnNext.BackColor = COLOR_GREEN
            Application.OnTime Now + TimeValue("00:00:05"), "LockPanel"
        Else
            btnNext.Enabled = False
            btnNext.BackColor = &HDCDCDC
        End If
    End If
    
    Do While lstStatus.ListCount < 12: lstStatus.AddItem: Loop
End Sub

Public Sub RenderOpenSheets(realFrame As MSForms.Frame)
    Dim lstOpen As MSForms.ListBox, calculatedData As Collection, sortedData As Collection
    Dim item As Variant, rIdx As Long
    Dim rawCell As String, dbName As String, displayCell As String
    
    ResetPanelButtons realFrame, "nextOpenButton", "updateOpen"
    Set lstOpen = realFrame.Controls("lst_OpenStatus")
    lstOpen.Clear
    
    lstOpen.AddItem
    lstOpen.List(0, 0) = LST_HDR_SHEET_NAME
    lstOpen.List(0, 1) = LST_HDR_REPORT_DATE
    lstOpen.List(0, 2) = LST_HDR_AFFECTED_DATE
    lstOpen.List(0, 3) = LST_HDR_CASE_COUNT
    lstOpen.List(0, 4) = LST_HDR_AFFECTED_CELL
    
    Set calculatedData = mod_Services.GetOpenCasesCalculated()
    If calculatedData.Count > 0 Then
        Set sortedData = mod_Services.SortCollectionByDateIndex(calculatedData, 2)
        rIdx = 1
        For Each item In sortedData
            lstOpen.AddItem
            lstOpen.List(rIdx, 0) = item(0)
            lstOpen.List(rIdx, 1) = item(1)
            lstOpen.List(rIdx, 2) = item(2)
            lstOpen.List(rIdx, 3) = item(3)
            
            rawCell = Trim("" & item(4))
            dbName = mod_Services.GetTargetDbSheetName(CStr(item(0)))
            
            If rawCell <> STATUS_NOT_FOUND And rawCell <> STATUS_NO_DB_SHEET And rawCell <> "" Then
                displayCell = rawCell & " (" & dbName & ")"
            Else
                displayCell = rawCell
            End If
            
            lstOpen.List(rIdx, 4) = displayCell
            rIdx = rIdx + 1
        Next item
    End If
    
    Do While lstOpen.ListCount < 12: lstOpen.AddItem: Loop
End Sub

Public Sub RenderClosedSheets(realFrame As MSForms.Frame)
    Dim lstClosed As MSForms.ListBox, calculatedData As Collection, item As Variant, rIdx As Long
    Dim rawCell As String, dbName As String, displayCell As String
    
    ResetPanelButtons realFrame, "nextClosedButton", "updateClosed"
    Set lstClosed = realFrame.Controls("lst_ClosedStatus")
    lstClosed.Clear
    
    lstClosed.AddItem
    lstClosed.List(0, 0) = LST_HDR_SHEET_NAME
    lstClosed.List(0, 1) = LST_HDR_REPORT_DATE
    lstClosed.List(0, 2) = LST_HDR_AFFECTED_DATE
    lstClosed.List(0, 3) = LST_HDR_PERSON_INFO
    lstClosed.List(0, 4) = LST_HDR_CASE
    lstClosed.List(0, 5) = LST_HDR_CELL
    
    Set calculatedData = mod_Services.GetClosedCasesCalculated()
    rIdx = 1
    If calculatedData.Count > 0 Then
        For Each item In calculatedData
            lstClosed.AddItem
            lstClosed.List(rIdx, 0) = item(0)
            lstClosed.List(rIdx, 1) = item(1)
            lstClosed.List(rIdx, 2) = item(2)
            lstClosed.List(rIdx, 3) = item(3)
            lstClosed.List(rIdx, 4) = item(4)
            
            rawCell = Trim("" & item(5))
            dbName = mod_Services.GetTargetDbSheetName(CStr(item(0)))
            
            If rawCell <> STATUS_NOT_FOUND And rawCell <> STATUS_NO_DB_SHEET And rawCell <> "" Then
                displayCell = rawCell & " (" & dbName & ")"
            Else
                displayCell = rawCell
            End If
            
            lstClosed.List(rIdx, 5) = displayCell
            rIdx = rIdx + 1
        Next item
    End If
    
    Do While lstClosed.ListCount < 12: lstClosed.AddItem: Loop
End Sub

Public Sub RenderPreviousMonthSheets(realFrame As MSForms.Frame, Optional isManualCheck As Boolean = False)
    Dim lst As MSForms.ListBox, results As Collection, item As Variant, rIdx As Long
    Dim btnNext As MSForms.CommandButton
    
    Set lst = realFrame.Controls("lst_PreviousMonthStatus")
    Set btnNext = realFrame.Controls("btn_NextStepPreviousMonth")
    
    Dim isFirstOpen As Boolean: isFirstOpen = (lst.ListCount = 0)
    lst.Clear
    
    lst.ColumnCount = 3
    lst.ColumnWidths = "70;140;370"
    
    lst.AddItem
    lst.List(0, 0) = LST_HDR_STATUS
    lst.List(0, 1) = LST_HDR_REPORT_TYPE
    lst.List(0, 2) = LST_HDR_SHEET_NAME
    
    Set results = mod_Services.ScanPreviousMonthSheets()
    
    If results.Count = 0 Then
        lst.AddItem
        lst.List(1, 0) = STATUS_ERROR
        lst.List(1, 1) = STATUS_NOT_FOUND
        lst.List(1, 2) = LST_VAL_MISSING_SHEET
        
        If Not btnNext Is Nothing Then
            btnNext.Enabled = False
            btnNext.BackColor = &HDCDCDC
        End If
    Else
        rIdx = 1
        For Each item In results
            lst.AddItem
            lst.List(rIdx, 0) = item(0)
            lst.List(rIdx, 1) = item(1)
            lst.List(rIdx, 2) = item(2)
            rIdx = rIdx + 1
        Next item
        
        If Not btnNext Is Nothing Then
            If Not isFirstOpen Or isManualCheck Then
                btnNext.Enabled = True
                btnNext.BackColor = COLOR_GREEN
                Application.OnTime Now + TimeValue("00:00:05"), "LockPreviousMonthPanel"
            Else
                btnNext.Enabled = False
                btnNext.BackColor = &HDCDCDC
            End If
        End If
    End If
    
    Do While lst.ListCount < 12: lst.AddItem: Loop
End Sub

Public Sub RenderSelectPreviousMonthPanel(realFrame As MSForms.Frame)
    Dim cmbMonth As MSForms.ComboBox, txtYear As MSForms.TextBox
    Dim months As Variant, i As Long
    
    Set cmbMonth = realFrame.Controls("cmb_MonthSelect")
    Set txtYear = realFrame.Controls("txt_YearInput")
    
    If SelectedPreviousMonth = 0 Then SelectedPreviousMonth = Month(Date)
    If SelectedPreviousYear = 0 Then SelectedPreviousYear = Year(Date)
    
    cmbMonth.Clear
    
    ' Ayları config üzerindeki sabitten dinamik olarak çeker
    months = Split(MONTH_NAMES_LIST, ",")
    For i = LBound(months) To UBound(months)
        cmbMonth.AddItem Trim(months(i))
    Next i
    
    cmbMonth.ListIndex = SelectedPreviousMonth - 1
    
    txtYear.Text = CStr(SelectedPreviousYear)
End Sub

Public Sub RenderPrevMonthClosedPanel(realFrame As MSForms.Frame)
    Dim lst As MSForms.ListBox, calculatedData As Collection, item As Variant, rIdx As Long
    
    ResetPanelButtons realFrame, "btn_NextFromPrevMonthClosed", "btn_UpdatePrevMonthClosed"
    
    Set lst = realFrame.Controls("lst_PrevMonthClosedStatus")
    lst.Clear
    
    lst.Width = 620
    lst.Height = 180
    lst.ColumnCount = 4
    lst.ColumnWidths = "180;100;120;150"
    
    lst.AddItem
    lst.List(0, 0) = LST_HDR_SHEET_NAME
    lst.List(0, 1) = LST_HDR_TYPE
    lst.List(0, 2) = LST_HDR_PREV_CLOSED_COUNT
    lst.List(0, 3) = LST_HDR_AFFECTED_CELL
    
    Set calculatedData = mod_Services.GetPreviousMonthClosedCalculated(SelectedPreviousMonth, SelectedPreviousYear)
    
    rIdx = 1
    If calculatedData.Count > 0 Then
        For Each item In calculatedData
            lst.AddItem
            lst.List(rIdx, 0) = item(0)
            lst.List(rIdx, 1) = item(1)
            lst.List(rIdx, 2) = item(2)
            lst.List(rIdx, 3) = item(3) & " (" & PREVIOUS_MONTH_SHEET_NAME & ")"
            rIdx = rIdx + 1
        Next item
    End If
    
    Do While lst.ListCount < 12: lst.AddItem: Loop
End Sub

Public Sub RenderPrevMonthOpenPanel(realFrame As MSForms.Frame)
    Dim lst As MSForms.ListBox, calculatedData As Collection, item As Variant, rIdx As Long
    
    ResetPanelButtons realFrame, "btn_NextFromPrevMonthOpen", "btn_UpdatePrevMonthOpen"
    
    Set lst = realFrame.Controls("lst_PrevMonthOpenStatus")
    lst.Clear
    
    lst.Width = 620
    lst.Height = 180
    lst.ColumnCount = 5
    lst.ColumnWidths = "140;80;150;90;120"
    
    lst.AddItem
    lst.List(0, 0) = LST_HDR_SHEET_NAME
    lst.List(0, 1) = LST_HDR_TYPE
    lst.List(0, 2) = LST_HDR_PERSON_INFO
    lst.List(0, 3) = LST_HDR_PREV_OPEN_COUNT
    lst.List(0, 4) = LST_HDR_AFFECTED_CELL
    
    Set calculatedData = mod_Services.GetPreviousMonthOpenPersonCalculated(SelectedPreviousMonth, SelectedPreviousYear)
    
    rIdx = 1
    If calculatedData.Count > 0 Then
        For Each item In calculatedData
            lst.AddItem
            lst.List(rIdx, 0) = item(0)
            lst.List(rIdx, 1) = item(1)
            lst.List(rIdx, 2) = item(2)
            lst.List(rIdx, 3) = item(3)
            lst.List(rIdx, 4) = item(4) & " (" & PREVIOUS_MONTH_SHEET_NAME & ")"
            rIdx = rIdx + 1
        Next item
    End If
    
    Do While lst.ListCount < 12: lst.AddItem: Loop
End Sub

' --- Private Builder & Renderer Helpers ---

Private Function AddControlButton(pFrame As MSForms.Frame, cName As String, cap As String, L As Double, T As Double, W As Double, H As Double, bg As Long, Optional fg As Long = &HFFFFFF) As MSForms.CommandButton
    Dim btn As MSForms.CommandButton, btnListener As clsDynamicButton
    Set btn = pFrame.Controls.Add("Forms.CommandButton.1", cName, True)
    With btn
        .Caption = cap
        .Font.Name = "Segoe UI"
        .Font.Size = 10
        .Font.Bold = True
        .Width = W
        .Height = H
        .Left = L
        .Top = T
        .BackColor = bg
        .ForeColor = fg
    End With
    Set btnListener = New clsDynamicButton
    Set btnListener.DynamicButton = btn
    ButtonCollection.Add btnListener
    Set AddControlButton = btn
End Function

Private Sub AddControlLabel(pFrame As MSForms.Frame, cName As String, cap As String, L As Double, T As Double)
    Dim lbl As MSForms.Label
    Set lbl = pFrame.Controls.Add("Forms.Label.1", cName, True)
    With lbl
        .Caption = cap
        .Font.Name = "Segoe UI"
        .Font.Bold = True
        .Font.Size = 14
        .ForeColor = &H333333
        .Left = L
        .Top = T
        .Width = 400
        .Height = 30
        .BackStyle = 0
    End With
End Sub

Private Sub AddControlListBox(pFrame As MSForms.Frame, cName As String, colCnt As Integer, colWidths As String)
    Dim lst As MSForms.ListBox
    Set lst = pFrame.Controls.Add("Forms.ListBox.1", cName, True)
    With lst
        .Font.Name = "Segoe UI"
        .Font.Size = 9
        .Left = 40
        .Top = 60
        .Width = 620
        .Height = 180
        .ColumnCount = colCnt
        .ColumnWidths = colWidths
        .BackColor = &HFFFFFF
        .BorderStyle = 1
    End With
End Sub

Private Sub ResetPanelButtons(pFrame As MSForms.Frame, nextBtnName As String, updateBtnName As String)
    On Error Resume Next
    With pFrame.Controls(nextBtnName)
        .Enabled = False
        .BackColor = RGB(220, 220, 220)
        .ForeColor = COLOR_WHITE
    End With
    With pFrame.Controls(updateBtnName)
        .Enabled = True
        .Caption = BTN_CAPTION_UPDATE
        .BackColor = COLOR_WHITE
        .ForeColor = &H333333
    End With
    With pFrame.Controls("btn_BackToLogin")
        .Caption = BTN_CAPTION_BACK
        .BackColor = COLOR_RED
        .ForeColor = COLOR_WHITE
    End With
    On Error GoTo 0
End Sub