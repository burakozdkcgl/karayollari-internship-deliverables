Attribute VB_Name = "mod_UI_Manager"
Option Explicit

'*****************************************************************************
' Module: mod_UI_Manager
' Purpose: This module contains public functions for managing the user interface, including panel switching and initialization.
'*****************************************************************************

Public Sub BuildNavigationSystem()
    Dim ws As Worksheet
    
    If Not Application.ActiveProtectedViewWindow Is Nothing Then Exit Sub
    
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(MAIN_SHEET_NAME)
    On Error GoTo 0
    
    If ws Is Nothing Then Exit Sub
    
    If Not ButtonCollection Is Nothing Then
        Dim btnItem As Object
        For Each btnItem In ButtonCollection
            Set btnItem.DynamicButton = Nothing
        Next btnItem
        Set ButtonCollection = Nothing
    End If
    Set ButtonCollection = New Collection
    
    On Error Resume Next
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    
    Dim obj As OLEObject, anchorExists As Boolean
    For Each obj In ws.OLEObjects
        If obj.Name = "bx_Anchor" Then
            obj.Visible = False
            anchorExists = True
        Else
            obj.Delete
        End If
    Next obj
    
    If Not anchorExists Then
        Set obj = ws.OLEObjects.Add(ClassType:="Forms.TextBox.1", Left:=0, Top:=0, Width:=1, Height:=1)
        obj.Name = "bx_Anchor"
        obj.Visible = False
    End If
    
    mod_UI_View.CreateLoginPanel ws
    mod_UI_View.CreateOSMPanel ws
    mod_UI_View.CreateOpenPanel ws
    mod_UI_View.CreateClosedPanel ws
    mod_UI_View.CreatePreviousMonthPanel ws
    mod_UI_View.CreateSelectPreviousMonthPanel ws
    mod_UI_View.CreatePrevMonthClosedPanel ws
    mod_UI_View.CreatePrevMonthOpenPanel ws
    
    ' Güvenli Activate İşlemi
    If ActiveSheet.Name <> ws.Name Then
        ws.Select
    End If
    
CleanUp:
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    On Error GoTo 0
End Sub

Public Sub SwitchTo(panelName As String)
    Dim ws As Worksheet, oleObj As OLEObject, realFrame As MSForms.Frame
    Dim lst As MSForms.ListBox
    
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(MAIN_SHEET_NAME)

    If ws.OLEObjects("pan_" & panelName) Is Nothing Then Exit Sub
    
    Application.ScreenUpdating = False
    
    ws.OLEObjects("pan_Login").Visible = False
    ws.OLEObjects("pan_OSM").Visible = False
    ws.OLEObjects("pan_Open").Visible = False
    ws.OLEObjects("pan_Closed").Visible = False
    ws.OLEObjects("pan_PreviousMonth").Visible = False
    ws.OLEObjects("pan_SelectPreviousMonth").Visible = False
    ws.OLEObjects("pan_PrevMonthClosed").Visible = False
    ws.OLEObjects("pan_PrevMonthOpen").Visible = False

    Set oleObj = ws.OLEObjects("pan_" & panelName)
    oleObj.Left = PANEL_LEFT
    oleObj.Top = PANEL_TOP
    oleObj.Width = PANEL_WIDTH
    oleObj.Height = PANEL_HEIGHT
    oleObj.Visible = True

    Set realFrame = oleObj.Object
    
    Select Case panelName
        Case "OSM"
            Set lst = realFrame.Controls("lst_Status")
            If Not lst Is Nothing Then
                lst.Width = 620
                lst.Height = 180
                lst.ColumnCount = 5
                lst.ColumnWidths = "60;110;100;100;250"
            End If
            mod_UI_View.RenderOSMSheetStatus realFrame, False
            
        Case "Open"
            Set lst = realFrame.Controls("lst_OpenStatus")
            If Not lst Is Nothing Then
                lst.Width = 620
                lst.Height = 180
                lst.ColumnCount = 5
                lst.ColumnWidths = "200;90;100;80;70"
            End If
            mod_UI_View.RenderOpenSheets realFrame
            
        Case "Closed"
            Set lst = realFrame.Controls("lst_ClosedStatus")
            If Not lst Is Nothing Then
                lst.Width = 620
                lst.Height = 180
                lst.ColumnCount = 6
                lst.ColumnWidths = "170;75;85;130;50;80"
            End If
            mod_UI_View.RenderClosedSheets realFrame

        Case "PreviousMonth"
            Set lst = realFrame.Controls("lst_PreviousMonthStatus")
            If Not lst Is Nothing Then
                lst.Width = 620
                lst.Height = 180
                lst.ColumnCount = 3
                lst.ColumnWidths = "70;140;370"
            End If
            mod_UI_View.RenderPreviousMonthSheets realFrame, False
            
        Case "SelectPreviousMonth"
            mod_UI_View.RenderSelectPreviousMonthPanel realFrame
        
        Case "PrevMonthClosed"
            Set lst = realFrame.Controls("lst_PrevMonthClosedStatus")
            If Not lst Is Nothing Then
                lst.Width = 620
                lst.Height = 180
                lst.ColumnCount = 4
                lst.ColumnWidths = "180;100;120;150"
            End If
            mod_UI_View.RenderPrevMonthClosedPanel realFrame
        
        Case "PrevMonthOpen"
            Set lst = realFrame.Controls("lst_PrevMonthOpenStatus")
            If Not lst Is Nothing Then
                lst.Width = 620
                lst.Height = 180
                lst.ColumnCount = 5
                lst.ColumnWidths = "140;80;150;90;120"
            End If
            mod_UI_View.RenderPrevMonthOpenPanel realFrame

    End Select
    
    On Error GoTo 0
    Application.ScreenUpdating = True
End Sub

Public Sub LockPanel()
    Dim ws As Worksheet, realFrame As MSForms.Frame, btnNext As MSForms.CommandButton
    Set ws = ThisWorkbook.Sheets(MAIN_SHEET_NAME)
    On Error Resume Next
    Set realFrame = ws.OLEObjects("pan_OSM").Object
    Set btnNext = realFrame.Controls("btn_NextStep")
    If Not btnNext Is Nothing Then
        btnNext.Enabled = False
        btnNext.BackColor = &HDCDCDC
        btnNext.Caption = BTN_CAPTION_NEXT
    End If
    On Error GoTo 0
End Sub

Public Sub LockPreviousMonthPanel()
    Dim ws As Worksheet, realFrame As MSForms.Frame, btnNext As MSForms.CommandButton
    Set ws = ThisWorkbook.Sheets(MAIN_SHEET_NAME)
    On Error Resume Next
    Set realFrame = ws.OLEObjects("pan_PreviousMonth").Object
    Set btnNext = realFrame.Controls("btn_NextStepPreviousMonth")
    If Not btnNext Is Nothing Then
        btnNext.Enabled = False
        btnNext.BackColor = &HDCDCDC
        btnNext.Caption = BTN_CAPTION_NEXT
    End If
    On Error GoTo 0
End Sub

Public Sub InitializeAppUI()
    If Not Application.ActiveProtectedViewWindow Is Nothing Then Exit Sub
    
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(MAIN_SHEET_NAME)
    On Error GoTo 0
    
    If ws Is Nothing Then Exit Sub
    
    BuildNavigationSystem
    SwitchTo "Login"
    
    With ws
        .ScrollArea = "A1"
        .EnableSelection = xlNoSelection
    End With
End Sub