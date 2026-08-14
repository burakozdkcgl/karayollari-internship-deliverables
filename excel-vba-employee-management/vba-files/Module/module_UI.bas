Attribute VB_Name = "module_UI"
' =========================================================================
' MODULE: module_UI
' Description: Controls UI creation, panel switching, and image rendering.
' =========================================================================

Public ButtonCollection As New Collection

' Clean up dynamic objects from worksheet
Public Sub ClearCanvas()
    Dim ws As Worksheet, obj As OLEObject, anchorExists As Boolean
    Set ws = ThisWorkbook.Sheets("PORTAL")
    Set ButtonCollection = New Collection
    
    On Error Resume Next
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
    On Error GoTo 0
End Sub

' Rebuild all UI elements and setup navigation
Public Sub BuildNavigationSystem()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("PORTAL")
    
    Call ClearCanvas
    Call CreateLoginPanel(ws)
    Call CreateAssetsPanel(ws)
    Call CreatePersonnelPanel(ws)
    Call CreateVerifiedPanel(ws)
    ws.Activate
End Sub

' Switch visible panel dynamically
Public Sub SwitchTo(ByVal PanelName As String)
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("PORTAL")
    
    Application.ScreenUpdating = False
    On Error Resume Next
    
    ws.OLEObjects("pan_Login").Visible = False
    ws.OLEObjects("pan_Assets").Visible = False
    ws.OLEObjects("pan_Personnel").Visible = False
    ws.OLEObjects("pan_Verified").Visible = False
    
    ws.OLEObjects("pan_" & PanelName).Visible = True
    
    If PanelName = "Verified" Then Call RefreshVerifiedPanel
    
    On Error GoTo 0
    Application.ScreenUpdating = True
End Sub

' Build central Login panel
Public Sub CreateLoginPanel(ws As Worksheet)
    Dim oleFrame As OLEObject, realFrame As MSForms.Frame
    Dim btnAsset As MSForms.CommandButton, btnPersonnel As MSForms.CommandButton
    Dim imgLogo As MSForms.Image
    
    Set oleFrame = ws.OLEObjects.Add("Forms.Frame.1", Left:=PANEL_LEFT, Top:=PANEL_TOP, Width:=PANEL_WIDTH, Height:=PANEL_HEIGHT)
    oleFrame.Name = "pan_Login"
    
    Set realFrame = oleFrame.Object
    realFrame.Caption = "": realFrame.BackColor = COLOR_GRAY_BG: realFrame.SpecialEffect = 0
    
    Set imgLogo = realFrame.Controls.Add("Forms.Image.1", "img_Logo", True)
    With imgLogo
        .Width = 180: .Height = 180: .Left = (PANEL_WIDTH - 180) / 2: .Top = (PANEL_HEIGHT - 180) / 2
        .BorderStyle = 0: .BackColor = COLOR_GRAY_BG
    End With
    Call LoadEmbeddedShapeToControl(ws, "logo", imgLogo, RGB(240, 240, 240))
    
    Set btnAsset = realFrame.Controls.Add("Forms.CommandButton.1", "btn_Asset", True)
    With btnAsset
        .Caption = UI_BTN_ASSET_CAPTION: .Font.Name = "Segoe UI": .Font.Bold = True
        .Width = 100: .Height = 100: .Left = imgLogo.Left - 35 - 100: .Top = imgLogo.Top + 40
        .WordWrap = True: .BackColor = COLOR_GREEN: .ForeColor = COLOR_WHITE
    End With
    
    Set btnPersonnel = realFrame.Controls.Add("Forms.CommandButton.1", "btn_Personnel", True)
    With btnPersonnel
        .Caption = UI_BTN_PERSONNEL_CAPTION: .Font.Name = "Segoe UI": .Font.Bold = True
        .Width = 100: .Height = 100: .Left = imgLogo.Left + 180 + 35: .Top = imgLogo.Top + 40
        .WordWrap = True: .BackColor = COLOR_GREEN: .ForeColor = COLOR_WHITE
    End With
    
    Call AddEngineerToFrame(ws, realFrame)
    Call RegisterButton(btnAsset): Call RegisterButton(btnPersonnel)
    oleFrame.Visible = True
End Sub

' Build Assets panel
Public Sub CreateAssetsPanel(ws As Worksheet)
    Dim oleFrame As OLEObject, realFrame As MSForms.Frame
    Dim btnBack As MSForms.CommandButton, lblTitle As MSForms.Label
    
    Set oleFrame = ws.OLEObjects.Add("Forms.Frame.1", Left:=PANEL_LEFT, Top:=PANEL_TOP, Width:=PANEL_WIDTH, Height:=PANEL_HEIGHT)
    oleFrame.Name = "pan_Assets"
    
    Set realFrame = oleFrame.Object
    realFrame.Caption = "": realFrame.BackColor = COLOR_GRAY_BG: realFrame.SpecialEffect = 0
    
    Set lblTitle = realFrame.Controls.Add("Forms.Label.1", "lbl_Title", True)
    With lblTitle
        .Caption = UI_TITLE_ASSETS: .Font.Name = "Segoe UI": .Font.Bold = True: .Font.Size = 12
        .Left = 140: .Top = 35: .Width = 200: .Height = 25: .BackStyle = 0
    End With
    
    Set btnBack = realFrame.Controls.Add("Forms.CommandButton.1", "btn_BackToLogin", True)
    With btnBack
        .Caption = UI_BTN_BACK: .Font.Name = "Segoe UI": .Font.Bold = True
        .Width = 100: .Height = 35: .Left = 20: .Top = 30: .BackColor = COLOR_GREEN: .ForeColor = COLOR_WHITE
    End With
    
    Call AddEngineerToFrame(ws, realFrame)
    Call RegisterButton(btnBack)
    oleFrame.Visible = False
End Sub

' Build Personnel verification panel
Public Sub CreatePersonnelPanel(ws As Worksheet)
    Dim oleFrame As OLEObject, realFrame As MSForms.Frame
    Dim btnBack As MSForms.CommandButton, btnVerify As MSForms.CommandButton
    Dim txtInput As MSForms.TextBox, lblTitle As MSForms.Label
    
    Set oleFrame = ws.OLEObjects.Add("Forms.Frame.1", Left:=PANEL_LEFT, Top:=PANEL_TOP, Width:=PANEL_WIDTH, Height:=PANEL_HEIGHT)
    oleFrame.Name = "pan_Personnel"
    
    Set realFrame = oleFrame.Object
    realFrame.Caption = "": realFrame.BackColor = COLOR_GRAY_BG: realFrame.SpecialEffect = 0
    
    Set lblTitle = realFrame.Controls.Add("Forms.Label.1", "lbl_InputTitle", True)
    With lblTitle
        .Caption = UI_LBL_PERSONNEL_ID: .Font.Name = "Segoe UI": .Font.Size = 10
        .Left = 210: .Top = (PANEL_HEIGHT - 20) / 2: .Width = 120: .Height = 20: .BackStyle = 0
    End With
    
    Set txtInput = realFrame.Controls.Add("Forms.TextBox.1", "txt_PersonnelID", True)
    With txtInput
        .Font.Name = "Segoe UI": .Font.Size = 11: .Left = 275: .Top = (PANEL_HEIGHT - 30) / 2
        .Width = 160: .Height = 30: .BorderStyle = 1: .SpecialEffect = 0
    End With
    
    Set btnVerify = realFrame.Controls.Add("Forms.CommandButton.1", "btn_VerifyPersonnel", True)
    With btnVerify
        .Caption = UI_BTN_VERIFY: .Font.Name = "Segoe UI": .Font.Bold = True
        .Width = 110: .Height = 30: .Left = 445: .Top = (PANEL_HEIGHT - 30) / 2
        .BackColor = COLOR_GREEN: .ForeColor = COLOR_WHITE
    End With
    
    Set btnBack = realFrame.Controls.Add("Forms.CommandButton.1", "btn_BackToLogin", True)
    With btnBack
        .Caption = UI_BTN_BACK: .Font.Name = "Segoe UI": .Font.Bold = True
        .Width = 100: .Height = 35: .Left = 20: .Top = PANEL_HEIGHT - 55
        .BackColor = COLOR_GREEN: .ForeColor = COLOR_WHITE
    End With
    
    Call AddEngineerToFrame(ws, realFrame)
    Call RegisterButton(btnBack): Call RegisterButton(btnVerify)
    oleFrame.Visible = False
End Sub

' Build Verified info panel
Public Sub CreateVerifiedPanel(ws As Worksheet)
    Dim oleFrame As OLEObject, realFrame As MSForms.Frame
    Dim lblWelcome As MSForms.Label, lblInfo As MSForms.Label, lblShiftPrefix As MSForms.Label
    Dim lblShiftValue As MSForms.Label, lblShiftSuffix As MSForms.Label, lblError As MSForms.Label
    Dim lblThanks As MSForms.Label, btnLogout As MSForms.CommandButton
    Dim activeID As String, personName As String, shiftCount As Double
    
    On Error Resume Next: ws.OLEObjects("pan_Verified").Delete: On Error GoTo 0
    
    Set oleFrame = ws.OLEObjects.Add("Forms.Frame.1", Left:=PANEL_LEFT, Top:=PANEL_TOP, Width:=PANEL_WIDTH, Height:=PANEL_HEIGHT)
    oleFrame.Name = "pan_Verified"
    
    Set realFrame = oleFrame.Object
    realFrame.Caption = "": realFrame.BackColor = COLOR_GRAY_BG: realFrame.SpecialEffect = 0
    
    activeID = GetActivePersonnel()
    personName = getNameByID(activeID)
    If personName = "" Then personName = activeID
    
    Set lblWelcome = realFrame.Controls.Add("Forms.Label.1", "lbl_Welcome", True)
    With lblWelcome
        .Caption = UI_VERIFIED_HELLO & personName: .Font.Name = "Segoe UI": .Font.Size = 14: .Font.Bold = True
        .Width = 500: .Left = (PANEL_WIDTH - 500) / 2: .Top = 55: .Height = 25: .BackStyle = 0: .TextAlign = 2
    End With
    
    Set lblInfo = realFrame.Controls.Add("Forms.Label.1", "lbl_Info", True)
    With lblInfo
        .Caption = GetVerifiedPersonnelInfoText(activeID, shiftCount): .Font.Name = "Segoe UI": .Font.Size = 10
        .Width = 600: .Left = (PANEL_WIDTH - 600) / 2: .Top = 90: .Height = 75: .BackStyle = 0: .WordWrap = True: .TextAlign = 2
    End With
    
    Set lblShiftPrefix = realFrame.Controls.Add("Forms.Label.1", "lbl_ShiftPrefix", True)
    With lblShiftPrefix
        .Caption = UI_VERIFIED_PREFIX: .Font.Name = "Segoe UI": .Font.Size = 10
        .Left = 160: .Top = 174: .Width = 105: .Height = 15: .BackStyle = 0: .ForeColor = vbBlack: .TextAlign = 3
    End With
    
    Set lblShiftValue = realFrame.Controls.Add("Forms.Label.1", "lbl_ShiftValue", True)
    With lblShiftValue
        .Caption = CStr(shiftCount): .Font.Name = "Segoe UI": .Font.Size = 10
        .Left = 265: .Top = 174: .Width = 35: .Height = 15: .BackStyle = 0: .TextAlign = 2
    End With
    
    Set lblShiftSuffix = realFrame.Controls.Add("Forms.Label.1", "lbl_ShiftSuffix", True)
    With lblShiftSuffix
        .Caption = UI_VERIFIED_SUFFIX: .Font.Name = "Segoe UI": .Font.Size = 10
        .Left = 300: .Top = 174: .Width = 240: .Height = 15: .BackStyle = 0: .ForeColor = vbBlack: .TextAlign = 1
    End With
    
    Set lblError = realFrame.Controls.Add("Forms.Label.1", "lbl_Error", True)
    With lblError
        .Font.Name = "Segoe UI": .Font.Size = 10: .Width = 500: .Left = (PANEL_WIDTH - 500) / 2: .Top = 202
        .Height = 15: .BackStyle = 0: .TextAlign = 2
    End With
    
    Set lblThanks = realFrame.Controls.Add("Forms.Label.1", "lbl_Thanks", True)
    With lblThanks
        .Caption = UI_VERIFIED_THANKS: .Font.Name = "Segoe UI": .Font.Size = 10: .Font.Bold = False
        .ForeColor = vbBlack: .Width = 500: .Left = (PANEL_WIDTH - 500) / 2: .Height = 15: .BackStyle = 0: .TextAlign = 2
    End With
    
    Set btnLogout = realFrame.Controls.Add("Forms.CommandButton.1", "btn_Logout", True)
    With btnLogout
        .Caption = UI_BTN_BACK: .Font.Name = "Segoe UI": .Font.Bold = True
        .Width = 100: .Height = 35: .Left = 20: .Top = PANEL_HEIGHT - 55: .BackColor = COLOR_RED: .ForeColor = COLOR_WHITE
    End With
    
    Call AddEngineerToFrame(ws, realFrame)
    Call RegisterButton(btnLogout)
    oleFrame.Visible = False
End Sub

' Refresh data inside Verified panel
Public Sub RefreshVerifiedPanel()
    Dim ws As Worksheet, realFrame As MSForms.Frame
    Dim activeID As String, personName As String, shiftCount As Double
    Set ws = ThisWorkbook.Sheets("PORTAL")
    Set realFrame = ws.OLEObjects("pan_Verified").Object
    
    activeID = GetActivePersonnel()
    personName = getNameByID(activeID)
    If personName = "" Then personName = activeID
    
    realFrame.Controls("lbl_Welcome").Caption = UI_VERIFIED_HELLO & personName
    realFrame.Controls("lbl_Info").Caption = GetVerifiedPersonnelInfoText(activeID, shiftCount)
    realFrame.Controls("lbl_ShiftValue").Caption = CStr(shiftCount)
    
    ' Update shift value style based on threshold
    With realFrame.Controls("lbl_ShiftValue")
        If shiftCount > 200 Then
            .ForeColor = RGB(255, 0, 0)
            .Font.Bold = True
        ElseIf shiftCount > 100 Then
            .ForeColor = RGB(255, 140, 0)
            .Font.Bold = False
        Else
            .ForeColor = vbBlack
            .Font.Bold = False
        End If
    End With
    
    With realFrame.Controls("lbl_Error")
        If shiftCount > 200 Then .Caption = UI_VERIFIED_WARN_LIMIT: .ForeColor = vbRed: .Font.Bold = True: .Visible = True _
        Else: .Caption = "": .Visible = False
    End With
    
    realFrame.Controls("lbl_Thanks").Top = IIf(shiftCount > 200, 230, 202)
End Sub

' Helper to attach click listener to dynamic button
Private Sub RegisterButton(btn As MSForms.CommandButton)
    Dim listener As New clsDynamicButton
    Set listener.DynamicButton = btn
    ButtonCollection.Add listener
End Sub

' Render shape into Image control
Public Sub LoadEmbeddedShapeToControl(ws As Worksheet, shapeName As String, targetControl As MSForms.Image, Optional bgRGB As Long = -1)
    Dim shp As Shape, chartObj As ChartObject, tempFilePath As String, origActiveSheet As Worksheet
    Set origActiveSheet = ActiveSheet
    
    On Error Resume Next: Set shp = ws.Shapes(shapeName): On Error GoTo 0
    
    If Not shp Is Nothing Then
        ws.Activate
        tempFilePath = Environ("TEMP") & "\temp_" & shapeName & ".jpg"
        shp.Visible = msoTrue: shp.Select: Selection.CopyPicture Appearance:=xlScreen, Format:=xlPicture: shp.Visible = msoFalse
        
        Set chartObj = ws.ChartObjects.Add(Left:=0, Top:=0, Width:=shp.Width, Height:=shp.Height)
        With chartObj
            .Border.LineStyle = xlNone: .Shadow = False
            If bgRGB <> -1 Then .Chart.ChartArea.Format.Fill.ForeColor.RGB = bgRGB
            .Activate: .Chart.Paste: .Chart.Export Filename:=tempFilePath, FilterName:="JPG": .Delete
        End With
        
        If Dir(tempFilePath) <> "" Then targetControl.Picture = LoadPicture(tempFilePath): Kill tempFilePath
        origActiveSheet.Activate
    Else
        MsgBox Replace(MSG_ERR_SHAPE_NOT_FOUND_BODY, "%s", shapeName), vbExclamation, MSG_ERR_SHAPE_NOT_FOUND_TITLE
    End If
End Sub

' Inject Shared Engineer Image
Public Sub AddEngineerToFrame(ws As Worksheet, targetFrame As MSForms.Frame)
    Dim imgEngineer As MSForms.Image
    Set imgEngineer = targetFrame.Controls.Add("Forms.Image.1", "img_SharedEngineer", True)
    With imgEngineer
        .Width = ENGINEER_WIDTH: .Height = ENGINEER_HEIGHT
        .Top = targetFrame.InsideHeight - .Height
        .Left = targetFrame.InsideWidth - .Width - 15
        .PictureSizeMode = 1: .BorderStyle = 0: .BackStyle = 0: .BackColor = COLOR_GRAY_BG
    End With
    Call LoadEmbeddedShapeToControl(ws, "engineers", imgEngineer, RGB(240, 240, 240))
End Sub