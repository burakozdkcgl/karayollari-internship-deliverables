VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsDynamicButton"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

'*****************************************************************************
' Class: clsDynamicButton
' Purpose: This class is used to handle events for dynamically created buttons.
'*****************************************************************************

Public WithEvents DynamicButton As MSForms.CommandButton

Private Sub DynamicButton_Click()
    Dim parentFrame As MSForms.Frame
    
    On Error Resume Next
    Set parentFrame = DynamicButton.Parent
    On Error GoTo 0

    If parentFrame Is Nothing Then Exit Sub

    Select Case parentFrame.Name
        Case "pan_Login":               mod_UI_Events.HandleLoginPanelEvents DynamicButton.Name, parentFrame
        Case "pan_OSM":                 mod_UI_Events.HandleOSMPanelEvents DynamicButton.Name, parentFrame
        Case "pan_Open":                mod_UI_Events.HandleOpenPanelEvents DynamicButton.Name, parentFrame
        Case "pan_Closed":              mod_UI_Events.HandleClosedPanelEvents DynamicButton.Name, parentFrame
        Case "pan_PreviousMonth":       mod_UI_Events.HandlePreviousMonthPanelEvents DynamicButton.Name, parentFrame
        Case "pan_SelectPreviousMonth": mod_UI_Events.HandleSelectPreviousMonthEvents DynamicButton.Name, parentFrame
        Case "pan_PrevMonthClosed":     mod_UI_Events.HandlePrevMonthClosedEvents DynamicButton.Name, parentFrame
        Case "pan_PrevMonthOpen":       mod_UI_Events.HandlePrevMonthOpenEvents DynamicButton.Name, parentFrame
    End Select
End Sub