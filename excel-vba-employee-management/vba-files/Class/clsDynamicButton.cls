VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsDynamicButton"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

' =========================================================================
' CLASS MODULE: clsDynamicButton
' Description: Class listening to click events for dynamic Frame buttons.
' =========================================================================

Public WithEvents DynamicButton As MSForms.CommandButton

Private Sub DynamicButton_Click()
    Dim wsPortal As Worksheet, framePersonnel As MSForms.Frame
    Dim txtInput As MSForms.TextBox, inputVal As String
    
    Set wsPortal = ThisWorkbook.Sheets("PORTAL")
    
    Select Case DynamicButton.Name
        Case "btn_Asset"
            Call SwitchTo("Assets")
            
        Case "btn_Personnel"
            Call SwitchTo("Personnel")
            
        Case "btn_BackToLogin", "btn_Logout"
            Call ClearActivePersonnel
            Call SwitchTo("Login")
            
        Case "btn_VerifyPersonnel"
            On Error Resume Next
            Set framePersonnel = wsPortal.OLEObjects("pan_Personnel").Object
            Set txtInput = framePersonnel.Controls("txt_PersonnelID")
            On Error GoTo 0
            
            If Not txtInput Is Nothing Then
                inputVal = Trim(txtInput.Value)
                
                If VerifyPersonnelID(inputVal) Then
                    Call SetActivePersonnel(inputVal)
                    txtInput.Value = ""
                    Call SwitchTo("Verified")
                Else
                    MsgBox MSG_ERR_VERIFY_FAIL_BODY, vbExclamation, MSG_ERR_VERIFY_FAIL_TITLE
                    txtInput.SetFocus
                End If
            End If
    End Select
End Sub