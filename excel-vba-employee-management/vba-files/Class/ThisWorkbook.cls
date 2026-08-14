VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True

' =========================================================================
' CLASS: ThisWorkbook
' Description: Application startup orchestrator.
' =========================================================================

Private Sub Workbook_Open()
    Dim ws As Worksheet
    
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("PORTAL")
    On Error GoTo 0
    
    If Not ws Is Nothing Then
        Call BuildNavigationSystem
        Call SwitchTo("Login")
        
        With ws
            .ScrollArea = "A1"
            .EnableSelection = xlNoSelection
        End With
    End If
End Sub