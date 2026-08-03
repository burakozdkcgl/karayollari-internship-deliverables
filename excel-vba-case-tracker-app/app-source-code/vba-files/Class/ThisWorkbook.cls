VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True

'*****************************************************************************
' Class: ThisWorkbook
' Purpose: This class is a standard class module that represents the excel file and handles workbook-level events.
'*****************************************************************************

Private Sub Workbook_Open()
    DoEvents
    Application.OnTime Now + TimeValue("00:00:01"), "InitializeAppUI"
End Sub