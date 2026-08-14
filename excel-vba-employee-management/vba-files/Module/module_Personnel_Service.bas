Attribute VB_Name = "module_Personnel_Service"
' =========================================================================
' MODULE: module_Personnel_Service
' Description: Service layer handling DB queries, verification, and session state.
' =========================================================================

Private CurrentPersonnelID As String

' Set active session personnel ID
Public Sub SetActivePersonnel(ByVal ID As String)
    CurrentPersonnelID = Trim(ID)
End Sub

' Get currently logged-in personnel ID
Public Function GetActivePersonnel() As String
    GetActivePersonnel = CurrentPersonnelID
End Function

' Clear active session
Public Sub ClearActivePersonnel()
    CurrentPersonnelID = ""
End Sub

' Fetch personnel name by ID
Public Function getNameByID(ByVal ID As String) As String
    getNameByID = GetDBValueByColumn(ID, MESAI_DB_PERSONAL_NAME_COLUMN)
End Function

' Fetch job start date by ID
Public Function getJobStartDateByID(ByVal ID As String) As Variant
    getJobStartDateByID = GetDBValueByColumn(ID, MESAI_DB_START_DATE_COLUMN)
End Function

' Fetch deserved off day count
Public Function getDeservedOffDayCountByID(ByVal ID As String) As Long
    getDeservedOffDayCountByID = Val(GetDBValueByColumn(ID, MESAI_DB_DESERVED_OFF_DAY_COLUMN))
End Function

' Fetch current quota
Public Function getCurrentQuotaByID(ByVal ID As String) As Double
    getCurrentQuotaByID = Val(GetDBValueByColumn(ID, MESAI_DB_CURRENT_QUOTA_COLUMN))
End Function

' Fetch used off day count
Public Function getUsedOffDayCountByID(ByVal ID As String) As Long
    getUsedOffDayCountByID = Val(GetDBValueByColumn(ID, MESAI_DB_USED_OFF_DAY_COLUMN))
End Function

' Fetch off day deserve date
Public Function getOffDayDeserveDate(ByVal ID As String) As Variant
    getOffDayDeserveDate = GetDBValueByColumn(ID, MESAI_DB_OFF_DAY_DESERVE_DATE_COLUMN)
End Function

' Check if personnel ID exists in DB
Public Function VerifyPersonnelID(ByVal ID As String) As Boolean
    VerifyPersonnelID = (GetTargetRowByID(ID) > 0)
End Function

' Traverse row to find last filled shift column value
Public Function getSummedShiftCountById(ByVal ID As String) As Variant
    Dim wsDb As Worksheet, targetRow As Long, colIndex As Long
    Dim emptyCount As Integer, lastFoundCol As Long, cellValue As Variant
    
    getSummedShiftCountById = ""
    targetRow = GetTargetRowByID(ID)
    If targetRow = 0 Then Exit Function
    
    Set wsDb = ThisWorkbook.Sheets(MESAI_DB_SHEET_NAME)
    colIndex = wsDb.Columns(MESAI_DB_PERSONNEL_COLUMN).Column + 1
    
    Do While colIndex <= wsDb.Columns.Count
        cellValue = wsDb.Cells(targetRow, colIndex).Value
        If IsEmpty(cellValue) Or Trim(CStr(cellValue)) = "" Then
            emptyCount = emptyCount + 1
        Else
            lastFoundCol = colIndex
            emptyCount = 0
        End If
        If emptyCount >= 30 Then Exit Do
        colIndex = colIndex + 1
    Loop
    
    If lastFoundCol > 0 Then getSummedShiftCountById = wsDb.Cells(targetRow, lastFoundCol).Value
End Function

' Compile verified info label text
Public Function GetVerifiedPersonnelInfoText(ByVal ID As String, ByRef outShiftCount As Double) As String
    Dim jobStartDate As Variant, deserveDate As Variant
    Dim deservedDays As Long, currentQuota As Double, usedDays As Long, shiftVal As Variant
    Dim strStartDate As String, strDeserveDate As String, cleanVal As String
    
    jobStartDate = getJobStartDateByID(ID)
    deserveDate = getOffDayDeserveDate(ID)
    deservedDays = getDeservedOffDayCountByID(ID)
    currentQuota = getCurrentQuotaByID(ID)
    usedDays = getUsedOffDayCountByID(ID)
    shiftVal = getSummedShiftCountById(ID)
    
    strStartDate = IIf(IsDate(jobStartDate), Format(jobStartDate, "dd.mm.yyyy"), UI_TEXT_NOT_SPECIFIED)
    strDeserveDate = IIf(IsDate(deserveDate), Format(deserveDate, "dd.mm.yyyy"), UI_TEXT_NOT_SPECIFIED)
    
    outShiftCount = 0
    If Not IsEmpty(shiftVal) And Trim(CStr(shiftVal)) <> "" Then
        cleanVal = Trim(CStr(shiftVal))
        cleanVal = Replace(cleanVal, IIf(Mid(Format(1 / 2, "0.0"), 2, 1) = ",", ".", ","), Mid(Format(1 / 2, "0.0"), 2, 1))
        If IsNumeric(cleanVal) Then outShiftCount = CDbl(cleanVal)
    End If

    ' Call template builder from module_Config
    GetVerifiedPersonnelInfoText = BuildVerifiedInfoText(strStartDate, strDeserveDate, deservedDays, usedDays, currentQuota)
End Function

' --- PRIVATE HELPER METHODS ---

Private Function GetDBWorksheet() As Worksheet
    On Error Resume Next
    Set GetDBWorksheet = ThisWorkbook.Sheets(MESAI_DB_SHEET_NAME)
    On Error GoTo 0
    If GetDBWorksheet Is Nothing Then
        MsgBox MSG_ERR_DB_NOT_FOUND_BODY, vbCritical, MSG_ERR_DB_NOT_FOUND_TITLE
    End If
End Function

Private Function GetTargetRowByID(ByVal ID As String) As Long
    Dim wsDb As Worksheet, lastRow As Long, cell As Range
    ID = Trim(ID): GetTargetRowByID = 0
    If ID = "" Then Exit Function
    
    Set wsDb = GetDBWorksheet()
    If wsDb Is Nothing Then Exit Function
    
    lastRow = wsDb.Cells(wsDb.Rows.Count, MESAI_DB_PERSONNEL_COLUMN).End(xlUp).Row
    For Each cell In wsDb.Range(MESAI_DB_PERSONNEL_COLUMN & "1:" & MESAI_DB_PERSONNEL_COLUMN & lastRow)
        If CStr(cell.Value) = ID Then
            GetTargetRowByID = cell.Row
            Exit Function
        End If
    Next cell
End Function

Private Function GetDBValueByColumn(ByVal ID As String, ByVal colLetter As String) As Variant
    Dim targetRow As Long, wsDb As Worksheet
    GetDBValueByColumn = ""
    targetRow = GetTargetRowByID(ID)
    If targetRow > 0 Then
        Set wsDb = GetDBWorksheet()
        GetDBValueByColumn = wsDb.Cells(targetRow, colLetter).Value
    End If
End Function