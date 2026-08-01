Attribute VB_Name = "mod_Services"
Option Explicit

'*****************************************************************************
' Module: mod_Services
' Purpose: This module contains public functions for handling business logic related to case and maintenance reporting.
'*****************************************************************************

' Retrieves the report date from the specified worksheet header row.
Public Function GetReportDate(ws As Worksheet) As String
    Dim col As Long, cellVal As Variant
    GetReportDate = STATUS_NOT_FOUND
    
    For col = 1 To 50
        cellVal = ws.Cells(1, col).Value
        If Not IsEmpty(cellVal) And IsDate(cellVal) Then
            GetReportDate = Format(CDate(cellVal), "dd.mm.yyyy")
            Exit Function
        End If
    Next col
End Function

' Maps source sheet names to corresponding target database sheets.
Public Function GetTargetDbSheetName(ByVal sheetName As String) As String
    If InStr(1, sheetName, OPEN_MAINT_SHEET_NAME, vbTextCompare) > 0 Or _
       InStr(1, sheetName, CLOSED_MAINT_SHEET_NAME, vbTextCompare) > 0 Or _
       InStr(1, sheetName, DATABASE_SHEET_NAME_BAKIM, vbTextCompare) > 0 Then
        GetTargetDbSheetName = DATABASE_SHEET_NAME_BAKIM
    Else
        GetTargetDbSheetName = DATABASE_SHEET_NAME_ARIZA
    End If
End Function

' Scans the workbook for relevant case/maintenance reporting sheets.
Public Function ScanWorkbookSheets() As Collection
    Dim ws As Worksheet, results As New Collection
    Dim isOpenCase As Boolean, isClosedCase As Boolean
    Dim isOpenMaint As Boolean, isClosedMaint As Boolean
    Dim rDate As String, eDate As String, rType As String
    
    For Each ws In ThisWorkbook.Worksheets
        isOpenCase = (InStr(1, ws.Name, OPEN_CASES_SHEET_NAME, vbTextCompare) > 0)
        isClosedCase = (InStr(1, ws.Name, CLOSED_CASES_SHEET_NAME, vbTextCompare) > 0)
        isOpenMaint = (InStr(1, ws.Name, OPEN_MAINT_SHEET_NAME, vbTextCompare) > 0)
        isClosedMaint = (InStr(1, ws.Name, CLOSED_MAINT_SHEET_NAME, vbTextCompare) > 0)
        
        If isOpenCase Or isClosedCase Or isOpenMaint Or isClosedMaint Then
            If isOpenCase Then rType = RPT_TYPE_OPEN_CASE
            If isClosedCase Then rType = RPT_TYPE_CLOSED_CASE
            If isOpenMaint Then rType = RPT_TYPE_OPEN_MAINT
            If isClosedMaint Then rType = RPT_TYPE_CLOSED_MAINT
            
            rDate = GetReportDate(ws)
            eDate = STATUS_NOT_FOUND
            
            If rDate <> STATUS_NOT_FOUND And IsDate(rDate) Then
                eDate = Format(CDate(rDate) - 1, "dd.mm.yyyy")
            End If
            
            results.Add Array(STATUS_OK, rType, rDate, eDate, ws.Name)
        End If
    Next ws
    
    Set ScanWorkbookSheets = results
End Function

' Checks if any open cases exist across sheets.
Public Function HasAnyOpenCase() As Boolean
    Dim item As Variant
    For Each item In ScanWorkbookSheets()
        If (item(1) = RPT_TYPE_OPEN_CASE Or item(1) = RPT_TYPE_OPEN_MAINT) And item(2) <> STATUS_NOT_FOUND And Trim(item(2)) <> "" Then
            HasAnyOpenCase = True
            Exit Function
        End If
    Next item
    HasAnyOpenCase = False
End Function

' Calculates metrics for open cases and finds target DB addresses.
Public Function GetOpenCasesCalculated() As Collection
    Dim item As Variant, targetWs As Worksheet
    Dim openCol As New Collection, cellAddr As String
    Dim lastRow As Long, startRow As Long, caseCount As Long
    Dim etkiDateStr As String, dbName As String
    
    For Each item In ScanWorkbookSheets()
        If (item(1) = RPT_TYPE_OPEN_CASE Or item(1) = RPT_TYPE_OPEN_MAINT) And item(2) <> STATUS_NOT_FOUND Then
            On Error Resume Next
            Set targetWs = ThisWorkbook.Sheets(CStr(item(4)))
            On Error GoTo 0
            
            caseCount = 0
            If Not targetWs Is Nothing Then
                If Trim(CStr(targetWs.Range(OPEN_SHEET_FIRST_CASE_CELL).Value)) <> "" Then
                    lastRow = targetWs.Cells(targetWs.Rows.Count, "A").End(xlUp).Row
                    startRow = targetWs.Range(OPEN_SHEET_FIRST_CASE_CELL).Row
                    caseCount = IIf(lastRow >= startRow, (lastRow - startRow) + 1, 1)
                End If
            End If
            
            etkiDateStr = CStr(item(3))
            dbName = GetTargetDbSheetName(CStr(item(4)))
            cellAddr = FindIntersectCell(dbName, etkiDateStr, CStr(DB_OPEN_CELL_COL), CStr(DB_OPEN_CELL_CONTEXT))
            
            openCol.Add Array(item(4), item(2), item(3), caseCount, cellAddr)
        End If
    Next item
    
    Set GetOpenCasesCalculated = openCol
End Function

' Checks if any closed cases exist across sheets.
Public Function HasAnyClosedCase() As Boolean
    Dim item As Variant
    For Each item In ScanWorkbookSheets()
        If (item(1) = RPT_TYPE_CLOSED_CASE Or item(1) = RPT_TYPE_CLOSED_MAINT) And item(2) <> STATUS_NOT_FOUND And Trim(item(2)) <> "" Then
            HasAnyClosedCase = True
            Exit Function
        End If
    Next item
    HasAnyClosedCase = False
End Function

' Calculates metrics for closed cases grouped by assigned personnel.
Public Function GetClosedCasesCalculated() As Collection
    Dim item As Variant, targetWs As Worksheet
    Dim closedCol As New Collection, dict As Object, key As Variant
    Dim startRow As Long, lastRow As Long, i As Long
    Dim pName As String, targetCell As String, dbName As String
    Dim nameCol As String, firstCell As String
    
    Set dict = CreateObject("Scripting.Dictionary")
    
    For Each item In ScanWorkbookSheets()
        If (item(1) = RPT_TYPE_CLOSED_CASE Or item(1) = RPT_TYPE_CLOSED_MAINT) And item(2) <> STATUS_NOT_FOUND Then
            On Error Resume Next
            Set targetWs = ThisWorkbook.Sheets(CStr(item(4)))
            On Error GoTo 0
            
            If Not targetWs Is Nothing Then
                If item(1) = RPT_TYPE_CLOSED_MAINT Then
                    nameCol = CLOSED_MAINT_SHEET_NAME_COL
                    firstCell = CLOSED_MAINT_SHEET_FIRST_CELL
                Else
                    nameCol = CLOSED_CASE_SHEET_NAME_COL
                    firstCell = CLOSED_CASE_SHEET_FIRST_CELL
                End If
                
                startRow = targetWs.Range(firstCell).Row
                lastRow = targetWs.Cells(targetWs.Rows.Count, nameCol).End(xlUp).Row
                
                If lastRow >= startRow Then
                    For i = startRow To lastRow
                        pName = Trim(CStr(targetWs.Cells(i, nameCol).Value))
                        If pName <> "" Then dict(pName) = dict(pName) + 1
                    Next i
                End If
            End If
            
            dbName = GetTargetDbSheetName(CStr(item(4)))
            For Each key In dict.Keys
                targetCell = FindIntersectCell(dbName, CStr(item(3)), DB_CLOSED_NAME_COL, CStr(key))
                
                If targetCell <> STATUS_NOT_FOUND And targetCell <> STATUS_NO_DB_SHEET Then
                    closedCol.Add Array(item(4), item(2), item(3), CStr(key), dict(key), targetCell)
                End If
            Next key
            
            dict.RemoveAll
        End If
    Next item
    
    Set GetClosedCasesCalculated = SortCollectionByDateIndex(closedCol, 3)
End Function

' Intersects target date columns and context rows to return cell addresses.
Public Function FindIntersectCell(ByVal dbSheetName As String, ByVal targetDateStr As String, ByVal searchColName As String, ByVal contextText As String) As String
    Dim dbWs As Worksheet, startCol As Long, endCol As Long, searchCol As Long
    Dim tRow As Long, tCol As Long, i As Long, dbLastRow As Long
    
    FindIntersectCell = STATUS_NOT_FOUND
    On Error Resume Next
    Set dbWs = ThisWorkbook.Sheets(dbSheetName)
    On Error GoTo 0
    
    If dbWs Is Nothing Then FindIntersectCell = STATUS_NO_DB_SHEET: Exit Function
    If Not IsDate(targetDateStr) Then Exit Function
    
    startCol = dbWs.Columns(DB_OSM_DATES_START_COL).Column
    endCol = dbWs.Columns(DB_OSM_DATES_END_COL).Column
    searchCol = dbWs.Columns(searchColName).Column
    
    For i = startCol To endCol
        If IsDate(dbWs.Cells(1, i).Value) Then
            If CDate(dbWs.Cells(1, i).Value) = CDate(targetDateStr) Then
                tCol = i: Exit For
            End If
        End If
    Next i
    
    dbLastRow = dbWs.Cells(dbWs.Rows.Count, searchCol).End(xlUp).Row
    For i = 1 To dbLastRow
        If InStr(1, CStr(dbWs.Cells(i, searchCol).Value), contextText, vbTextCompare) > 0 Then
            tRow = i: Exit For
        End If
    Next i
    
    If tRow > 0 And tCol > 0 Then
        FindIntersectCell = dbWs.Cells(tRow, tCol).Address(RowAbsolute:=False, ColumnAbsolute:=False)
    End If
End Function

' Writes updated ListBox values into database cells.
Public Function ExecuteDatabaseWrite(lst As MSForms.ListBox, cellColIdx As Long, valColIdx As Long, sheetNameColIdx As Long) As Boolean
    Dim dbWs As Worksheet, i As Long, addr As String, caseVal As Long
    Dim srcSheetName As String, targetDbName As String
    Dim rawAddr As Variant, rawSheet As Variant
    
    For i = 1 To lst.ListCount - 1
        rawAddr = lst.List(i, cellColIdx)
        rawSheet = lst.List(i, sheetNameColIdx)
        
        addr = Trim("" & rawAddr)
        srcSheetName = Trim("" & rawSheet)
        
        If InStr(addr, "(") > 0 Then
            addr = Trim(Split(addr, "(")(0))
        End If
        
        If addr <> "" And addr <> STATUS_NOT_FOUND And addr <> STATUS_NO_DB_SHEET Then
            targetDbName = GetTargetDbSheetName(srcSheetName)
            
            Set dbWs = Nothing
            On Error Resume Next
            Set dbWs = ThisWorkbook.Sheets(targetDbName)
            On Error GoTo 0
            
            If dbWs Is Nothing Then
                MsgBox targetDbName & MSG_ERR_TARGET_DB_NOT_FOUND, vbCritical, MSG_TITLE_ERROR
                ExecuteDatabaseWrite = False
                Exit Function
            End If
            
            caseVal = Val("" & lst.List(i, valColIdx))
            dbWs.Range(addr).Value = caseVal
        End If
    Next i
    
    ExecuteDatabaseWrite = True
End Function

' Sorts array collection elements chronologically based on date column index.
Public Function SortCollectionByDateIndex(col As Collection, dateIdx As Long) As Collection
    If col.Count <= 1 Then Set SortCollectionByDateIndex = col: Exit Function
    
    Dim items() As Variant, temp As Variant, i As Long, j As Long
    ReDim items(1 To col.Count)
    For i = 1 To col.Count: items(i) = col(i): Next i
    
    For i = 1 To UBound(items) - 1
        For j = i + 1 To UBound(items)
            If IsDate(items(i)(dateIdx - 1)) And IsDate(items(j)(dateIdx - 1)) Then
                If CDate(items(i)(dateIdx - 1)) > CDate(items(j)(dateIdx - 1)) Then
                    temp = items(i): items(i) = items(j): items(j) = temp
                End If
            End If
        Next j
    Next i
    
    Dim sortedCol As New Collection
    For i = 1 To UBound(items): sortedCol.Add items(i): Next i
    Set SortCollectionByDateIndex = sortedCol
End Function

' Maps type IDs to human-readable UI strings.
Public Function GetTypeStringByID(ByVal caseTypeID As String) As String
    Select Case StrConv(Trim(caseTypeID), vbUpperCase)
        Case RPT_TYPE_OPEN_CASE:    GetTypeStringByID = CASE_TYPE_OPEN_STR
        Case RPT_TYPE_CLOSED_CASE:  GetTypeStringByID = CASE_TYPE_CLOSED_STR
        Case RPT_TYPE_OPEN_MAINT:   GetTypeStringByID = CASE_TYPE_MAINT_OPEN_STR
        Case RPT_TYPE_CLOSED_MAINT: GetTypeStringByID = CASE_TYPE_MAINT_CLOSED_STR
        Case Else:                  GetTypeStringByID = caseTypeID
    End Select
End Function

' Scans sheets required for previous month analysis.
Public Function ScanPreviousMonthSheets() As Collection
    Dim results As New Collection, ws As Worksheet, rType As String
    Dim isAllCases As Boolean, isAllMaint As Boolean
    
    For Each ws In ThisWorkbook.Worksheets
        isAllCases = (InStr(1, ws.Name, SHEET_NAME_ALL_CASES, vbTextCompare) > 0)
        isAllMaint = (InStr(1, ws.Name, SHEET_NAME_ALL_MAINT, vbTextCompare) > 0)
        
        If isAllCases Or isAllMaint Then
            rType = IIf(isAllCases, RPT_TYPE_CASE_REPORT, RPT_TYPE_MAINT_REPORT)
            results.Add Array(STATUS_OK, rType, ws.Name)
        End If
    Next ws
    
    Set ScanPreviousMonthSheets = results
End Function

' Calculates metrics for closed cases of previous month.
Public Function GetPreviousMonthClosedCalculated(targetMonth As Integer, targetYear As Integer) As Collection
    Dim prevCol As New Collection
    Dim gWs As Worksheet, wsCases As Worksheet, wsMaint As Worksheet
    Dim totalCases As Long, totalMaint As Long
    Dim cellAddrCases As String, cellAddrMaint As String
    
    On Error Resume Next
    Set gWs = ThisWorkbook.Sheets(PREVIOUS_MONTH_SHEET_NAME)
    Set wsCases = ThisWorkbook.Sheets(SHEET_NAME_ALL_CASES)
    Set wsMaint = ThisWorkbook.Sheets(SHEET_NAME_ALL_MAINT)
    On Error GoTo 0
    
    If Not wsCases Is Nothing Then
        totalCases = CountClosedCasesByPerson(wsCases, ALL_CASES_ASSIGNED_COL, ALL_CASES_CLOSE_DATE_COL, ALL_CASES_STATUS_COL, targetMonth, targetYear)
        cellAddrCases = FindPreviousMonthIntersectCell(DATABASE_SHEET_NAME_ARIZA)
        prevCol.Add Array(SHEET_NAME_ALL_CASES, CATEGORY_ARIZA, totalCases, cellAddrCases)
    End If
    
    If Not wsMaint Is Nothing Then
        totalMaint = CountClosedCasesByPerson(wsMaint, ALL_MAINT_ASSIGNED_COL, ALL_MAINT_CLOSE_DATE_COL, ALL_MAINT_STATUS_COL, targetMonth, targetYear)
        cellAddrMaint = FindPreviousMonthIntersectCell(DATABASE_SHEET_NAME_BAKIM)
        prevCol.Add Array(SHEET_NAME_ALL_MAINT, CATEGORY_BAKIM, totalMaint, cellAddrMaint)
    End If
    
    Set GetPreviousMonthClosedCalculated = prevCol
End Function

' Writes updated values to previous month's sheet.
Public Function ExecutePreviousMonthWrite(lst As MSForms.ListBox, Optional cellColIdx As Long = 3, Optional valColIdx As Long = 2) As Boolean
    Dim gWs As Worksheet, i As Long, addr As String, valCount As Long
    Dim rawAddr As Variant
    
    On Error Resume Next
    Set gWs = ThisWorkbook.Sheets(PREVIOUS_MONTH_SHEET_NAME)
    On Error GoTo 0
    
    If gWs Is Nothing Then
        MsgBox PREVIOUS_MONTH_SHEET_NAME & MSG_ERR_PREV_MONTH_SHEET_NOT_FOUND, vbCritical, MSG_TITLE_ERROR
        ExecutePreviousMonthWrite = False
        Exit Function
    End If
    
    For i = 1 To lst.ListCount - 1
        rawAddr = lst.List(i, cellColIdx)
        addr = Trim("" & rawAddr)
        
        If InStr(addr, "(") > 0 Then addr = Trim(Split(addr, "(")(0))
        
        If addr <> "" And addr <> STATUS_NOT_FOUND And addr <> STATUS_NO_DB_SHEET And addr <> STATUS_ERROR Then
            If IsCellAddressValid(gWs, addr) Then
                valCount = Val("" & lst.List(i, valColIdx))
                gWs.Range(addr).Value = valCount
            End If
        End If
    Next i
    
    ExecutePreviousMonthWrite = True
End Function

' Calculates metrics for open cases per person for previous month.
Public Function GetPreviousMonthOpenPersonCalculated(targetMonth As Integer, targetYear As Integer) As Collection
    Dim openCol As New Collection
    Dim wsCases As Worksheet, wsMaint As Worksheet
    Dim dictCases As Object, dictMaint As Object
    Dim key As Variant, targetCell As String
    
    On Error Resume Next
    Set wsCases = ThisWorkbook.Sheets(SHEET_NAME_ALL_CASES)
    Set wsMaint = ThisWorkbook.Sheets(SHEET_NAME_ALL_MAINT)
    On Error GoTo 0
    
    If Not wsCases Is Nothing Then
        Set dictCases = CountOpenCasesByPerson(wsCases, ALL_CASES_ASSIGNED_COL, ALL_CASES_CLOSE_DATE_COL, ALL_CASES_STATUS_COL, targetMonth, targetYear)
        For Each key In dictCases.Keys
            targetCell = FindPreviousMonthPersonIntersectCell(CStr(key), DATABASE_SHEET_NAME_ARIZA)
            If targetCell <> STATUS_NOT_FOUND Then
                openCol.Add Array(SHEET_NAME_ALL_CASES, CATEGORY_ARIZA, CStr(key), dictCases(key), targetCell)
            End If
        Next key
    End If
    
    If Not wsMaint Is Nothing Then
        Set dictMaint = CountOpenCasesByPerson(wsMaint, ALL_MAINT_ASSIGNED_COL, ALL_MAINT_CLOSE_DATE_COL, ALL_MAINT_STATUS_COL, targetMonth, targetYear)
        For Each key In dictMaint.Keys
            targetCell = FindPreviousMonthPersonIntersectCell(CStr(key), DATABASE_SHEET_NAME_BAKIM)
            If targetCell <> STATUS_NOT_FOUND Then
                openCol.Add Array(SHEET_NAME_ALL_MAINT, CATEGORY_BAKIM, CStr(key), dictMaint(key), targetCell)
            End If
        Next key
    End If
    
    Set GetPreviousMonthOpenPersonCalculated = openCol
End Function

' --- Private Helper Functions ---

Private Function CountClosedCasesByPerson(ws As Worksheet, nameColLetter As String, dateColLetter As String, statusColLetter As String, targetMonth As Integer, targetYear As Integer) As Long
    Dim gWs As Worksheet, nameDict As Object
    Dim gLastRow As Long, lastRow As Long, i As Long
    Dim gName As String, personName As String, statusStr As String
    Dim nameColIdx As Long, dateColIdx As Long, statusColIdx As Long
    Dim namesArray As Variant, datesArray As Variant, statusesArray As Variant
    Dim rawDateVal As Variant, dDate As Date
    Dim countTotal As Long, isValidDate As Boolean
    Dim validStatuses() As String, st As Variant, statusMatched As Boolean
    
    Set nameDict = CreateObject("Scripting.Dictionary")
    validStatuses = Split(VALID_CLOSED_STATUSES, ",")
    
    On Error Resume Next
    Set gWs = ThisWorkbook.Sheets(PREVIOUS_MONTH_SHEET_NAME)
    On Error GoTo 0
    
    If Not gWs Is Nothing Then
        gLastRow = gWs.Cells(gWs.Rows.Count, PREV_MONTH_NAME_COL).End(xlUp).Row
        If gLastRow >= 2 Then
            For i = 2 To gLastRow
                gName = StrConv(Trim(CStr(gWs.Cells(i, PREV_MONTH_NAME_COL).Value)), vbUpperCase)
                If gName <> "" Then nameDict(gName) = True
            Next i
        End If
    End If
    
    If nameDict.Count = 0 Then Exit Function
    
    nameColIdx = ws.Columns(nameColLetter).Column
    dateColIdx = ws.Columns(dateColLetter).Column
    statusColIdx = ws.Columns(statusColLetter).Column
    
    lastRow = ws.Cells(ws.Rows.Count, dateColIdx).End(xlUp).Row
    If lastRow < 2 Then Exit Function
    
    namesArray = ws.Range(ws.Cells(2, nameColIdx), ws.Cells(lastRow, nameColIdx)).Value
    datesArray = ws.Range(ws.Cells(2, dateColIdx), ws.Cells(lastRow, dateColIdx)).Value
    statusesArray = ws.Range(ws.Cells(2, statusColIdx), ws.Cells(lastRow, statusColIdx)).Value
    
    countTotal = 0
    For i = 1 To UBound(datesArray, 1)
        rawDateVal = datesArray(i, 1)
        isValidDate = False
        
        If Not IsEmpty(rawDateVal) Then
            If IsDate(rawDateVal) Then
                dDate = CDate(rawDateVal)
                isValidDate = True
            ElseIf IsNumeric(rawDateVal) Then
                On Error Resume Next
                dDate = CDate(rawDateVal)
                isValidDate = (Err.Number = 0)
                On Error GoTo 0
            End If
            
            If isValidDate Then
                If Month(dDate) = targetMonth And Year(dDate) = targetYear Then
                    statusStr = Trim(CStr(statusesArray(i, 1)))
                    statusMatched = False
                    
                    For Each st In validStatuses
                        If InStr(1, statusStr, Trim(CStr(st)), vbTextCompare) > 0 Then
                            statusMatched = True
                            Exit For
                        End If
                    Next st
                    
                    If statusMatched Then
                        personName = StrConv(Trim(CStr(namesArray(i, 1))), vbUpperCase)
                        If nameDict.Exists(personName) Then countTotal = countTotal + 1
                    End If
                End If
            End If
        End If
    Next i
    
    CountClosedCasesByPerson = countTotal
End Function

Private Function FindPreviousMonthIntersectCell(ByVal caseTypeContext As String) As String
    Dim gWs As Worksheet
    Dim searchColIdx As Long, startColIdx As Long, endColIdx As Long
    Dim tRow As Long, tCol As Long, i As Long, gLastRow As Long
    
    FindPreviousMonthIntersectCell = STATUS_NOT_FOUND
    On Error Resume Next
    Set gWs = ThisWorkbook.Sheets(PREVIOUS_MONTH_SHEET_NAME)
    On Error GoTo 0
    
    If gWs Is Nothing Then Exit Function
    
    searchColIdx = gWs.Columns(PREV_MONTH_SEARCH_COL).Column
    startColIdx = gWs.Columns(PREV_MONTH_START_COL).Column
    endColIdx = gWs.Columns(PREV_MONTH_END_COL).Column
    
    gLastRow = gWs.Cells(gWs.Rows.Count, searchColIdx).End(xlUp).Row
    For i = 1 To gLastRow
        If InStr(1, CStr(gWs.Cells(i, searchColIdx).Value), PREV_MONTH_SEARCH_ROW_CONTEXT, vbTextCompare) > 0 Then
            tRow = i: Exit For
        End If
    Next i
    
    For i = startColIdx To endColIdx
        If InStr(1, CStr(gWs.Cells(PREV_MONTH_HEADER_ROW, i).Value), caseTypeContext, vbTextCompare) > 0 Then
            tCol = i: Exit For
        End If
    Next i
    
    If tRow > 0 And tCol > 0 Then
        FindPreviousMonthIntersectCell = gWs.Cells(tRow, tCol).Address(RowAbsolute:=False, ColumnAbsolute:=False)
    End If
End Function

Private Function CountOpenCasesByPerson(ws As Worksheet, nameColLetter As String, dateColLetter As String, statusColLetter As String, targetMonth As Integer, targetYear As Integer) As Object
    Dim gWs As Worksheet, personDict As Object, countDict As Object
    Dim gLastRow As Long, lastRow As Long, i As Long
    Dim gName As String, personName As String, statusStr As String
    Dim nameColIdx As Long, dateColIdx As Long, statusColIdx As Long
    Dim namesArray As Variant, datesArray As Variant, statusesArray As Variant
    Dim rawDateVal As Variant, dDate As Date
    Dim isValidDate As Boolean, validStatuses() As String, st As Variant, statusMatched As Boolean
    
    Set personDict = CreateObject("Scripting.Dictionary")
    Set countDict = CreateObject("Scripting.Dictionary")
    validStatuses = Split(VALID_OPEN_STATUSES, ",")
    
    On Error Resume Next
    Set gWs = ThisWorkbook.Sheets(PREVIOUS_MONTH_SHEET_NAME)
    On Error GoTo 0
    
    If Not gWs Is Nothing Then
        gLastRow = gWs.Cells(gWs.Rows.Count, PREV_MONTH_NAME_COL).End(xlUp).Row
        If gLastRow >= 2 Then
            For i = 2 To gLastRow
                gName = Trim(CStr(gWs.Cells(i, PREV_MONTH_NAME_COL).Value))
                If gName <> "" Then
                    personDict(StrConv(gName, vbUpperCase)) = gName
                    countDict(gName) = 0
                End If
            Next i
        End If
    End If
    
    If personDict.Count = 0 Then
        Set CountOpenCasesByPerson = countDict
        Exit Function
    End If
    
    nameColIdx = ws.Columns(nameColLetter).Column
    dateColIdx = ws.Columns(dateColLetter).Column
    statusColIdx = ws.Columns(statusColLetter).Column
    
    lastRow = ws.Cells(ws.Rows.Count, dateColIdx).End(xlUp).Row
    If lastRow < 2 Then
        Set CountOpenCasesByPerson = countDict
        Exit Function
    End If
    
    namesArray = ws.Range(ws.Cells(2, nameColIdx), ws.Cells(lastRow, nameColIdx)).Value
    datesArray = ws.Range(ws.Cells(2, dateColIdx), ws.Cells(lastRow, dateColIdx)).Value
    statusesArray = ws.Range(ws.Cells(2, statusColIdx), ws.Cells(lastRow, statusColIdx)).Value
    
    For i = 1 To UBound(datesArray, 1)
        rawDateVal = datesArray(i, 1)
        isValidDate = False
        
        If Not IsEmpty(rawDateVal) Then
            If IsDate(rawDateVal) Then
                dDate = CDate(rawDateVal)
                isValidDate = True
            ElseIf IsNumeric(rawDateVal) Then
                On Error Resume Next
                dDate = CDate(rawDateVal)
                isValidDate = (Err.Number = 0)
                On Error GoTo 0
            End If
            
            If isValidDate Then
                If Month(dDate) = targetMonth And Year(dDate) = targetYear Then
                    statusStr = Trim(CStr(statusesArray(i, 1)))
                    statusMatched = False
                    
                    For Each st In validStatuses
                        If InStr(1, statusStr, Trim(CStr(st)), vbTextCompare) > 0 Then
                            statusMatched = True
                            Exit For
                        End If
                    Next st
                    
                    If statusMatched Then
                        personName = StrConv(Trim(CStr(namesArray(i, 1))), vbUpperCase)
                        If personDict.Exists(personName) Then
                            gName = personDict(personName)
                            countDict(gName) = countDict(gName) + 1
                        End If
                    End If
                End If
            End If
        End If
    Next i
    
    Set CountOpenCasesByPerson = countDict
End Function

Private Function FindPreviousMonthPersonIntersectCell(ByVal personName As String, ByVal caseTypeContext As String) As String
    Dim gWs As Worksheet
    Dim startColIdx As Long, endColIdx As Long
    Dim tRow As Long, tCol As Long, i As Long, gLastRow As Long
    
    FindPreviousMonthPersonIntersectCell = STATUS_NOT_FOUND
    On Error Resume Next
    Set gWs = ThisWorkbook.Sheets(PREVIOUS_MONTH_SHEET_NAME)
    On Error GoTo 0
    
    If gWs Is Nothing Then Exit Function
    
    startColIdx = gWs.Columns(PREV_MONTH_START_COL).Column
    endColIdx = gWs.Columns(PREV_MONTH_END_COL).Column
    
    gLastRow = gWs.Cells(gWs.Rows.Count, PREV_MONTH_NAME_COL).End(xlUp).Row
    For i = 2 To gLastRow
        If InStr(1, CStr(gWs.Cells(i, PREV_MONTH_NAME_COL).Value), personName, vbTextCompare) > 0 Then
            tRow = i: Exit For
        End If
    Next i
    
    For i = startColIdx To endColIdx
        If InStr(1, CStr(gWs.Cells(PREV_MONTH_HEADER_ROW, i).Value), caseTypeContext, vbTextCompare) > 0 Then
            tCol = i: Exit For
        End If
    Next i
    
    If tRow > 0 And tCol > 0 Then
        FindPreviousMonthPersonIntersectCell = gWs.Cells(tRow, tCol).Address(RowAbsolute:=False, ColumnAbsolute:=False)
    End If
End Function

Private Function IsCellAddressValid(ws As Worksheet, ByVal addr As String) As Boolean
    On Error Resume Next
    Dim testRng As Range
    Set testRng = ws.Range(addr)
    IsCellAddressValid = (Err.Number = 0 And Not testRng Is Nothing)
    Err.Clear
End Function