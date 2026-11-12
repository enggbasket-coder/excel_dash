Attribute VB_Name = "RFI_Duplicate_Audit"
Option Explicit

' ==============================================================
'  RFI Duplicate Audit (with cross-sheet detection)
'  - Scans every sheet in the active workbook
'  - Finds any column whose header contains "RFI" (case-insensitive)
'    anywhere in the first 20 rows (handles header rows up to 20)
'  - Detects in-sheet AND cross-sheet duplicates
'  - Writes audit result in a NEW column ("Duplicate Audit")
'    placed right after the last used column
'  - Skips blank RFI cells
'  - Exact text match (trimmed, case-insensitive)
'
'  Output values in the audit column:
'      Unique
'      Duplicate of row 7
'      Also in 'BC' row 5
'      Duplicate of row 7; Also in 'BC' row 5
' ==============================================================

' Module-level index of every RFI value across the whole workbook.
' Key   = normalized RFI string
' Value = pipe-separated list of "SheetName||Row" entries
Private g_index As Collection
Private Const SCAN_ROWS As Long = 20
Private Const SUMMARY_NAME As String = "RFI Audit Summary"

Public Sub RFI_dupli()
    Dim ws As Worksheet
    Dim sheetsScanned As Long, sheetsWithRFI As Long
    Dim totalInSheetDupes As Long, totalCrossSheetDupes As Long

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.DisplayAlerts = False

    ' Remove old summary sheet if it exists
    On Error Resume Next
    ActiveWorkbook.Worksheets(SUMMARY_NAME).Delete
    On Error GoTo 0
    Application.DisplayAlerts = True

    ' ---- PASS 1: build workbook-wide RFI index ----
    Set g_index = New Collection
    For Each ws In ActiveWorkbook.Worksheets
        If ws.Name <> SUMMARY_NAME Then
            BuildIndexForSheet ws
        End If
    Next ws

    ' ---- Create summary sheet ----
    Dim summary As Worksheet
    Set summary = ActiveWorkbook.Worksheets.Add(After:=ActiveWorkbook.Worksheets(ActiveWorkbook.Worksheets.Count))
    summary.Name = SUMMARY_NAME
    summary.Range("A1:E1").Value = Array("Sheet Name", "Has RFI Column?", "Header Row", "In-sheet Duplicates", "Cross-sheet Duplicates")
    summary.Range("A1:E1").Font.Bold = True
    summary.Range("A1:E1").Interior.Color = RGB(255, 242, 204)

    Dim outRow As Long: outRow = 2

    ' ---- PASS 2: write audit column on each sheet ----
    For Each ws In ActiveWorkbook.Worksheets
        If ws.Name <> SUMMARY_NAME Then
            sheetsScanned = sheetsScanned + 1
            Dim found As Boolean, inSheet As Long, crossSheet As Long, hdrRow As Long
            found = AuditSheet(ws, inSheet, crossSheet, hdrRow)
            summary.Cells(outRow, 1).Value = ws.Name
            If found Then
                sheetsWithRFI = sheetsWithRFI + 1
                totalInSheetDupes = totalInSheetDupes + inSheet
                totalCrossSheetDupes = totalCrossSheetDupes + crossSheet
                summary.Cells(outRow, 2).Value = "Yes"
                summary.Cells(outRow, 3).Value = hdrRow
                summary.Cells(outRow, 4).Value = inSheet
                summary.Cells(outRow, 5).Value = crossSheet
                If inSheet > 0 Then
                    summary.Cells(outRow, 4).Interior.Color = RGB(255, 199, 206)
                    summary.Cells(outRow, 4).Font.Color = RGB(156, 0, 6)
                End If
                If crossSheet > 0 Then
                    summary.Cells(outRow, 5).Interior.Color = RGB(255, 235, 156)
                    summary.Cells(outRow, 5).Font.Color = RGB(156, 87, 0)
                End If
            Else
                summary.Cells(outRow, 2).Value = "No"
                summary.Cells(outRow, 2).Interior.Color = RGB(217, 217, 217)
                summary.Cells(outRow, 3).Value = "-"
                summary.Cells(outRow, 4).Value = "-"
                summary.Cells(outRow, 5).Value = "-"
            End If
            outRow = outRow + 1
        End If
    Next ws

    summary.Columns("A:E").AutoFit

    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True

    MsgBox "RFI Duplicate Audit complete." & vbCrLf & vbCrLf & _
           "Sheets scanned: " & sheetsScanned & vbCrLf & _
           "Sheets with RFI column: " & sheetsWithRFI & vbCrLf & _
           "In-sheet duplicate rows: " & totalInSheetDupes & vbCrLf & _
           "Cross-sheet duplicate rows: " & totalCrossSheetDupes & vbCrLf & vbCrLf & _
           "See sheet '" & SUMMARY_NAME & "' for per-sheet details.", _
           vbInformation, "RFI Duplicate Audit"
End Sub

' --------------------------------------------------------------
' PASS 1: scan a sheet and add every RFI value to g_index.
' --------------------------------------------------------------
Private Sub BuildIndexForSheet(ws As Worksheet)
    Dim headerRow As Long, rfiCol As Long
    If Not FindRFIColumn(ws, headerRow, rfiCol) Then Exit Sub

    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, rfiCol).End(xlUp).Row
    If lastRow <= headerRow Then Exit Sub

    Dim r As Long, raw As String, key As String, entry As String
    Dim existing As Variant
    For r = headerRow + 1 To lastRow
        raw = CStr(ws.Cells(r, rfiCol).Value)
        key = LCase$(Trim$(raw))
        If Len(key) > 0 Then
            entry = ws.Name & "||" & r
            existing = CollectionGet(g_index, key)
            If IsEmpty(existing) Then
                g_index.Add entry, key
            Else
                ' Append to existing list
                g_index.Remove key
                g_index.Add existing & "@@" & entry, key
            End If
        End If
    Next r
End Sub

' --------------------------------------------------------------
' PASS 2: write the audit column for one sheet.
' Returns True if an RFI column was found.
' --------------------------------------------------------------
Private Function AuditSheet(ws As Worksheet, ByRef inSheetDupes As Long, _
                            ByRef crossSheetDupes As Long, ByRef headerRowOut As Long) As Boolean
    Dim headerRow As Long, rfiCol As Long
    inSheetDupes = 0
    crossSheetDupes = 0

    If Not FindRFIColumn(ws, headerRow, rfiCol) Then
        AuditSheet = False
        Exit Function
    End If
    AuditSheet = True
    headerRowOut = headerRow

    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, rfiCol).End(xlUp).Row
    If lastRow <= headerRow Then Exit Function

    Dim auditCol As Long
    auditCol = LastUsedColumn(ws) + 1

    ws.Cells(headerRow, auditCol).Value = "Duplicate Audit"
    ws.Cells(headerRow, auditCol).Font.Bold = True
    ws.Cells(headerRow, auditCol).Interior.Color = RGB(255, 242, 204)

    ' Track first-seen row within THIS sheet
    Dim seen As Collection
    Set seen = New Collection

    Dim r As Long, raw As String, key As String
    For r = headerRow + 1 To lastRow
        raw = CStr(ws.Cells(r, rfiCol).Value)
        key = LCase$(Trim$(raw))
        If Len(key) = 0 Then
            ws.Cells(r, auditCol).Value = ""
        Else
            Dim msgPart As String: msgPart = ""
            Dim isInSheetDup As Boolean: isInSheetDup = False
            Dim isCrossSheetDup As Boolean: isCrossSheetDup = False

            ' --- In-sheet check ---
            Dim firstRow As Variant
            firstRow = CollectionGet(seen, key)
            If IsEmpty(firstRow) Then
                seen.Add r, key
            Else
                msgPart = "Duplicate of row " & firstRow
                isInSheetDup = True
            End If

            ' --- Cross-sheet check ---
            Dim crossMsg As String
            crossMsg = BuildCrossSheetMessage(key, ws.Name, r)
            If Len(crossMsg) > 0 Then
                If Len(msgPart) > 0 Then
                    msgPart = msgPart & "; " & crossMsg
                Else
                    msgPart = crossMsg
                End If
                isCrossSheetDup = True
            End If

            ' --- Write result + colour ---
            If Len(msgPart) = 0 Then
                ws.Cells(r, auditCol).Value = "Unique"
                ws.Cells(r, auditCol).Interior.Color = RGB(198, 239, 206) ' green
                ws.Cells(r, auditCol).Font.Color = RGB(0, 97, 0)
            Else
                ws.Cells(r, auditCol).Value = msgPart
                If isInSheetDup Then
                    ws.Cells(r, auditCol).Interior.Color = RGB(255, 199, 206) ' red
                    ws.Cells(r, auditCol).Font.Color = RGB(156, 0, 6)
                Else
                    ws.Cells(r, auditCol).Interior.Color = RGB(255, 235, 156) ' amber
                    ws.Cells(r, auditCol).Font.Color = RGB(156, 87, 0)
                End If
            End If

            If isInSheetDup Then inSheetDupes = inSheetDupes + 1
            If isCrossSheetDup Then crossSheetDupes = crossSheetDupes + 1
        End If
    Next r

    ws.Columns(auditCol).AutoFit
End Function

' --------------------------------------------------------------
' For a given RFI key, build a string listing all OTHER sheets it
' appears on. Returns "" if it appears only on the current sheet.
' Format: "Also in 'BC' row 5" or "Also in 'BC' row 5, 'WMM' row 12"
' --------------------------------------------------------------
Private Function BuildCrossSheetMessage(key As String, currentSheet As String, currentRow As Long) As String
    Dim raw As Variant
    raw = CollectionGet(g_index, key)
    If IsEmpty(raw) Then Exit Function

    Dim entries() As String
    entries = Split(CStr(raw), "@@")

    Dim parts As String, i As Long, parts2() As String
    For i = LBound(entries) To UBound(entries)
        parts2 = Split(entries(i), "||")
        If UBound(parts2) >= 1 Then
            Dim shName As String, shRow As String
            shName = parts2(0)
            shRow = parts2(1)
            ' Skip entries on the current sheet
            If shName <> currentSheet Then
                If Len(parts) > 0 Then parts = parts & ", "
                parts = parts & "'" & shName & "' row " & shRow
            End If
        End If
    Next i

    If Len(parts) > 0 Then
        BuildCrossSheetMessage = "Also in " & parts
    End If
End Function

' --------------------------------------------------------------
' Look for a header containing "RFI" anywhere in the first
' SCAN_ROWS rows. Returns True and sets headerRow / colIndex.
' --------------------------------------------------------------
Private Function FindRFIColumn(ws As Worksheet, ByRef headerRow As Long, ByRef colIndex As Long) As Boolean
    Dim lastCol As Long: lastCol = LastUsedColumn(ws)
    If lastCol = 0 Then Exit Function

    Dim r As Long, c As Long, v As String
    For r = 1 To SCAN_ROWS
        For c = 1 To lastCol
            v = CStr(ws.Cells(r, c).Value)
            If Len(v) > 0 Then
                If InStr(1, v, "RFI", vbTextCompare) > 0 Then
                    headerRow = r
                    colIndex = c
                    FindRFIColumn = True
                    Exit Function
                End If
            End If
        Next c
    Next r
End Function

' --------------------------------------------------------------
' Safe Collection lookup. Returns Empty if key missing.
' --------------------------------------------------------------
Private Function CollectionGet(col As Collection, key As String) As Variant
    On Error Resume Next
    CollectionGet = col(key)
    If Err.Number <> 0 Then
        Err.Clear
        CollectionGet = Empty
    End If
    On Error GoTo 0
End Function

' --------------------------------------------------------------
' Last column that contains data anywhere on the sheet.
' --------------------------------------------------------------
Private Function LastUsedColumn(ws As Worksheet) As Long
    Dim lc As Range
    On Error Resume Next
    Set lc = ws.Cells.Find(What:="*", LookIn:=xlFormulas, _
                           SearchOrder:=xlByColumns, SearchDirection:=xlPrevious)
    On Error GoTo 0
    If lc Is Nothing Then
        LastUsedColumn = 0
    Else
        LastUsedColumn = lc.Column
    End If
End Function
