Attribute VB_Name = "StripChart"
Option Explicit

' =====================================================================
'  StripChart Plot  v2.0  —  Combined multi-layer strip chart
'
'  Design:
'    - One strip-chart sheet for ALL selected data sheets
'    - Each selected data sheet becomes ONE layer row
'    - Layer name = sheet name; Layer color = sheet tab color (or palette)
'    - Each row of a data sheet adds a coloured strip across the chainage
'      band, on the row's SIDE (RHS/LHS/BHS)
'
'  Data-sheet column layout (auto-detected start row):
'    A: SL.NO  (data starts at first numeric row in col A)
'    E: From chainage  (text "152+780" or numeric)
'    F: To   chainage  (text "152+800" or numeric)
'    G: SIDE  (LHS / RHS / BHS / BS / BOTH)
'
'  Config layout in "_SC_Cfg" (written by Streamlit dashboard):
'    B1: from chainage (metres)
'    B2: to chainage   (metres)
'    B3: interval      (metres per column)
'    B4: assume_km     (1 = numeric data in km, 0 = metres)
'    B5..B11: slot 1..7 data-sheet name (= layer name)
'    C5      : combined strip-sheet name (slots 2..7's C cells are ignored)
'    Z column: status log written by macro
' =====================================================================

Private Const COLS_PER_BLOCK As Long = 50
Private Const C_DATA_START   As Long = 3      ' column C
Private Const CFG_SHEET      As String = "_SC_Cfg"
Private Const COLOR_RED      As Long = 255      ' RGB(255,0,0) — overlap cells
Private Const COLOR_PURPLE   As Long = 8388736  ' RGB(128,0,128) — missing-lower cells
Private Const COLOR_OK_GREEN As Long = 32768    ' RGB(0,128,0)

Private g_cfg As Worksheet
Private g_logRow As Long

' Tracking grids for audit. Indexed (1 To N, 0 To totalCols-1).
' g_paintedXHS = a row of any layer painted this cell
' g_overlapXHS = the cell was painted twice by the same layer (overlap)
Private g_paintedRHS() As Boolean
Private g_paintedLHS() As Boolean
Private g_overlapRHS() As Boolean
Private g_overlapLHS() As Boolean
Private g_totalCols As Long


' Parameterless wrapper — xlwings on Mac can't reliably pass arguments via
' AppleScript "run VB macro", so the dashboard activates the target book first
' and we always read ActiveWorkbook here.
Public Sub StripChart_Run()
    StripChart_RunImpl ""
End Sub

' The real entry point. Can be called directly from VBA with a target name.
Public Sub StripChart_RunImpl(Optional targetBookName As String = "")
    On Error GoTo Fail

    Dim wb As Workbook
    If Len(targetBookName) > 0 Then
        On Error Resume Next
        Set wb = Workbooks(targetBookName)
        On Error GoTo 0
        If wb Is Nothing Then
            MsgBox "Target workbook '" & targetBookName & "' not open in Excel.", vbExclamation
            Exit Sub
        End If
    Else
        ' Find the data workbook by scanning every open book for _SC_Cfg.
        ' This works regardless of what the macro workbook is named (so
        ' "StripChart_macros.xlsm", "StripChart_macros V2.xlsm", a copy with
        ' any name — all behave the same). Falls back to ActiveWorkbook
        ' if nothing has _SC_Cfg.
        Set wb = Nothing
        Dim wbCand As Workbook
        For Each wbCand In Application.Workbooks
            If Not (wbCand Is ThisWorkbook) Then
                Dim shCand As Worksheet
                For Each shCand In wbCand.Sheets
                    If LCase(shCand.Name) = LCase(CFG_SHEET) Then
                        Set wb = wbCand
                        Exit For
                    End If
                Next shCand
                If Not wb Is Nothing Then Exit For
            End If
        Next wbCand
        If wb Is Nothing Then
            Set wb = ActiveWorkbook
            If wb Is Nothing Then MsgBox "No active workbook.", vbExclamation: Exit Sub
        End If
    End If

    On Error Resume Next
    Set g_cfg = wb.Sheets(CFG_SHEET)
    On Error GoTo 0
    If g_cfg Is Nothing Then
        MsgBox "Config sheet '" & CFG_SHEET & "' not found in '" & wb.Name & "'.", vbExclamation
        Exit Sub
    End If

    g_cfg.Columns("Z").Clear
    g_logRow = 0
    ' Make sure cfg is visible so Mac VBA writes are not silently dropped
    On Error Resume Next
    g_cfg.Visible = xlSheetVisible
    On Error GoTo 0
    Log_ "START | Workbook=" & wb.Name & " | " & Now

    Dim fromM As Long, toM As Long, intervalM As Long
    Dim assumeKm As Boolean
    fromM     = CLng(g_cfg.Range("B1").Value)
    toM       = CLng(g_cfg.Range("B2").Value)
    intervalM = CLng(g_cfg.Range("B3").Value)
    assumeKm  = (CStr(g_cfg.Range("B4").Value) = "1")

    Log_ "Range: " & fromM & " to " & toM & " | interval=" & intervalM & " | assumeKm=" & assumeKm

    If toM <= fromM Then Log_ "ABORT: To must be > From": Exit Sub
    If intervalM <= 0 Then Log_ "ABORT: Interval must be > 0": Exit Sub

    ' ---- Read slot rows → build layer list ----
    Dim layerSheets As Collection: Set layerSheets = New Collection
    Dim layerNames  As Collection: Set layerNames = New Collection
    Dim layerColors As Collection: Set layerColors = New Collection

    Dim slot As Long
    For slot = 5 To 11
        Dim dsName As String
        dsName = Trim(CStr(g_cfg.Cells(slot, "B").Value))
        If Len(dsName) = 0 Then GoTo NextSlot

        Dim wsData As Worksheet
        Set wsData = FindSheet(wb, dsName)
        If wsData Is Nothing Then
            Log_ "SKIP slot " & (slot - 4) & ": data sheet '" & dsName & "' not found"
            GoTo NextSlot
        End If

        layerSheets.Add wsData
        layerNames.Add dsName
        layerColors.Add LayerColorFor(wsData, layerNames.Count)
        Log_ "Layer " & layerNames.Count & ": " & dsName & " (color=" & layerColors(layerNames.Count) & ")"
NextSlot:
    Next slot

    If layerNames.Count = 0 Then
        Log_ "ABORT: no valid layers configured"
        Exit Sub
    End If

    ' ---- Decide combined strip-sheet name ----
    Dim stripName As String
    stripName = Trim(CStr(g_cfg.Cells(5, "C").Value))
    If Len(stripName) = 0 Then stripName = "StripChart"

    ' Validate: must not collide with any data-sheet name (would overwrite data)
    Dim ic As Long
    For ic = 1 To layerNames.Count
        If LCase(stripName) = LCase(layerNames(ic)) Then
            Log_ "ABORT: strip sheet name '" & stripName & "' matches data sheet '" & layerNames(ic) & "' — pick a different name"
            Exit Sub
        End If
    Next ic
    Log_ "Strip sheet: " & stripName

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    ' ---- Recreate strip sheet ----
    Dim wsStrip As Worksheet
    On Error Resume Next
    Set wsStrip = wb.Sheets(stripName)
    On Error GoTo Fail
    If Not wsStrip Is Nothing Then
        wsStrip.Cells.Clear
        wsStrip.Cells.UnMerge
    Else
        Set wsStrip = wb.Sheets.Add(After:=wb.Sheets(wb.Sheets.Count))
        wsStrip.Name = stripName
    End If

    Dim N As Long: N = layerNames.Count
    Dim BLOCK_ROWS As Long: BLOCK_ROWS = 1 + N + 1 + N + 1
    Dim metresPerBlock As Long: metresPerBlock = COLS_PER_BLOCK * intervalM
    Dim numBlocks As Long
    numBlocks = Int((toM - fromM + metresPerBlock - 1) / metresPerBlock)
    Dim lastCol As Long: lastCol = C_DATA_START + COLS_PER_BLOCK - 1

    wsStrip.Columns("A").ColumnWidth = 5
    wsStrip.Columns("B").ColumnWidth = 11
    Dim widCol As Long
    For widCol = C_DATA_START To lastCol
        wsStrip.Columns(widCol).ColumnWidth = 3.2
    Next widCol

    Dim b As Long
    For b = 0 To numBlocks - 1
        Dim blockTop As Long: blockTop = b * BLOCK_ROWS + 1
        Dim blockStartM As Long: blockStartM = fromM + b * metresPerBlock
        BuildOneBlock wsStrip, blockTop, blockStartM, intervalM, lastCol, layerNames, layerColors
    Next b
    Log_ "Built " & numBlocks & " block(s), " & N & " layer(s) per block"

    ' Initialise audit-tracking grids — every cell starts unpainted
    g_totalCols = numBlocks * COLS_PER_BLOCK
    ReDim g_paintedRHS(1 To N, 0 To g_totalCols - 1)
    ReDim g_paintedLHS(1 To N, 0 To g_totalCols - 1)
    ReDim g_overlapRHS(1 To N, 0 To g_totalCols - 1)
    ReDim g_overlapLHS(1 To N, 0 To g_totalCols - 1)

    ' ---- Paint strips for each layer ----
    Dim totalPainted As Long: totalPainted = 0
    Dim i As Long
    For i = 1 To N
        Dim painted As Long, attempted As Long, skipped As Long
        painted = 0: attempted = 0: skipped = 0
        On Error Resume Next
        painted = PaintLayer(layerSheets(i), wsStrip, i, N, fromM, toM, intervalM, _
                             metresPerBlock, BLOCK_ROWS, lastCol, layerColors(i), assumeKm, _
                             numBlocks, attempted, skipped)
        If Err.Number <> 0 Then
            Log_ "Layer " & i & " [" & layerNames(i) & "]: ERROR " & Err.Number & " - " & Err.Description
            Err.Clear
        Else
            Log_ "Layer " & i & " [" & layerNames(i) & "]: rows_seen=" & attempted & _
                 " skipped=" & skipped & " cells_painted=" & painted
            totalPainted = totalPainted + painted
        End If
        On Error GoTo Fail
    Next i

    ' ---- Audit pass ----
    Dim auditCount As Long: auditCount = 0
    On Error Resume Next
    auditCount = RunAudit(wsStrip, layerNames, fromM, intervalM, lastCol)
    If Err.Number <> 0 Then
        Log_ "AUDIT ERROR: " & Err.Number & " - " & Err.Description
        Err.Clear
    End If
    On Error GoTo Fail

    Application.DisplayAlerts = True
    Application.ScreenUpdating = True

    Log_ "DONE | layers=" & N & " | blocks=" & numBlocks & _
         " | total_painted=" & totalPainted & " | audit_issues=" & auditCount
    Exit Sub

Fail:
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    Log_ "FATAL: " & Err.Description
End Sub


' ---------------------------------------------------------------------
'  Logging — appends one line to col Z of _SC_Cfg
' ---------------------------------------------------------------------
Private Sub Log_(msg As String)
    If g_cfg Is Nothing Then Exit Sub
    g_logRow = g_logRow + 1
    g_cfg.Cells(g_logRow, "Z").Value = msg
End Sub


' ---------------------------------------------------------------------
'  Find a worksheet case-insensitively
' ---------------------------------------------------------------------
Private Function FindSheet(wb As Workbook, name As String) As Worksheet
    Dim sh As Worksheet
    Dim target As String: target = LCase(Trim(name))
    For Each sh In wb.Worksheets
        If LCase(Trim(sh.Name)) = target Then
            Set FindSheet = sh
            Exit Function
        End If
    Next sh
End Function


' ---------------------------------------------------------------------
'  Pick a colour for a layer:
'    1. Sheet tab colour, if set
'    2. Cell A1 background, if not white
'    3. Distinct palette colour by index
' ---------------------------------------------------------------------
Private Function LayerColorFor(ws As Worksheet, idx As Long) As Long
    On Error Resume Next
    Dim c As Variant

    c = Empty
    c = ws.Tab.Color
    If Err.Number = 0 And IsNumeric(c) Then
        Dim cl As Long: cl = CLng(c)
        If cl > 0 And cl <> 16777215 Then
            LayerColorFor = cl
            On Error GoTo 0
            Exit Function
        End If
    End If
    Err.Clear

    c = Empty
    c = ws.Cells(1, 1).Interior.Color
    If Err.Number = 0 And IsNumeric(c) Then
        Dim cl2 As Long: cl2 = CLng(c)
        If cl2 > 0 And cl2 <> 16777215 Then
            LayerColorFor = cl2
            On Error GoTo 0
            Exit Function
        End If
    End If
    On Error GoTo 0

    LayerColorFor = PaletteColor(idx)
End Function


Private Function PaletteColor(idx As Long) As Long
    Select Case ((idx - 1) Mod 7) + 1
        Case 1: PaletteColor = RGB(212, 165, 116)   ' tan
        Case 2: PaletteColor = RGB(168, 199, 122)   ' light green
        Case 3: PaletteColor = RGB(122, 168, 199)   ' sky blue
        Case 4: PaletteColor = RGB(199, 122, 168)   ' pink
        Case 5: PaletteColor = RGB(255, 193, 7)     ' gold
        Case 6: PaletteColor = RGB(122, 199, 199)   ' light teal
        Case 7: PaletteColor = RGB(199, 122, 122)   ' salmon
    End Select
End Function


' ---------------------------------------------------------------------
'  Layout one chainage block (header row + RHS rows + axis + LHS rows)
' ---------------------------------------------------------------------
Private Sub BuildOneBlock(ws As Worksheet, blockTop As Long, startM As Long, _
                          intervalM As Long, lastCol As Long, _
                          layerNames As Collection, layerColors As Collection)
    Dim N As Long: N = layerNames.Count
    Dim hdr As Long: hdr = blockTop
    Dim rhsTop As Long: rhsTop = blockTop + 1
    Dim rhsBot As Long: rhsBot = blockTop + N
    Dim axis As Long: axis = blockTop + N + 1
    Dim lhsTop As Long: lhsTop = blockTop + N + 2
    Dim lhsBot As Long: lhsBot = blockTop + 2 * N + 1

    ws.Rows(hdr).RowHeight = 22
    Dim rh As Long
    For rh = rhsTop To lhsBot
        ws.Rows(rh).RowHeight = IIf(rh = axis, 64, 18)
    Next rh

    ws.Cells(hdr, "A").Value = "SIDE"
    ws.Cells(hdr, "B").Value = "LAYER"
    With ws.Range(ws.Cells(hdr, C_DATA_START), ws.Cells(hdr, lastCol))
        .Merge: .Value = "CHAINAGE"
        .HorizontalAlignment = xlCenter: .VerticalAlignment = xlCenter
        .Font.Bold = True: .Font.Size = 14
    End With
    With ws.Range(ws.Cells(hdr, "A"), ws.Cells(hdr, "B"))
        .Font.Bold = True
        .HorizontalAlignment = xlCenter: .VerticalAlignment = xlCenter
    End With

    With ws.Range(ws.Cells(rhsTop, "A"), ws.Cells(rhsBot, "A"))
        .Merge: .Value = "RHS"
        .Orientation = 90: .Font.Bold = True
        .HorizontalAlignment = xlCenter: .VerticalAlignment = xlCenter
    End With
    With ws.Range(ws.Cells(lhsTop, "A"), ws.Cells(lhsBot, "A"))
        .Merge: .Value = "LHS"
        .Orientation = 90: .Font.Bold = True
        .HorizontalAlignment = xlCenter: .VerticalAlignment = xlCenter
    End With

    Dim endM As Long: endM = startM + COLS_PER_BLOCK * intervalM
    With ws.Cells(axis, "A")
        .Value = FormatChainage(startM) & " to " & FormatChainage(endM)
        .Orientation = 90: .Font.Bold = True: .Font.Size = 9
        .HorizontalAlignment = xlCenter: .VerticalAlignment = xlCenter
    End With

    ' Layer order: first-selected layer is closest to the chainage axis
    ' (just-above-axis on RHS, just-below-axis on LHS). Last-selected layer
    ' is at the outer edges (RHS top, LHS bottom).
    Dim i As Long
    For i = 1 To N
        With ws.Cells(rhsTop + i - 1, "B")
            .Value = layerNames(N - i + 1)
            .Interior.Color = layerColors(N - i + 1)
            .Font.Bold = True
            .HorizontalAlignment = xlCenter: .VerticalAlignment = xlCenter
        End With
    Next i
    For i = 1 To N
        With ws.Cells(lhsTop + i - 1, "B")
            .Value = layerNames(i)
            .Interior.Color = layerColors(i)
            .Font.Bold = True
            .HorizontalAlignment = xlCenter: .VerticalAlignment = xlCenter
        End With
    Next i

    Dim col As Long, m As Long
    For col = 0 To COLS_PER_BLOCK - 1
        m = startM + col * intervalM
        With ws.Cells(axis, C_DATA_START + col)
            .Value = FormatChainage(m)
            .Orientation = 90
            .HorizontalAlignment = xlCenter: .VerticalAlignment = xlCenter
            .Font.Size = 8
        End With
    Next col

    With ws.Range(ws.Cells(hdr, "A"), ws.Cells(lhsBot, lastCol))
        .Borders.LineStyle = xlContinuous
        .Borders.Weight = xlThin
        .Borders.Color = RGB(170, 170, 170)
    End With
End Sub


' ---------------------------------------------------------------------
'  Paint one layer's strips. Reads From/To from cols E,F and SIDE from G.
'  Auto-detects data start row (first numeric row in col A).
' ---------------------------------------------------------------------
Private Function PaintLayer(wsData As Worksheet, wsStrip As Worksheet, _
                             layerIdx As Long, N As Long, _
                             fromM As Long, toM As Long, intervalM As Long, _
                             metresPerBlock As Long, blockRows As Long, _
                             lastCol As Long, layerColor As Long, _
                             assumeKm As Boolean, numBlocks As Long, _
                             ByRef attempted As Long, ByRef skipped As Long) As Long
    Dim painted As Long: painted = 0
    attempted = 0: skipped = 0

    Dim startRow As Long: startRow = 0
    Dim r As Long
    For r = 1 To 30
        Dim aText As String
        aText = SafeCellText(wsData, r, 1)
        If Len(aText) > 0 And IsNumeric(aText) Then
            startRow = r
            Exit For
        End If
    Next r
    If startRow = 0 Then
        Log_ "  [" & wsData.Name & "] no numeric SL.NO found in first 30 rows of col A — sheet skipped"
        skipped = 0
        PaintLayer = 0
        Exit Function
    End If

    ' Auto-detect From / To / Side columns from header rows
    Dim fromCol As Long, toCol As Long, sideCol As Long
    FindHeaderColumns wsData, fromCol, toCol, sideCol

    Log_ "  [" & wsData.Name & "] startRow=" & startRow & _
         " | cols: From=" & ColLetter(fromCol) & _
         " To=" & ColLetter(toCol) & _
         " Side=" & ColLetter(sideCol) & _
         " | sample: " & ColLetter(fromCol) & startRow & "=" & SafeCellText(wsData, startRow, fromCol) & _
         ", " & ColLetter(toCol) & startRow & "=" & SafeCellText(wsData, startRow, toCol) & _
         ", " & ColLetter(sideCol) & startRow & "=" & SafeCellText(wsData, startRow, sideCol)

    Dim lastRow As Long
    On Error Resume Next
    lastRow = wsData.Cells(wsData.Rows.Count, "A").End(xlUp).Row
    If Err.Number <> 0 Then lastRow = 0: Err.Clear
    On Error GoTo 0
    If lastRow < startRow Then PaintLayer = 0: Exit Function

    For r = startRow To lastRow
        attempted = attempted + 1

        Dim fromTxt As String, toTxt As String
        fromTxt = SafeCellText(wsData, r, fromCol)
        toTxt = SafeCellText(wsData, r, toCol)

        Dim fM As Long, tM As Long
        fM = ParseChainage(fromTxt, assumeKm)
        tM = ParseChainage(toTxt, assumeKm)
        If fM = -1 Or tM = -1 Then skipped = skipped + 1: GoTo NextRow

        If tM < fM Then
            Dim sw As Long: sw = fM: fM = tM: tM = sw
        End If

        Dim sideStr As String
        sideStr = UCase(Trim(SafeCellText(wsData, r, sideCol)))
        If sideStr <> "RHS" And sideStr <> "LHS" And _
           sideStr <> "BHS" And sideStr <> "BS" And sideStr <> "BOTH" Then
            skipped = skipped + 1
            GoTo NextRow
        End If

        ' Paint by column-overlap, not by stride. Each grid column at index
        ' colIdx represents bin [fromM + colIdx*intervalM, fromM + (colIdx+1)*intervalM).
        ' Data ranges are interpreted half-open [fM, tM) — so adjacent rows
        ' that meet at the same chainage (row1: 105+700→106+000, row2:
        ' 106+000→106+300) do NOT report as overlaps. This matches how
        ' "length" in the source data is computed: tM - fM.
        Dim totalCols As Long: totalCols = numBlocks * COLS_PER_BLOCK
        If tM <= fromM Or fM >= toM Then GoTo NextRow   ' wholly out of range
        If tM <= fM Then GoTo NextRow                   ' degenerate / zero-length

        Dim firstCol As Long: firstCol = (fM - fromM) \ intervalM
        Dim lastColIdx As Long: lastColIdx = (tM - fromM - 1) \ intervalM
        If firstCol < 0 Then firstCol = 0
        If lastColIdx >= totalCols Then lastColIdx = totalCols - 1
        If lastColIdx < firstCol Then GoTo NextRow

        Dim isRHS As Boolean: isRHS = (sideStr = "RHS" Or sideStr = "BHS" Or sideStr = "BS" Or sideStr = "BOTH")
        Dim isLHS As Boolean: isLHS = (sideStr = "LHS" Or sideStr = "BHS" Or sideStr = "BS" Or sideStr = "BOTH")

        Dim colIdx As Long
        For colIdx = firstCol To lastColIdx
            Dim blockNo As Long: blockNo = colIdx \ COLS_PER_BLOCK
            Dim colInBlock As Long: colInBlock = colIdx Mod COLS_PER_BLOCK
            Dim blockTop As Long: blockTop = blockNo * blockRows + 1
            Dim col As Long: col = C_DATA_START + colInBlock

            ' RHS: first-selected layer just-above-axis, last-selected at top.
            If isRHS Then
                If g_paintedRHS(layerIdx, colIdx) Then
                    ' Same layer + side already painted this column — overlap
                    wsStrip.Cells(blockTop + (N - layerIdx + 1), col).Interior.Color = COLOR_RED
                    g_overlapRHS(layerIdx, colIdx) = True
                Else
                    wsStrip.Cells(blockTop + (N - layerIdx + 1), col).Interior.Color = layerColor
                    g_paintedRHS(layerIdx, colIdx) = True
                End If
                painted = painted + 1
            End If

            ' LHS: first-selected layer just-below-axis, last-selected at bottom.
            If isLHS Then
                If g_paintedLHS(layerIdx, colIdx) Then
                    wsStrip.Cells(blockTop + N + 1 + layerIdx, col).Interior.Color = COLOR_RED
                    g_overlapLHS(layerIdx, colIdx) = True
                Else
                    wsStrip.Cells(blockTop + N + 1 + layerIdx, col).Interior.Color = layerColor
                    g_paintedLHS(layerIdx, colIdx) = True
                End If
                painted = painted + 1
            End If
        Next colIdx
NextRow:
    Next r
    PaintLayer = painted
End Function


' ---------------------------------------------------------------------
'  Parse a chainage string to metres. Accepts:
'    "152+780"  → 152780
'    "152.78"   → 152780
'    "152780"   → 152780  (or 152780000 if assumeKm=True)
'  Returns -1 on parse failure.
' ---------------------------------------------------------------------
Private Function ParseChainage(s As String, assumeKm As Boolean) As Long
    Dim t As String: t = Trim(s)
    If Len(t) = 0 Then ParseChainage = -1: Exit Function

    Dim plusPos As Long: plusPos = InStr(t, "+")
    If plusPos > 0 Then
        Dim kmStr As String, mStr As String
        kmStr = Trim(Mid(t, 1, plusPos - 1))
        mStr  = Trim(Mid(t, plusPos + 1))
        If IsNumeric(kmStr) And IsNumeric(mStr) Then
            ParseChainage = CLng(kmStr) * 1000 + CLng(mStr)
        Else
            ParseChainage = -1
        End If
        Exit Function
    End If

    Dim dotPos As Long: dotPos = InStr(t, ".")
    If dotPos > 0 Then
        Dim kmStr2 As String, fracStr As String
        kmStr2 = Trim(Mid(t, 1, dotPos - 1))
        fracStr = Trim(Mid(t, dotPos + 1))
        If Len(fracStr) < 3 Then fracStr = fracStr & String(3 - Len(fracStr), "0")
        If Len(fracStr) > 3 Then fracStr = Left(fracStr, 3)
        If IsNumeric(kmStr2) And IsNumeric(fracStr) Then
            ParseChainage = CLng(kmStr2) * 1000 + CLng(fracStr)
        Else
            ParseChainage = -1
        End If
        Exit Function
    End If

    If IsNumeric(t) Then
        Dim n As Double: n = CDbl(t)
        ' Smart units: values < 1000 are clearly km (e.g. 152.78 or 152);
        ' values >= 1000 are already metres (e.g. 152780). The assumeKm
        ' flag only applies in the ambiguous low-magnitude case.
        If n < 1000 And assumeKm Then
            ParseChainage = CLng(n * 1000)
        Else
            ParseChainage = CLng(n)
        End If
        Exit Function
    End If

    ParseChainage = -1
End Function


Private Function FormatChainage(metres As Long) As String
    Dim km As Long, m As Long
    km = metres \ 1000
    m  = metres Mod 1000
    FormatChainage = km & "+" & Format(m, "000")
End Function


' ---------------------------------------------------------------------
'  Scan first 30 rows × 26 cols of a data sheet for headers labelled
'  "From", "To", "Side" (case-insensitive). Falls back to E/F/G if not
'  found, matching the SPS-12 layout we developed against.
' ---------------------------------------------------------------------
Private Sub FindHeaderColumns(wsData As Worksheet, _
                              ByRef fromCol As Long, _
                              ByRef toCol As Long, _
                              ByRef sideCol As Long)
    fromCol = 5   ' default E
    toCol = 6     ' default F
    sideCol = 7   ' default G

    Dim r As Long, c As Long
    Dim foundFrom As Boolean, foundTo As Boolean, foundSide As Boolean

    For r = 1 To 30
        For c = 1 To 26
            Dim t As String
            t = SafeCellText(wsData, r, c)
            If Len(t) = 0 Then GoTo NextCell
            t = LCase(Trim(t))
            If Not foundFrom Then
                If t = "from" Or t = "from chainage" Or t = "ch from" Or t = "from ch" Then
                    fromCol = c
                    foundFrom = True
                End If
            End If
            If Not foundTo Then
                If t = "to" Or t = "to chainage" Or t = "ch to" Or t = "to ch" Then
                    toCol = c
                    foundTo = True
                End If
            End If
            If Not foundSide Then
                If t = "side" Or t = "lhs/rhs" Or t = "rhs/lhs" Then
                    sideCol = c
                    foundSide = True
                End If
            End If
NextCell:
        Next c
        If foundFrom And foundTo And foundSide Then Exit For
    Next r
End Sub


' Read a cell as text, tolerating errors, merged-range quirks, and non-string values.
Private Function SafeCellText(ws As Worksheet, r As Long, c As Long) As String
    On Error Resume Next
    SafeCellText = ""
    Dim v As Variant
    v = ws.Cells(r, c).Value
    If Err.Number = 0 Then
        If Not IsError(v) Then
            If Not IsNull(v) And Not IsEmpty(v) Then
                SafeCellText = CStr(v)
            End If
        End If
    End If
    Err.Clear
    On Error GoTo 0
End Function


' Convert column number (1=A, 2=B, …, 27=AA, …) to letter.
' ByVal is critical: without it, the loop mutates the caller's variable to 0
' because VBA passes by reference by default.
Private Function ColLetter(ByVal c As Long) As String
    Dim out As String
    Do While c > 0
        Dim rem_ As Long
        rem_ = (c - 1) Mod 26
        out = Chr(65 + rem_) & out
        c = (c - 1) \ 26
    Loop
    ColLetter = out
End Function


' =====================================================================
'  AUDIT PASS — runs after all painting is done
'    1. Lists every overlap range (per layer + side)
'    2. Lists every chainage range where an upper layer is painted but
'       the layer immediately below is missing
'  Writes results in column (lastCol + 2) — one blank column gap from
'  the chart. Header in row 1, issues from row 2, OK in green if clean.
'  Returns the number of issues written.
' =====================================================================
Private Function RunAudit(wsStrip As Worksheet, layerNames As Collection, _
                          ByVal fromM As Long, ByVal intervalM As Long, _
                          ByVal lastCol As Long) As Long
    Dim auditCol As Long: auditCol = lastCol + 2
    Dim N As Long: N = layerNames.Count
    Dim BLOCK_ROWS As Long: BLOCK_ROWS = 1 + N + 1 + N + 1
    Dim numBlocks As Long: numBlocks = g_totalCols \ COLS_PER_BLOCK

    wsStrip.Columns(auditCol).ColumnWidth = 70

    Dim issueCount As Long: issueCount = 0

    ' --- For every block × every layer × every side, write a per-row audit ---
    Dim b As Long
    For b = 0 To numBlocks - 1
        Dim blockStartCol As Long: blockStartCol = b * COLS_PER_BLOCK
        Dim blockEndCol As Long: blockEndCol = blockStartCol + COLS_PER_BLOCK - 1
        Dim blockTop As Long: blockTop = b * BLOCK_ROWS + 1

        Dim i As Long
        For i = 1 To N
            Dim rhsRow As Long: rhsRow = blockTop + (N - i + 1)
            Dim lhsRow As Long: lhsRow = blockTop + N + 1 + i

            Dim rhsMsg As String, lhsMsg As String
            rhsMsg = BuildRowAudit(g_overlapRHS, g_paintedRHS, i, N, _
                                    blockStartCol, blockEndCol, fromM, intervalM, layerNames)
            lhsMsg = BuildRowAudit(g_overlapLHS, g_paintedLHS, i, N, _
                                    blockStartCol, blockEndCol, fromM, intervalM, layerNames)

            ' Paint missing-lower cells purple in this block (RHS)
            If i >= 2 Then
                PaintMissingInBlock wsStrip, g_paintedRHS, i, N, BLOCK_ROWS, True, _
                                    blockStartCol, blockEndCol, COLOR_PURPLE
                PaintMissingInBlock wsStrip, g_paintedLHS, i, N, BLOCK_ROWS, False, _
                                    blockStartCol, blockEndCol, COLOR_PURPLE
            End If

            If Len(rhsMsg) > 0 Then
                WriteAudit wsStrip, rhsRow, auditCol, rhsMsg, True
                issueCount = issueCount + 1
            ElseIf HasAnyPaintInBlock(g_paintedRHS, i, blockStartCol, blockEndCol) Then
                WriteAudit wsStrip, rhsRow, auditCol, "OK", False
            End If

            If Len(lhsMsg) > 0 Then
                WriteAudit wsStrip, lhsRow, auditCol, lhsMsg, True
                issueCount = issueCount + 1
            ElseIf HasAnyPaintInBlock(g_paintedLHS, i, blockStartCol, blockEndCol) Then
                WriteAudit wsStrip, lhsRow, auditCol, "OK", False
            End If
        Next i
    Next b

    RunAudit = issueCount
End Function


' Build the audit text for one (layer, side) row within one block:
'   * Overlap range(s) for this layer + side, if any
'   * Missing-lower range(s) for upper layers, if any
' Returns "" if no issues.
Private Function BuildRowAudit(overlapArr() As Boolean, paintedArr() As Boolean, _
                                ByVal layerIdx As Long, ByVal N As Long, _
                                ByVal startCol As Long, ByVal endCol As Long, _
                                ByVal fromM As Long, ByVal intervalM As Long, _
                                layerNames As Collection) As String
    Dim msg As String: msg = ""

    Dim ov As String
    ov = RangesInBlock(overlapArr, layerIdx, startCol, endCol, fromM, intervalM)
    If Len(ov) > 0 Then
        msg = "Overlap: " & ov
    End If

    If layerIdx >= 2 Then
        Dim ml As String
        ml = MissingLowerInBlock(paintedArr, layerIdx, startCol, endCol, fromM, intervalM)
        If Len(ml) > 0 Then
            If Len(msg) > 0 Then msg = msg & " | "
            msg = msg & "Missing lower '" & layerNames(layerIdx - 1) & "' at " & ml
        End If
    End If

    BuildRowAudit = msg
End Function


' Find consecutive True ranges in row `layerIdx` of a 2D Boolean array
' restricted to columns [startCol, endCol]. Returns chainage list.
Private Function RangesInBlock(arr() As Boolean, ByVal layerIdx As Long, _
                                ByVal startCol As Long, ByVal endCol As Long, _
                                ByVal fromM As Long, ByVal intervalM As Long) As String
    Dim out As String: out = ""
    Dim inRange As Boolean: inRange = False
    Dim startC As Long
    Dim c As Long

    If g_totalCols <= 0 Then Exit Function
    If endCol >= g_totalCols Then endCol = g_totalCols - 1

    For c = startCol To endCol
        If arr(layerIdx, c) Then
            If Not inRange Then
                startC = c
                inRange = True
            End If
        Else
            If inRange Then
                If Len(out) > 0 Then out = out & "; "
                out = out & "Km " & FormatChainage(fromM + startC * intervalM) & _
                      " to Km " & FormatChainage(fromM + c * intervalM)
                inRange = False
            End If
        End If
    Next c
    If inRange Then
        If Len(out) > 0 Then out = out & "; "
        out = out & "Km " & FormatChainage(fromM + startC * intervalM) & _
              " to Km " & FormatChainage(fromM + (endCol + 1) * intervalM)
    End If
    RangesInBlock = out
End Function


' Like RangesInBlock but for cells where layer upperIdx is painted yet (upperIdx-1) is not.
Private Function MissingLowerInBlock(arr() As Boolean, ByVal upperIdx As Long, _
                                      ByVal startCol As Long, ByVal endCol As Long, _
                                      ByVal fromM As Long, ByVal intervalM As Long) As String
    Dim out As String: out = ""
    Dim inRange As Boolean: inRange = False
    Dim startC As Long
    Dim c As Long

    If upperIdx < 2 Then Exit Function
    If g_totalCols <= 0 Then Exit Function
    If endCol >= g_totalCols Then endCol = g_totalCols - 1

    For c = startCol To endCol
        Dim isIssue As Boolean
        isIssue = arr(upperIdx, c) And Not arr(upperIdx - 1, c)
        If isIssue Then
            If Not inRange Then
                startC = c
                inRange = True
            End If
        Else
            If inRange Then
                If Len(out) > 0 Then out = out & "; "
                out = out & "Km " & FormatChainage(fromM + startC * intervalM) & _
                      " to Km " & FormatChainage(fromM + c * intervalM)
                inRange = False
            End If
        End If
    Next c
    If inRange Then
        If Len(out) > 0 Then out = out & "; "
        out = out & "Km " & FormatChainage(fromM + startC * intervalM) & _
              " to Km " & FormatChainage(fromM + (endCol + 1) * intervalM)
    End If
    MissingLowerInBlock = out
End Function


Private Function HasAnyPaintInBlock(arr() As Boolean, ByVal layerIdx As Long, _
                                     ByVal startCol As Long, ByVal endCol As Long) As Boolean
    If g_totalCols <= 0 Then Exit Function
    If endCol >= g_totalCols Then endCol = g_totalCols - 1
    Dim c As Long
    For c = startCol To endCol
        If arr(layerIdx, c) Then
            HasAnyPaintInBlock = True
            Exit Function
        End If
    Next c
End Function


' Override fill in cells where upperIdx is painted but (upperIdx - 1) is not.
Private Sub PaintMissingInBlock(wsStrip As Worksheet, arr() As Boolean, _
                                ByVal upperIdx As Long, ByVal N As Long, _
                                ByVal blockRows As Long, ByVal isRHS As Boolean, _
                                ByVal startCol As Long, ByVal endCol As Long, _
                                ByVal color As Long)
    If upperIdx < 2 Then Exit Sub
    If endCol >= g_totalCols Then endCol = g_totalCols - 1
    Dim c As Long
    For c = startCol To endCol
        If arr(upperIdx, c) And Not arr(upperIdx - 1, c) Then
            Dim blockNo As Long: blockNo = c \ COLS_PER_BLOCK
            Dim colInBlock As Long: colInBlock = c Mod COLS_PER_BLOCK
            Dim blockTop As Long: blockTop = blockNo * blockRows + 1
            Dim cellCol As Long: cellCol = C_DATA_START + colInBlock
            Dim cellRow As Long
            If isRHS Then
                cellRow = blockTop + (N - upperIdx + 1)
            Else
                cellRow = blockTop + N + 1 + upperIdx
            End If
            wsStrip.Cells(cellRow, cellCol).Interior.Color = color
        End If
    Next c
End Sub


' Write one audit cell. isIssue=True → red text; False → green "OK" text.
Private Sub WriteAudit(ws As Worksheet, r As Long, c As Long, _
                       msg As String, isIssue As Boolean)
    With ws.Cells(r, c)
        .Value = msg
        .WrapText = True
        .VerticalAlignment = xlCenter
        .HorizontalAlignment = xlLeft
        .Font.Bold = True
        If isIssue Then
            .Font.Color = COLOR_RED
        Else
            .Font.Color = COLOR_OK_GREEN
        End If
    End With
End Sub
