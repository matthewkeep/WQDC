Option Explicit
' SimLog: Date-centric live log with UPSERT logic.
' Dependencies: Core, Schema, Helpers, Setup
'
' tblLive_{site}: One row per date with Std/Enh predictions side-by-side.
' Standard run creates/updates rows, Enhanced updates existing rows.
' RunId stored without prefix (same as History table).
' Columns: Date, Days, StdVol, Std[7 chem], EnhVol, Enh[7 chem], EnhHid[7 chem], ErrVol, ErrEC, RunId

' ==== Write Functions =======================================================

Public Sub WriteLog(ByRef r As Result, ByRef cfg As Config, ByVal runId As String, ByVal site As String, ByVal mode As String)
    ' UPSERT to site's live table - creates/updates rows by date
    ' mode = "Standard" or "Enhanced"
    On Error GoTo Fail

    If mode = "Standard" Then
        WriteLiveStandard r, cfg, runId, site
    Else
        WriteLiveEnhanced r, cfg, runId, site
    End If
    Exit Sub

Fail:
    Error.TraceErr "SimLog.WriteLog"
End Sub

Private Sub WriteLiveStandard(ByRef r As Result, ByRef cfg As Config, ByVal runId As String, ByVal site As String)
    ' Writes Standard predictions - creates rows if needed
    ' Simple overlay: writes from sample date forward, Days = 0,1,2...
    Dim tbl As ListObject
    Dim i As Long, j As Long, n As Long, rowIdx As Long
    Dim logDate As Date
    Dim daysCol As Long, volCol As Long, runIdCol As Long
    Dim chemCols(1 To 7) As Long
    Dim runDate As Date

    Set tbl = GetLiveTable(site)
    If tbl Is Nothing Then Exit Sub
    runDate = GetRunDate()

    ' Clear old trigger formatting before writing new data
    ClearTriggerFormatting tbl, "Std"

    ' Pre-fetch all column indices (avoids O(n*7) lookups in loop)
    daysCol = Helpers.ColIdx(tbl, Schema.LIVE_COL_DAYS)
    volCol = Helpers.ColIdx(tbl, Schema.LIVE_COL_STD_VOL)
    runIdCol = Helpers.ColIdx(tbl, Schema.LIVE_COL_RUNID)
    For j = 1 To Core.METRIC_COUNT
        chemCols(j) = Helpers.ColIdx(tbl, Schema.StdChemColName(j))
    Next j

    n = UBound(r.Snaps)
    For i = 0 To n
        logDate = cfg.StartDate + i

        ' Find or create row for this date
        rowIdx = EnsureRowForDate(tbl, logDate)
        If rowIdx = 0 Then Exit Sub  ' Failed to create row

        ' Write Days column (relative to run date: negative=past, 0=today, positive=future)
        With tbl.DataBodyRange
            If daysCol > 0 Then .Cells(rowIdx, daysCol) = CLng(logDate - runDate)

            ' Write Standard columns: Volume + all 7 chemistry metrics
            If volCol > 0 Then .Cells(rowIdx, volCol) = r.Snaps(i).Vol
            For j = 1 To Core.METRIC_COUNT
                If chemCols(j) > 0 Then .Cells(rowIdx, chemCols(j)) = r.Snaps(i).Chem(j)
            Next j
            If runIdCol > 0 Then .Cells(rowIdx, runIdCol) = runId
        End With
    Next i

    ' Calculate discrepancy from telemetry
    WriteDiscrepancy tbl, site

    ' Apply all row formatting in single pass (shading + trigger)
    Dim triggerDate As Date, triggerMetric As String
    If r.TriggerDay <> Core.NO_TRIGGER Then
        triggerDate = cfg.StartDate + r.TriggerDay
        triggerMetric = r.TriggerMetric
    End If
    ApplyRowFormatting tbl, cfg.StartDate, runDate, triggerDate, triggerMetric, "Std"
End Sub

Private Sub WriteLiveEnhanced(ByRef r As Result, ByRef cfg As Config, ByVal runId As String, ByVal site As String)
    ' Writes Enhanced predictions + hidden layer - updates existing rows
    Dim tbl As ListObject
    Dim i As Long, j As Long, n As Long, rowIdx As Long
    Dim logDate As Date
    Dim daysCol As Long, volCol As Long, runIdCol As Long
    Dim chemCols(1 To 7) As Long, hidCols(1 To 7) As Long
    Dim runDate As Date

    Set tbl = GetLiveTable(site)
    If tbl Is Nothing Then Exit Sub
    runDate = GetRunDate()

    ' Clear old trigger formatting before writing new data
    ClearTriggerFormatting tbl, "Enh"

    ' Pre-fetch all column indices (avoids O(n*14) lookups in loop)
    daysCol = Helpers.ColIdx(tbl, Schema.LIVE_COL_DAYS)
    volCol = Helpers.ColIdx(tbl, Schema.LIVE_COL_ENH_VOL)
    runIdCol = Helpers.ColIdx(tbl, Schema.LIVE_COL_RUNID)
    For j = 1 To Core.METRIC_COUNT
        chemCols(j) = Helpers.ColIdx(tbl, Schema.EnhChemColName(j))
        hidCols(j) = Helpers.ColIdx(tbl, Schema.EnhHidColName(j))
    Next j

    n = UBound(r.Snaps)
    For i = 0 To n
        logDate = cfg.StartDate + i

        ' Find row for this date (should exist from Standard run)
        rowIdx = Helpers.FindRowByDate(tbl, logDate)
        If rowIdx = 0 Then
            ' Row doesn't exist - create it (Enhanced-only run)
            rowIdx = EnsureRowForDate(tbl, logDate)
            If rowIdx = 0 Then Exit Sub
        End If

        ' Write Days column (relative to run date: negative=past, 0=today, positive=future)
        With tbl.DataBodyRange
            If daysCol > 0 Then .Cells(rowIdx, daysCol) = CLng(logDate - runDate)

            ' Write Enhanced columns: Volume + all 7 chemistry visible + hidden
            If volCol > 0 Then .Cells(rowIdx, volCol) = r.Snaps(i).Vol

            For j = 1 To Core.METRIC_COUNT
                If chemCols(j) > 0 Then .Cells(rowIdx, chemCols(j)) = r.Snaps(i).Chem(j)
                If hidCols(j) > 0 Then .Cells(rowIdx, hidCols(j)) = r.Snaps(i).Hidden(j)
            Next j

            If runIdCol > 0 Then .Cells(rowIdx, runIdCol) = runId
        End With
    Next i

    ' Calculate discrepancy from telemetry
    WriteDiscrepancy tbl, site

    ' Apply Enh trigger formatting only (shading already done by Standard, preserves Std trigger box)
    If r.TriggerDay <> Core.NO_TRIGGER Then
        ApplyTriggerFormatting tbl, cfg.StartDate + r.TriggerDay, r.TriggerMetric, "Enh"
    End If
End Sub

Private Sub WriteDiscrepancy(ByVal tbl As ListObject, ByVal site As String)
    ' Calculates ErrVol/ErrEC = Telemetry - Predicted (Enhanced if available, else Standard)
    ' Blank if no telemetry for date
    Dim tblTelem As ListObject, i As Long, rowIdx As Long
    Dim logDate As Date, telemEC As Variant, telemVol As Variant
    Dim predEC As Double, predVol As Double
    Dim ecCol As Long, volCol As Long, errVolCol As Long, errECCol As Long
    Dim enhVolCol As Long, enhECCol As Long, stdVolCol As Long, stdECCol As Long

    If Not Helpers.HasData(tbl) Then Exit Sub
    Set tblTelem = Helpers.GetTable(Schema.SHEET_RESULTS, Schema.TABLE_TELEMETRY)
    If Not Helpers.HasData(tblTelem) Then Exit Sub

    ecCol = Helpers.ColIdx(tblTelem, Helpers.TelemECColName(site))
    volCol = Helpers.ColIdx(tblTelem, Helpers.TelemVolColName(site))
    If ecCol = 0 And volCol = 0 Then Exit Sub

    errVolCol = Helpers.ColIdx(tbl, Schema.LIVE_COL_ERR_VOL)
    errECCol = Helpers.ColIdx(tbl, Schema.LIVE_COL_ERR_EC)
    enhVolCol = Helpers.ColIdx(tbl, Schema.LIVE_COL_ENH_VOL)
    enhECCol = Helpers.ColIdx(tbl, Schema.LIVE_COL_ENH_EC)
    stdVolCol = Helpers.ColIdx(tbl, Schema.LIVE_COL_STD_VOL)
    stdECCol = Helpers.ColIdx(tbl, Schema.LIVE_COL_STD_EC)

    For i = 1 To tbl.ListRows.Count
        logDate = tbl.DataBodyRange.Cells(i, 1).Value
        rowIdx = Helpers.FindRowByDate(tblTelem, logDate)

        If rowIdx > 0 Then
            If ecCol > 0 Then telemEC = tblTelem.DataBodyRange.Cells(rowIdx, ecCol).Value
            If volCol > 0 Then telemVol = tblTelem.DataBodyRange.Cells(rowIdx, volCol).Value

            ' EC error
            If errECCol > 0 And Not IsEmpty(telemEC) Then
                If enhECCol > 0 And Not IsEmpty(tbl.DataBodyRange.Cells(i, enhECCol).Value) Then
                    predEC = tbl.DataBodyRange.Cells(i, enhECCol).Value
                Else
                    predEC = tbl.DataBodyRange.Cells(i, stdECCol).Value
                End If
                tbl.DataBodyRange.Cells(i, errECCol).Value = CDbl(telemEC) - predEC
            End If

            ' Volume error
            If errVolCol > 0 And Not IsEmpty(telemVol) Then
                If enhVolCol > 0 And Not IsEmpty(tbl.DataBodyRange.Cells(i, enhVolCol).Value) Then
                    predVol = tbl.DataBodyRange.Cells(i, enhVolCol).Value
                Else
                    predVol = tbl.DataBodyRange.Cells(i, stdVolCol).Value
                End If
                tbl.DataBodyRange.Cells(i, errVolCol).Value = CDbl(telemVol) - predVol
            End If
        Else
            If errECCol > 0 Then tbl.DataBodyRange.Cells(i, errECCol).ClearContents
            If errVolCol > 0 Then tbl.DataBodyRange.Cells(i, errVolCol).ClearContents
        End If
    Next i
End Sub

' ==== Formatting Helpers ====================================================

Private Sub ClearTriggerFormatting(ByVal tbl As ListObject, ByVal prefix As String)
    ' Clears red+bold trigger formatting from Vol + chemistry columns for Std or Enh
    ' Uses bulk column operations (O(1) per column, not O(n) per row)
    Dim j As Long
    If Not Helpers.HasData(tbl) Then Exit Sub

    ClearColumnFormat tbl, IIf(prefix = "Std", Schema.LIVE_COL_STD_VOL, Schema.LIVE_COL_ENH_VOL)
    For j = 1 To Core.METRIC_COUNT
        If prefix = "Std" Then
            ClearColumnFormat tbl, Schema.StdChemColName(j)
        Else
            ClearColumnFormat tbl, Schema.EnhChemColName(j)
        End If
    Next j
End Sub

Private Sub ClearColumnFormat(ByVal tbl As ListObject, ByVal colName As String)
    ' Clears font formatting (bold + color) from a table column - bulk operation
    Dim col As Long
    col = Helpers.ColIdx(tbl, colName)
    If col > 0 Then
        With tbl.DataBodyRange.Columns(col).Font
            .Bold = False
            .ColorIndex = xlAutomatic
        End With
    End If
End Sub

Private Sub ApplyRowFormatting(ByVal tbl As ListObject, ByVal sampleDate As Date, ByVal runDate As Date, _
                                ByVal triggerDate As Date, ByVal triggerMetric As String, ByVal prefix As String)
    ' Single pass: clears old formatting, applies shading + trigger formatting
    ' Efficient: only touches rows that need changes
    Dim i As Long, rowDate As Date
    Dim triggerCol As Long, colName As String
    If Not Helpers.HasData(tbl) Then Exit Sub

    ' Get trigger column index
    If Len(triggerMetric) > 0 Then
        If triggerMetric = "Volume" Then
            colName = IIf(prefix = "Std", Schema.LIVE_COL_STD_VOL, Schema.LIVE_COL_ENH_VOL)
        Else
            colName = prefix & triggerMetric
        End If
        triggerCol = Helpers.ColIdx(tbl, colName)
    End If

    ' Apply shading to all rows, clear old borders
    Dim triggerRow As Long
    For i = 1 To tbl.ListRows.Count
        rowDate = tbl.DataBodyRange.Cells(i, 1).Value
        tbl.ListRows(i).Range.Borders.LineStyle = xlNone

        If rowDate = sampleDate Then
            tbl.ListRows(i).Range.Interior.Color = Schema.COLOR_SAMPLE_DATE
        ElseIf rowDate = runDate Then
            tbl.ListRows(i).Range.Interior.Color = Schema.COLOR_RUN_DATE
        Else
            tbl.ListRows(i).Range.Interior.ColorIndex = xlNone
        End If

        If rowDate = triggerDate Then triggerRow = i
    Next i

    ' Apply trigger formatting last (bold + red + box)
    If triggerRow > 0 And triggerCol > 0 Then
        With tbl.DataBodyRange.Cells(triggerRow, triggerCol)
            .Font.Bold = True
            .Font.Color = Schema.COLOR_TRIGGER_FONT
        End With
        tbl.ListRows(triggerRow).Range.BorderAround xlContinuous, xlThin
    End If
End Sub

Private Sub ApplyTriggerFormatting(ByVal tbl As ListObject, ByVal triggerDate As Date, _
                                    ByVal triggerMetric As String, ByVal prefix As String)
    ' Applies trigger formatting (bold + red + box) to a specific row
    ' Used by Enhanced run to add its trigger without clearing Standard's
    Dim rowIdx As Long, colName As String, col As Long
    If Not Helpers.HasData(tbl) Then Exit Sub

    rowIdx = Helpers.FindRowByDate(tbl, triggerDate)
    If rowIdx = 0 Then Exit Sub

    ' Build column name based on prefix and metric
    If triggerMetric = "Volume" Then
        colName = IIf(prefix = "Std", Schema.LIVE_COL_STD_VOL, Schema.LIVE_COL_ENH_VOL)
    Else
        colName = prefix & triggerMetric
    End If

    col = Helpers.ColIdx(tbl, colName)
    If col > 0 Then
        With tbl.DataBodyRange.Cells(rowIdx, col)
            .Font.Bold = True
            .Font.Color = Schema.COLOR_TRIGGER_FONT
        End With
    End If

    tbl.ListRows(rowIdx).Range.BorderAround xlContinuous, xlThin
End Sub

' ==== Row Lookup/Creation ===================================================

Private Function EnsureRowForDate(ByVal tbl As ListObject, ByVal targetDate As Date) As Long
    ' Finds row for date or creates new row in sorted position
    ' Returns row index (1-based)
    Dim i As Long, insertPos As Long, newRow As ListRow
    Dim rowDate As Date

    ' Check if row exists
    EnsureRowForDate = Helpers.FindRowByDate(tbl, targetDate)
    If EnsureRowForDate > 0 Then Exit Function

    ' Find insert position (keep sorted by date)
    insertPos = 0
    If Not tbl.DataBodyRange Is Nothing Then
        For i = 1 To tbl.ListRows.Count
            rowDate = tbl.DataBodyRange.Cells(i, 1).Value
            If targetDate < rowDate Then
                insertPos = i
                Exit For
            End If
        Next i
    End If

    ' Insert new row
    If insertPos > 0 Then
        Set newRow = tbl.ListRows.Add(insertPos)
        EnsureRowForDate = insertPos
    Else
        Set newRow = tbl.ListRows.Add
        EnsureRowForDate = tbl.ListRows.Count
    End If

    ' Set date value
    newRow.Range.Cells(1, 1).Value = targetDate
End Function

' ==== Clear Functions =======================================================

Public Sub ClearEnhancedColumns(ByVal site As String)
    ' Clears all Enhanced columns (Vol, Chem, Hidden) - used when running Standard-only
    Dim tbl As ListObject, j As Long, col As Long

    Set tbl = GetLiveTable(site)
    If Not Helpers.HasData(tbl) Then Exit Sub

    ' Clear Enhanced visible columns
    col = Helpers.ColIdx(tbl, Schema.LIVE_COL_ENH_VOL)
    If col > 0 Then tbl.DataBodyRange.Columns(col).ClearContents

    For j = 1 To Core.METRIC_COUNT
        col = Helpers.ColIdx(tbl, Schema.EnhChemColName(j))
        If col > 0 Then tbl.DataBodyRange.Columns(col).ClearContents
        col = Helpers.ColIdx(tbl, Schema.EnhHidColName(j))
        If col > 0 Then tbl.DataBodyRange.Columns(col).ClearContents
    Next j

    ' Clear Enhanced trigger formatting
    ClearTriggerFormatting tbl, "Enh"
End Sub

' ==== Delete Functions ======================================================

Public Sub DeleteAfterDate(ByVal cutoffDate As Date, ByVal site As String)
    ' Deletes all rows with Date > cutoffDate (for rollback)
    Dim tbl As ListObject
    Dim i As Long, rowDate As Date

    Set tbl = GetLiveTable(site)
    If Not Helpers.HasData(tbl) Then Exit Sub

    ' Delete from bottom up to avoid index issues
    For i = tbl.ListRows.Count To 1 Step -1
        rowDate = tbl.DataBodyRange.Cells(i, 1).Value
        If rowDate > cutoffDate Then
            tbl.ListRows(i).Delete
        End If
    Next i
End Sub

' ==== Table Access ===========================================================

Private Function GetRunDate() As Date
    ' Returns run date from Inputs sheet, or today as fallback
    GetRunDate = Helpers.GetDateVal(Helpers.GetSheet(Schema.SHEET_INPUT), Schema.NAME_RUN_DATE)
    If GetRunDate = 0 Then GetRunDate = Date
End Function

Private Function GetLiveTable(ByVal site As String) As ListObject
    ' Returns site's live table, creating it if necessary
    Set GetLiveTable = Helpers.GetSiteTable(Schema.SHEET_LOG, Schema.LIVE_TABLE_PREFIX, site)
    If GetLiveTable Is Nothing Then
        Setup.EnsureSiteLiveTable site
        Set GetLiveTable = Helpers.GetSiteTable(Schema.SHEET_LOG, Schema.LIVE_TABLE_PREFIX, site)
    End If
End Function
