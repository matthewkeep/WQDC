Option Explicit
' SimLog: Date-centric live log with UPSERT logic.
' Dependencies: Core, Schema, Setup, Data (for telemetry)
'
' tblLive_{site}: One row per date with Std/Enh predictions side-by-side.
' Standard run creates/updates rows, Enhanced updates existing rows.
' Columns: Date, Days, StdVol, Std[7 chem], EnhVol, Enh[7 chem], EnhHid[7 chem], ErrVol, ErrEC, RunId

' ==== Write Functions =======================================================

Public Sub WriteLog(ByRef r As Result, ByRef cfg As Config, ByVal runId As String, ByVal site As String)
    ' UPSERT to site's live table - creates/updates rows by date
    ' Detects Standard vs Enhanced from runId prefix (STD- or ENH-)
    On Error GoTo Fail

    If Left$(runId, 3) = "STD" Then
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
    Dim tbl As ListObject
    Dim i As Long, j As Long, n As Long, rowIdx As Long
    Dim logDate As Date, col As Long, daysCol As Long
    Dim runDate As Date

    Set tbl = GetLiveTable(site)
    If tbl Is Nothing Then Exit Sub

    runDate = Date  ' Run date is always today
    daysCol = Helpers.ColIdx(tbl, Schema.LIVE_COL_DAYS)

    n = UBound(r.Snaps)
    For i = 0 To n
        logDate = cfg.StartDate + i

        ' Find or create row for this date
        rowIdx = EnsureRowForDate(tbl, logDate)
        If rowIdx = 0 Then Exit Sub  ' Failed to create row

        ' Write Days column (relative to run date)
        With tbl.DataBodyRange
            If daysCol > 0 Then .Cells(rowIdx, daysCol) = CLng(logDate - runDate)

            ' Write Standard columns: Volume + all 7 chemistry metrics
            .Cells(rowIdx, Helpers.ColIdx(tbl, Schema.LIVE_COL_STD_VOL)) = r.Snaps(i).Vol
            For j = 1 To Schema.ChemistryCount()
                col = Helpers.ColIdx(tbl, Schema.StdChemColName(j))
                If col > 0 Then .Cells(rowIdx, col) = r.Snaps(i).Chem(j)
            Next j
            .Cells(rowIdx, Helpers.ColIdx(tbl, Schema.LIVE_COL_RUNID)) = runId
        End With
    Next i

    ' Calculate discrepancy from telemetry
    WriteDiscrepancy tbl, site

    ' Apply row shading for sample and run dates
    ApplyRowShading tbl, cfg.StartDate, runDate

    ' Format triggered cell if trigger occurred
    If r.TriggerDay <> Core.NO_TRIGGER Then
        FormatLiveTriggerCell tbl, cfg.StartDate + r.TriggerDay, r.TriggerMetric, "Std"
    End If
End Sub

Private Sub WriteLiveEnhanced(ByRef r As Result, ByRef cfg As Config, ByVal runId As String, ByVal site As String)
    ' Writes Enhanced predictions + hidden layer - updates existing rows
    Dim tbl As ListObject
    Dim i As Long, j As Long, n As Long, rowIdx As Long
    Dim logDate As Date, col As Long, daysCol As Long
    Dim runDate As Date

    Set tbl = GetLiveTable(site)
    If tbl Is Nothing Then Exit Sub

    runDate = Date  ' Run date is always today
    daysCol = Helpers.ColIdx(tbl, Schema.LIVE_COL_DAYS)

    n = UBound(r.Snaps)
    For i = 0 To n
        logDate = cfg.StartDate + i

        ' Find row for this date (should exist from Standard run)
        rowIdx = FindRowByDate(tbl, logDate)
        If rowIdx = 0 Then
            ' Row doesn't exist - create it (Enhanced-only run)
            rowIdx = EnsureRowForDate(tbl, logDate)
            If rowIdx = 0 Then Exit Sub
        End If

        ' Write Days column (relative to run date) - update in case row was just created
        With tbl.DataBodyRange
            If daysCol > 0 Then .Cells(rowIdx, daysCol) = CLng(logDate - runDate)

            ' Write Enhanced columns: Volume + all 7 chemistry visible + hidden
            .Cells(rowIdx, Helpers.ColIdx(tbl, Schema.LIVE_COL_ENH_VOL)) = r.Snaps(i).Vol

            For j = 1 To Schema.ChemistryCount()
                ' Visible layer chemistry
                col = Helpers.ColIdx(tbl, Schema.EnhChemColName(j))
                If col > 0 Then .Cells(rowIdx, col) = r.Snaps(i).Chem(j)
                ' Hidden layer mass (for TwoBucket continuity)
                col = Helpers.ColIdx(tbl, Schema.EnhHidColName(j))
                If col > 0 Then .Cells(rowIdx, col) = r.Snaps(i).Hidden(j)
            Next j

            .Cells(rowIdx, Helpers.ColIdx(tbl, Schema.LIVE_COL_RUNID)) = runId
        End With
    Next i

    ' Calculate discrepancy from telemetry
    WriteDiscrepancy tbl, site

    ' Format triggered cell if trigger occurred
    If r.TriggerDay <> Core.NO_TRIGGER Then
        FormatLiveTriggerCell tbl, cfg.StartDate + r.TriggerDay, r.TriggerMetric, "Enh"
    End If
End Sub

Private Sub WriteDiscrepancy(ByVal tbl As ListObject, ByVal site As String)
    ' Calculates ErrVol = TelemetryVol - PredictedVol (Enhanced if available, else Standard)
    ' Calculates ErrEC = TelemetryEC - PredictedEC
    ' Leaves blank if no telemetry for that date
    Dim tblTelem As ListObject
    Dim i As Long, rowIdx As Long
    Dim logDate As Date, telemEC As Variant, telemVol As Variant
    Dim predEC As Double, predVol As Double
    Dim ecCol As Long, volCol As Long
    Dim errVolCol As Long, errECCol As Long
    Dim enhVolCol As Long, enhECCol As Long
    Dim stdVolCol As Long, stdECCol As Long

    If Not Helpers.HasData(tbl) Then Exit Sub

    ' Get telemetry table
    Set tblTelem = Helpers.WithTableData(Schema.SHEET_RESULTS, Schema.TABLE_TELEMETRY)
    If tblTelem Is Nothing Then Exit Sub

    ' Get telemetry column indices for this site
    ecCol = Helpers.ColIdx(tblTelem, Helpers.TelemECColName(site))
    volCol = Helpers.ColIdx(tblTelem, Helpers.TelemVolColName(site))
    If ecCol = 0 And volCol = 0 Then Exit Sub  ' No telemetry columns for this site

    ' Get live table column indices
    errVolCol = Helpers.ColIdx(tbl, Schema.LIVE_COL_ERR_VOL)
    errECCol = Helpers.ColIdx(tbl, Schema.LIVE_COL_ERR_EC)
    enhVolCol = Helpers.ColIdx(tbl, Schema.LIVE_COL_ENH_VOL)
    enhECCol = Helpers.ColIdx(tbl, Schema.LIVE_COL_ENH_EC)
    stdVolCol = Helpers.ColIdx(tbl, Schema.LIVE_COL_STD_VOL)
    stdECCol = Helpers.ColIdx(tbl, Schema.LIVE_COL_STD_EC)

    ' Process each row in live table
    For i = 1 To tbl.ListRows.Count
        logDate = tbl.DataBodyRange.Cells(i, 1).Value

        ' Find matching telemetry row
        rowIdx = FindTelemRowByDate(tblTelem, logDate)
        If rowIdx > 0 Then
            ' Get telemetry values (may be empty)
            If ecCol > 0 Then telemEC = tblTelem.DataBodyRange.Cells(rowIdx, ecCol).Value
            If volCol > 0 Then telemVol = tblTelem.DataBodyRange.Cells(rowIdx, volCol).Value

            ' Calculate EC discrepancy
            If errECCol > 0 And Not IsEmpty(telemEC) Then
                ' Use Enhanced if available, else Standard
                predEC = 0
                If enhECCol > 0 And Not IsEmpty(tbl.DataBodyRange.Cells(i, enhECCol).Value) Then
                    predEC = tbl.DataBodyRange.Cells(i, enhECCol).Value
                ElseIf stdECCol > 0 And Not IsEmpty(tbl.DataBodyRange.Cells(i, stdECCol).Value) Then
                    predEC = tbl.DataBodyRange.Cells(i, stdECCol).Value
                End If
                tbl.DataBodyRange.Cells(i, errECCol).Value = CDbl(telemEC) - predEC
            End If

            ' Calculate Volume discrepancy
            If errVolCol > 0 And Not IsEmpty(telemVol) Then
                ' Use Enhanced if available, else Standard
                predVol = 0
                If enhVolCol > 0 And Not IsEmpty(tbl.DataBodyRange.Cells(i, enhVolCol).Value) Then
                    predVol = tbl.DataBodyRange.Cells(i, enhVolCol).Value
                ElseIf stdVolCol > 0 And Not IsEmpty(tbl.DataBodyRange.Cells(i, stdVolCol).Value) Then
                    predVol = tbl.DataBodyRange.Cells(i, stdVolCol).Value
                End If
                tbl.DataBodyRange.Cells(i, errVolCol).Value = CDbl(telemVol) - predVol
            End If
        Else
            ' No telemetry for this date - clear discrepancy
            If errECCol > 0 Then tbl.DataBodyRange.Cells(i, errECCol).ClearContents
            If errVolCol > 0 Then tbl.DataBodyRange.Cells(i, errVolCol).ClearContents
        End If
    Next i
End Sub

' ==== Formatting Helpers ====================================================

Private Sub ApplyRowShading(ByVal tbl As ListObject, ByVal sampleDate As Date, ByVal runDate As Date)
    ' Applies background color to sample date and run date rows
    Dim i As Long, rowDate As Date
    If Not Helpers.HasData(tbl) Then Exit Sub

    For i = 1 To tbl.ListRows.Count
        rowDate = tbl.DataBodyRange.Cells(i, 1).Value
        tbl.ListRows(i).Range.Interior.ColorIndex = xlNone  ' Clear first

        If rowDate = sampleDate Then
            tbl.ListRows(i).Range.Interior.Color = Schema.COLOR_SAMPLE_DATE
        ElseIf rowDate = runDate Then
            tbl.ListRows(i).Range.Interior.Color = Schema.COLOR_RUN_DATE
        End If
    Next i
End Sub

Private Sub FormatLiveTriggerCell(ByVal tbl As ListObject, ByVal triggerDate As Date, _
                                   ByVal metricName As String, ByVal prefix As String)
    ' Formats the triggered metric cell red + bold in tblLive
    Dim rowIdx As Long, colName As String, col As Long

    rowIdx = Helpers.FindRowByDate(tbl, triggerDate)
    If rowIdx = 0 Then Exit Sub

    ' Build column name based on prefix and metric
    If metricName = "Volume" Then
        colName = IIf(prefix = "Std", Schema.LIVE_COL_STD_VOL, Schema.LIVE_COL_ENH_VOL)
    Else
        colName = prefix & metricName  ' e.g., "StdEC" or "EnhEC"
    End If

    col = Helpers.ColIdx(tbl, colName)
    If col > 0 Then
        With tbl.DataBodyRange.Cells(rowIdx, col)
            .Font.Bold = True
            .Font.Color = Schema.COLOR_TRIGGER_FONT
        End With
    End If
End Sub

' ==== Row Lookup/Creation ===================================================

Private Function FindRowByDate(ByVal tbl As ListObject, ByVal targetDate As Date) As Long
    ' Returns row index (1-based) for date, or 0 if not found
    ' Delegates to shared O(1) utility
    FindRowByDate = Helpers.FindRowByDate(tbl, targetDate)
End Function

Private Function EnsureRowForDate(ByVal tbl As ListObject, ByVal targetDate As Date) As Long
    ' Finds row for date or creates new row in sorted position
    ' Returns row index (1-based)
    Dim i As Long, insertPos As Long, newRow As ListRow
    Dim rowDate As Date

    ' Check if row exists
    EnsureRowForDate = FindRowByDate(tbl, targetDate)
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

Private Function FindTelemRowByDate(ByVal tbl As ListObject, ByVal targetDate As Date) As Long
    ' Returns row index (1-based) for date in telemetry table, or 0 if not found
    ' Delegates to shared O(1) utility
    FindTelemRowByDate = Helpers.FindRowByDate(tbl, targetDate)
End Function

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

Public Sub ClearSiteLog(ByVal site As String)
    ' Clears entire live table for site
    Dim tbl As ListObject

    Set tbl = GetLiveTable(site)
    If tbl Is Nothing Then Exit Sub
    If Not tbl.DataBodyRange Is Nothing Then tbl.DataBodyRange.Delete
End Sub

' ==== Read Functions ========================================================

Public Function GetLatestLogDate(ByVal site As String) As Date
    ' Returns the most recent date in site's live table (0 if empty)
    Dim tbl As ListObject
    Dim i As Long, d As Date, maxDate As Date

    Set tbl = GetLiveTable(site)
    If Not Helpers.HasData(tbl) Then Exit Function

    maxDate = 0
    For i = 1 To tbl.ListRows.Count
        d = tbl.DataBodyRange.Cells(i, 1).Value
        If d > maxDate Then maxDate = d
    Next i
    GetLatestLogDate = maxDate
End Function

' ==== Table Access ===========================================================

Private Function GetLiveTable(ByVal site As String) As ListObject
    ' Returns site's live table, creating it if necessary
    Dim ws As Worksheet, tblName As String

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(Schema.SHEET_LOG)
    On Error GoTo 0
    If ws Is Nothing Then Exit Function

    tblName = Helpers.LiveTableName(site)

    ' Try to get existing table
    On Error Resume Next
    Set GetLiveTable = ws.ListObjects(tblName)
    On Error GoTo 0

    ' Create if doesn't exist
    If GetLiveTable Is Nothing Then
        Setup.EnsureSiteLiveTable site
        On Error Resume Next
        Set GetLiveTable = ws.ListObjects(tblName)
        On Error GoTo 0
    End If
End Function
