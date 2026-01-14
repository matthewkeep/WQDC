Option Explicit
' SimLog: Date-centric live log with UPSERT logic.
' Dependencies: Core, Schema, Setup, Data (for telemetry)
'
' tblLive_{site}: One row per date with Std/Enh predictions side-by-side.
' Standard run creates/updates rows, Enhanced updates existing rows.
' Columns: Date, StdVol, StdEC, EnhVol, EnhEC, EnhHid1-7, ErrVol, ErrEC, RunId

' ==== Write Functions =======================================================

Public Sub WriteLog(ByRef r As Result, ByRef cfg As Config, ByVal runId As String, ByVal site As String)
    ' UPSERT to site's live table - creates/updates rows by date
    ' Detects Standard vs Enhanced from runId prefix (STD- or ENH-)
    If Left$(runId, 3) = "STD" Then
        WriteLiveStandard r, cfg, runId, site
    Else
        WriteLiveEnhanced r, cfg, runId, site
    End If
End Sub

Private Sub WriteLiveStandard(ByRef r As Result, ByRef cfg As Config, ByVal runId As String, ByVal site As String)
    ' Writes Standard predictions - creates rows if needed
    Dim tbl As ListObject
    Dim i As Long, n As Long, rowIdx As Long
    Dim logDate As Date

    Set tbl = GetLiveTable(site)
    If tbl Is Nothing Then Exit Sub

    n = UBound(r.Snaps)
    For i = 0 To n
        logDate = cfg.StartDate + i

        ' Find or create row for this date
        rowIdx = EnsureRowForDate(tbl, logDate)
        If rowIdx = 0 Then Exit Sub  ' Failed to create row

        ' Write Standard columns
        With tbl.DataBodyRange
            .Cells(rowIdx, Schema.ColIdx(tbl, Schema.LIVE_COL_STD_VOL)) = r.Snaps(i).Vol
            .Cells(rowIdx, Schema.ColIdx(tbl, Schema.LIVE_COL_STD_EC)) = r.Snaps(i).Chem(1)
            .Cells(rowIdx, Schema.ColIdx(tbl, Schema.LIVE_COL_RUNID)) = runId
        End With
    Next i

    ' Calculate discrepancy from telemetry
    WriteDiscrepancy tbl, site
End Sub

Private Sub WriteLiveEnhanced(ByRef r As Result, ByRef cfg As Config, ByVal runId As String, ByVal site As String)
    ' Writes Enhanced predictions + hidden layer - updates existing rows
    Dim tbl As ListObject
    Dim i As Long, j As Long, n As Long, rowIdx As Long
    Dim logDate As Date, hidCol As Long

    Set tbl = GetLiveTable(site)
    If tbl Is Nothing Then Exit Sub

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

        ' Write Enhanced columns
        With tbl.DataBodyRange
            .Cells(rowIdx, Schema.ColIdx(tbl, Schema.LIVE_COL_ENH_VOL)) = r.Snaps(i).Vol
            .Cells(rowIdx, Schema.ColIdx(tbl, Schema.LIVE_COL_ENH_EC)) = r.Snaps(i).Chem(1)

            ' Write hidden layer (for TwoBucket continuity)
            For j = 1 To Core.METRIC_COUNT
                hidCol = Schema.ColIdx(tbl, Schema.EnhHidColName(j))
                If hidCol > 0 Then .Cells(rowIdx, hidCol) = r.Snaps(i).Hidden(j)
            Next j

            .Cells(rowIdx, Schema.ColIdx(tbl, Schema.LIVE_COL_RUNID)) = runId
        End With
    Next i

    ' Calculate discrepancy from telemetry
    WriteDiscrepancy tbl, site
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

    If tbl.DataBodyRange Is Nothing Then Exit Sub

    ' Get telemetry table
    Set tblTelem = Schema.GetTable(Schema.SHEET_TELEMETRY, Schema.TABLE_TELEMETRY)
    If tblTelem Is Nothing Then Exit Sub
    If tblTelem.DataBodyRange Is Nothing Then Exit Sub

    ' Get telemetry column indices for this site
    ecCol = Schema.ColIdx(tblTelem, Schema.TelemECColName(site))
    volCol = Schema.ColIdx(tblTelem, Schema.TelemVolColName(site))
    If ecCol = 0 And volCol = 0 Then Exit Sub  ' No telemetry columns for this site

    ' Get live table column indices
    errVolCol = Schema.ColIdx(tbl, Schema.LIVE_COL_ERR_VOL)
    errECCol = Schema.ColIdx(tbl, Schema.LIVE_COL_ERR_EC)
    enhVolCol = Schema.ColIdx(tbl, Schema.LIVE_COL_ENH_VOL)
    enhECCol = Schema.ColIdx(tbl, Schema.LIVE_COL_ENH_EC)
    stdVolCol = Schema.ColIdx(tbl, Schema.LIVE_COL_STD_VOL)
    stdECCol = Schema.ColIdx(tbl, Schema.LIVE_COL_STD_EC)

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
                If enhECCol > 0 And Not IsEmpty(tbl.DataBodyRange.Cells(i, enhECCol).Value) Then
                    predEC = tbl.DataBodyRange.Cells(i, enhECCol).Value
                ElseIf stdECCol > 0 Then
                    predEC = tbl.DataBodyRange.Cells(i, stdECCol).Value
                End If
                If predEC > 0 Then
                    tbl.DataBodyRange.Cells(i, errECCol).Value = CDbl(telemEC) - predEC
                End If
            End If

            ' Calculate Volume discrepancy
            If errVolCol > 0 And Not IsEmpty(telemVol) Then
                ' Use Enhanced if available, else Standard
                If enhVolCol > 0 And Not IsEmpty(tbl.DataBodyRange.Cells(i, enhVolCol).Value) Then
                    predVol = tbl.DataBodyRange.Cells(i, enhVolCol).Value
                ElseIf stdVolCol > 0 Then
                    predVol = tbl.DataBodyRange.Cells(i, stdVolCol).Value
                End If
                If predVol > 0 Then
                    tbl.DataBodyRange.Cells(i, errVolCol).Value = CDbl(telemVol) - predVol
                End If
            End If
        Else
            ' No telemetry for this date - clear discrepancy
            If errECCol > 0 Then tbl.DataBodyRange.Cells(i, errECCol).ClearContents
            If errVolCol > 0 Then tbl.DataBodyRange.Cells(i, errVolCol).ClearContents
        End If
    Next i
End Sub

' ==== Row Lookup/Creation ===================================================

Private Function FindRowByDate(ByVal tbl As ListObject, ByVal targetDate As Date) As Long
    ' Returns row index (1-based) for date, or 0 if not found
    Dim i As Long
    If tbl.DataBodyRange Is Nothing Then Exit Function
    For i = 1 To tbl.ListRows.Count
        If CDate(tbl.DataBodyRange.Cells(i, 1).Value) = targetDate Then
            FindRowByDate = i
            Exit Function
        End If
    Next i
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
    Dim i As Long
    If tbl.DataBodyRange Is Nothing Then Exit Function
    For i = 1 To tbl.ListRows.Count
        If CDate(tbl.DataBodyRange.Cells(i, 1).Value) = targetDate Then
            FindTelemRowByDate = i
            Exit Function
        End If
    Next i
End Function

' ==== Delete Functions ======================================================

Public Sub DeleteAfterDate(ByVal cutoffDate As Date, ByVal site As String)
    ' Deletes all rows with Date > cutoffDate (for rollback)
    Dim tbl As ListObject
    Dim i As Long, rowDate As Date

    Set tbl = GetLiveTable(site)
    If tbl Is Nothing Then Exit Sub
    If tbl.DataBodyRange Is Nothing Then Exit Sub

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
    If tbl Is Nothing Then Exit Function
    If tbl.DataBodyRange Is Nothing Then Exit Function

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

    tblName = Schema.LiveTableName(site)

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

' ==== Legacy Support (to be removed in Phase 7) =============================

Public Sub DeleteRun(ByVal runId As String, ByVal site As String)
    ' Legacy: Deletes rows by RunId (no longer used - kept for compatibility)
    ' New architecture uses DeleteAfterDate for rollback
    Dim tbl As ListObject
    Dim i As Long, runIdCol As Long

    Set tbl = GetLiveTable(site)
    If tbl Is Nothing Then Exit Sub
    If tbl.DataBodyRange Is Nothing Then Exit Sub

    runIdCol = Schema.ColIdx(tbl, Schema.LIVE_COL_RUNID)
    If runIdCol = 0 Then Exit Sub

    ' Delete from bottom up
    For i = tbl.ListRows.Count To 1 Step -1
        If tbl.DataBodyRange.Cells(i, runIdCol).Value = runId Then
            tbl.ListRows(i).Delete
        End If
    Next i
End Sub
