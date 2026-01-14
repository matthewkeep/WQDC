Option Explicit
' History: Audit trail for simulation runs.
' Dependencies: Core, Schema, SimLog, Setup (for EnsureSiteHistoryTable)
'
' All runs are stored per-site. Tables created on-demand: tblHistory_RP1, etc.
' No Site column in table - site is encoded in table name.

Public Sub RecordRun(ByRef cfg As Config, ByRef r As Result, ByVal runId As String, ByVal site As String)
    ' Records run metadata to site's history table. RunId must match SimLog entry.
    Dim tbl As ListObject, row As ListRow, i As Long, actionCol As Long

    Set tbl = GetHistoryTable(site)
    If tbl Is Nothing Then Exit Sub

    actionCol = Schema.ColIdx(tbl, Schema.HISTORY_COL_ACTION)
    If actionCol = 0 Then Exit Sub

    ' Update existing rows' action to "Rollback"
    If Not tbl.DataBodyRange Is Nothing Then
        For i = 1 To tbl.ListRows.Count
            tbl.DataBodyRange.Cells(i, actionCol).Value = Schema.ACTION_ROLLBACK
            Schema.StyleActionCell tbl.DataBodyRange.Cells(i, actionCol)
        Next i
    End If

    Set row = tbl.ListRows.Add: If row Is Nothing Then Exit Sub

    With row.Range
        .Cells(1, 1).Value = runId
        .Cells(1, 2).Value = Now
        .Cells(1, 3).Value = cfg.StartDate
        .Cells(1, 4).Value = cfg.Days
        .Cells(1, 5).Value = cfg.Mode
        .Cells(1, 6).Value = r.TriggerDay
        .Cells(1, 7).Value = r.TriggerMetric
        .Cells(1, actionCol).Value = Schema.ACTION_CURRENT
        Schema.StyleActionCell .Cells(1, actionCol)
    End With
End Sub

Public Function GetLastRun(ByVal site As String) As Variant
    ' Returns last run's row data for site
    Dim tbl As ListObject

    Set tbl = GetHistoryTable(site)
    If tbl Is Nothing Then Exit Function
    If tbl.ListRows.Count = 0 Then Exit Function

    GetLastRun = tbl.ListRows(tbl.ListRows.Count).Range.Value
End Function

Public Function CountRuns(ByVal site As String) As Long
    ' Returns count of runs for site
    Dim tbl As ListObject

    Set tbl = GetHistoryTable(site)
    If tbl Is Nothing Then Exit Function

    CountRuns = tbl.ListRows.Count
End Function

Public Function RollbackLast(ByVal site As String) As Boolean
    ' Deletes last run from history AND log entries after that run's start date
    Dim tbl As ListObject, startDate As Date

    Set tbl = GetHistoryTable(site)
    If tbl Is Nothing Then Exit Function
    If tbl.ListRows.Count = 0 Then Exit Function

    ' Get start date of last run (column 3 = RunDate)
    startDate = tbl.ListRows(tbl.ListRows.Count).Range.Cells(1, 3).Value

    ' Delete log entries after the previous run's start date
    If tbl.ListRows.Count > 1 Then
        ' Roll back to previous run's end date
        Dim prevStartDate As Date
        prevStartDate = tbl.ListRows(tbl.ListRows.Count - 1).Range.Cells(1, 3).Value
        SimLog.DeleteAfterDate prevStartDate, site
    Else
        ' Last run - delete all log entries before this run
        SimLog.DeleteAfterDate startDate - 1, site
    End If

    tbl.ListRows(tbl.ListRows.Count).Delete

    ' Update new last row to Current
    If tbl.ListRows.Count > 0 Then
        Dim actionCol As Long
        actionCol = Schema.ColIdx(tbl, Schema.HISTORY_COL_ACTION)
        If actionCol > 0 Then
            tbl.DataBodyRange.Cells(tbl.ListRows.Count, actionCol).Value = Schema.ACTION_CURRENT
            Schema.StyleActionCell tbl.DataBodyRange.Cells(tbl.ListRows.Count, actionCol)
        End If
    End If

    RollbackLast = True
End Function

Public Function RollbackTo(ByVal targetRunId As String, ByVal site As String) As Long
    ' Deletes all runs AFTER targetRunId for site (Jenga model)
    ' Returns count of runs removed
    Dim tbl As ListObject
    Dim runId As String, i As Long, targetIdx As Long, removed As Long
    Dim targetStartDate As Date

    Set tbl = GetHistoryTable(site)
    If tbl Is Nothing Then Exit Function
    If tbl.ListRows.Count = 0 Then Exit Function

    ' Find the target run to get its start date
    targetIdx = 0
    For i = 1 To tbl.ListRows.Count
        If tbl.ListRows(i).Range.Cells(1, 1).Value = targetRunId Then
            targetIdx = i
            targetStartDate = tbl.ListRows(i).Range.Cells(1, 3).Value  ' RunDate column
            Exit For
        End If
    Next i
    If targetIdx = 0 Then Exit Function

    ' Delete log entries after target run's start date
    SimLog.DeleteAfterDate targetStartDate, site

    ' Delete all history rows that come AFTER target
    ' Work backwards from end to avoid index issues
    For i = tbl.ListRows.Count To targetIdx + 1 Step -1
        tbl.ListRows(i).Delete
        removed = removed + 1
    Next i

    ' Update target row to Current
    If tbl.ListRows.Count > 0 Then
        Dim actionCol As Long
        actionCol = Schema.ColIdx(tbl, Schema.HISTORY_COL_ACTION)
        If actionCol > 0 Then
            tbl.DataBodyRange.Cells(tbl.ListRows.Count, actionCol).Value = Schema.ACTION_CURRENT
            Schema.StyleActionCell tbl.DataBodyRange.Cells(tbl.ListRows.Count, actionCol)
        End If
    End If

    RollbackTo = removed
End Function

Public Function GetRunHistory(ByVal site As String) As Variant
    ' Returns array of runs for site (for display/recall)
    ' Each row: (RunId, Timestamp, StartDate, TriggerDay, TriggerMetric)
    Dim tbl As ListObject
    Dim result() As Variant, i As Long

    Set tbl = GetHistoryTable(site)
    If tbl Is Nothing Then Exit Function
    If tbl.ListRows.Count = 0 Then Exit Function

    ' Build result array
    ReDim result(1 To tbl.ListRows.Count, 1 To 5)
    For i = 1 To tbl.ListRows.Count
        result(i, 1) = tbl.ListRows(i).Range.Cells(1, 1).Value  ' RunId
        result(i, 2) = tbl.ListRows(i).Range.Cells(1, 2).Value  ' Timestamp
        result(i, 3) = tbl.ListRows(i).Range.Cells(1, 3).Value  ' StartDate
        result(i, 4) = tbl.ListRows(i).Range.Cells(1, 6).Value  ' TriggerDay
        result(i, 5) = tbl.ListRows(i).Range.Cells(1, 7).Value  ' TriggerMetric
    Next i

    GetRunHistory = result
End Function

' ==== Table Access ===========================================================

Private Function GetHistoryTable(ByVal site As String) As ListObject
    ' Returns site's history table, creating it if necessary
    Dim ws As Worksheet, tblName As String

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(Schema.SHEET_HISTORY)
    On Error GoTo 0
    If ws Is Nothing Then Exit Function

    tblName = Schema.HistoryTableName(site)

    ' Try to get existing table
    On Error Resume Next
    Set GetHistoryTable = ws.ListObjects(tblName)
    On Error GoTo 0

    ' Create if doesn't exist
    If GetHistoryTable Is Nothing Then
        Setup.EnsureSiteHistoryTable site
        On Error Resume Next
        Set GetHistoryTable = ws.ListObjects(tblName)
        On Error GoTo 0
    End If
End Function

