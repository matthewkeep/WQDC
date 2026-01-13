Option Explicit
' History: Audit trail for simulation runs.
' Dependencies: Core, Schema, SimLog

Public Sub RecordRun(ByRef cfg As Config, ByRef r As Result, ByVal runId As String)
    ' Records run metadata to history table. RunId must match SimLog entry.
    Dim ws As Worksheet, tbl As ListObject, row As ListRow, i As Long
    On Error Resume Next
    Set ws = GetHSheet(): If ws Is Nothing Then Exit Sub
    Set tbl = GetHTbl(ws): If tbl Is Nothing Then Exit Sub

    ' Update existing rows' action to "Rollback"
    If Not tbl.DataBodyRange Is Nothing Then
        For i = 1 To tbl.ListRows.Count
            tbl.DataBodyRange.Cells(i, 10).Value = Schema.ACTION_ROLLBACK
            StyleActionCell tbl.DataBodyRange.Cells(i, 10)
        Next i
    End If

    Set row = tbl.ListRows.Add: If row Is Nothing Then Exit Sub

    With row.Range
        .Cells(1, 1).Value = runId
        .Cells(1, 2).Value = Now
        .Cells(1, 3).Value = cfg.StartDate
        .Cells(1, 4).Value = GetSite()
        .Cells(1, 5).Value = cfg.Days
        .Cells(1, 6).Value = cfg.Mode
        .Cells(1, 7).Value = r.TriggerDay
        .Cells(1, 8).Value = r.TriggerMetric
        .Cells(1, 9).Value = Schema.HISTORY_STATUS_ACTIVE
        .Cells(1, 10).Value = Schema.ACTION_CURRENT
        StyleActionCell .Cells(1, 10)
    End With
    On Error GoTo 0
End Sub

Public Function GetLastRun() As Variant
    Dim ws As Worksheet, tbl As ListObject, site As String, i As Long
    Set ws = GetHSheet(): If ws Is Nothing Then Exit Function
    Set tbl = GetHTbl(ws): If tbl Is Nothing Then Exit Function
    If tbl.ListRows.Count = 0 Then Exit Function

    site = GetSite()
    For i = tbl.ListRows.Count To 1 Step -1
        If tbl.ListRows(i).Range.Cells(1, 4).Value = site Then
            If tbl.ListRows(i).Range.Cells(1, 9).Value = Schema.HISTORY_STATUS_ACTIVE Then
                GetLastRun = tbl.ListRows(i).Range.Value
                Exit Function
            End If
        End If
    Next i
End Function

Public Function CountRuns() As Long
    Dim ws As Worksheet, tbl As ListObject, site As String, i As Long
    Set ws = GetHSheet(): If ws Is Nothing Then Exit Function
    Set tbl = GetHTbl(ws): If tbl Is Nothing Then Exit Function

    site = GetSite()
    For i = 1 To tbl.ListRows.Count
        If tbl.ListRows(i).Range.Cells(1, 4).Value = site Then
            If tbl.ListRows(i).Range.Cells(1, 9).Value = Schema.HISTORY_STATUS_ACTIVE Then
                CountRuns = CountRuns + 1
            End If
        End If
    Next i
End Function

Public Function RollbackLast() As Boolean
    ' Marks last run as rolled back AND deletes its SimLog entries
    Dim ws As Worksheet, tbl As ListObject, site As String, runId As String, i As Long
    On Error Resume Next
    Set ws = GetHSheet(): If ws Is Nothing Then Exit Function
    Set tbl = GetHTbl(ws): If tbl Is Nothing Then Exit Function
    If tbl.ListRows.Count = 0 Then Exit Function

    site = GetSite()
    For i = tbl.ListRows.Count To 1 Step -1
        If tbl.ListRows(i).Range.Cells(1, 4).Value = site Then
            If tbl.ListRows(i).Range.Cells(1, 9).Value = Schema.HISTORY_STATUS_ACTIVE Then
                runId = tbl.ListRows(i).Range.Cells(1, 1).Value
                tbl.ListRows(i).Range.Cells(1, 9).Value = Schema.HISTORY_STATUS_ROLLEDBACK
                ' Delete corresponding SimLog entries
                SimLog.DeleteRun runId
                RollbackLast = True
                Exit Function
            End If
        End If
    Next i
    On Error GoTo 0
End Function

Public Function RollbackTo(ByVal targetRunId As String) As Long
    ' Pops all runs AFTER targetRunId for current site (Jenga model)
    ' Returns count of runs removed
    Dim ws As Worksheet, tbl As ListObject, site As String
    Dim runId As String, i As Long, foundTarget As Boolean, removed As Long

    On Error Resume Next
    Set ws = GetHSheet(): If ws Is Nothing Then Exit Function
    Set tbl = GetHTbl(ws): If tbl Is Nothing Then Exit Function
    If tbl.ListRows.Count = 0 Then Exit Function

    site = GetSite()

    ' First pass: find the target run to verify it exists
    For i = 1 To tbl.ListRows.Count
        If tbl.ListRows(i).Range.Cells(1, 4).Value = site Then
            If tbl.ListRows(i).Range.Cells(1, 1).Value = targetRunId Then
                foundTarget = True
                Exit For
            End If
        End If
    Next i
    If Not foundTarget Then Exit Function

    ' Second pass: remove all ACTIVE runs for this site that come AFTER target
    ' Work backwards from end to avoid index issues
    For i = tbl.ListRows.Count To 1 Step -1
        If tbl.ListRows(i).Range.Cells(1, 4).Value = site Then
            runId = tbl.ListRows(i).Range.Cells(1, 1).Value
            If runId = targetRunId Then Exit For  ' Stop at target
            If tbl.ListRows(i).Range.Cells(1, 9).Value = Schema.HISTORY_STATUS_ACTIVE Then
                tbl.ListRows(i).Range.Cells(1, 9).Value = Schema.HISTORY_STATUS_ROLLEDBACK
                SimLog.DeleteRun runId
                removed = removed + 1
            End If
        End If
    Next i

    RollbackTo = removed
    On Error GoTo 0
End Function

Public Function GetRunHistory() As Variant
    ' Returns array of active runs for current site (for display/recall)
    ' Each row: (RunId, Timestamp, StartDate, TriggerDay, TriggerMetric)
    Dim ws As Worksheet, tbl As ListObject, site As String
    Dim result() As Variant, cnt As Long, i As Long

    Set ws = GetHSheet(): If ws Is Nothing Then Exit Function
    Set tbl = GetHTbl(ws): If tbl Is Nothing Then Exit Function
    If tbl.ListRows.Count = 0 Then Exit Function

    site = GetSite()

    ' Count active runs for this site
    cnt = 0
    For i = 1 To tbl.ListRows.Count
        If tbl.ListRows(i).Range.Cells(1, 4).Value = site Then
            If tbl.ListRows(i).Range.Cells(1, 9).Value = Schema.HISTORY_STATUS_ACTIVE Then
                cnt = cnt + 1
            End If
        End If
    Next i
    If cnt = 0 Then Exit Function

    ' Build result array
    ReDim result(1 To cnt, 1 To 5)
    cnt = 0
    For i = 1 To tbl.ListRows.Count
        If tbl.ListRows(i).Range.Cells(1, 4).Value = site Then
            If tbl.ListRows(i).Range.Cells(1, 9).Value = Schema.HISTORY_STATUS_ACTIVE Then
                cnt = cnt + 1
                result(cnt, 1) = tbl.ListRows(i).Range.Cells(1, 1).Value  ' RunId
                result(cnt, 2) = tbl.ListRows(i).Range.Cells(1, 2).Value  ' Timestamp
                result(cnt, 3) = tbl.ListRows(i).Range.Cells(1, 3).Value  ' StartDate
                result(cnt, 4) = tbl.ListRows(i).Range.Cells(1, 7).Value  ' TriggerDay
                result(cnt, 5) = tbl.ListRows(i).Range.Cells(1, 8).Value  ' TriggerMetric
            End If
        End If
    Next i

    GetRunHistory = result
End Function

' ==== Helpers ================================================================

Private Function GetHSheet() As Worksheet
    On Error Resume Next
    Set GetHSheet = ThisWorkbook.Worksheets(Schema.SHEET_HISTORY)
    On Error GoTo 0
End Function

Private Function GetHTbl(ByVal ws As Worksheet) As ListObject
    On Error Resume Next
    Set GetHTbl = ws.ListObjects(Schema.TABLE_HISTORY)
    On Error GoTo 0
End Function

Private Function GetSite() As String
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(Schema.SHEET_INPUT)
    If Not ws Is Nothing Then GetSite = CStr(ws.Range(Schema.NAME_SITE).Value)
    On Error GoTo 0
End Function

Private Sub StyleActionCell(ByVal cell As Range)
    With cell
        .Font.Color = Schema.COLOR_ACTION_FONT
        .Font.Underline = xlUnderlineStyleSingle
    End With
End Sub

