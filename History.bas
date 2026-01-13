Option Explicit
' History: Audit trail for simulation runs.
' Dependencies: Core, Schema

Public Sub RecordRun(ByRef cfg As Config, ByRef r As Result)
    Dim ws As Worksheet, tbl As ListObject, row As ListRow, id As String
    On Error Resume Next
    Set ws = GetHSheet(): If ws Is Nothing Then Exit Sub
    Set tbl = GetHTbl(ws): If tbl Is Nothing Then Exit Sub

    id = GenId()
    Set row = tbl.ListRows.Add: If row Is Nothing Then Exit Sub

    With row.Range
        .Cells(1, 1).Value = id
        .Cells(1, 2).Value = Now
        .Cells(1, 3).Value = cfg.StartDate
        .Cells(1, 4).Value = GetSite()
        .Cells(1, 5).Value = cfg.StartDate
        .Cells(1, 6).Value = cfg.Mode
        .Cells(1, 7).Value = r.TriggerDay
        .Cells(1, 8).Value = r.TriggerMetric
        .Cells(1, 9).Value = Schema.HISTORY_STATUS_ACTIVE
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
    Dim ws As Worksheet, tbl As ListObject, site As String, i As Long
    On Error Resume Next
    Set ws = GetHSheet(): If ws Is Nothing Then Exit Function
    Set tbl = GetHTbl(ws): If tbl Is Nothing Then Exit Function
    If tbl.ListRows.Count = 0 Then Exit Function

    site = GetSite()
    For i = tbl.ListRows.Count To 1 Step -1
        If tbl.ListRows(i).Range.Cells(1, 4).Value = site Then
            If tbl.ListRows(i).Range.Cells(1, 9).Value = Schema.HISTORY_STATUS_ACTIVE Then
                tbl.ListRows(i).Range.Cells(1, 9).Value = Schema.HISTORY_STATUS_ROLLEDBACK
                RollbackLast = True
                Exit Function
            End If
        End If
    Next i
    On Error GoTo 0
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

Private Function GenId() As String
    Randomize
    GenId = Format$(Now, "yyyymmdd_hhmmss") & "_" & Right$(Format$(Rnd() * 10000, "0000"), 4)
End Function
