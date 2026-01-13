Attribute VB_Name = "History"
Option Explicit
' History: Audit trail for simulation runs.
' Purpose: Record each run for review, learning, and rollback.
' Dependencies: Types, Schema

' ==== Record Run ==============================================================

' Record a completed run to history table
Public Sub RecordRun(ByRef cfg As Config, ByRef r As Result)
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim newRow As ListRow
    Dim runId As String

    On Error Resume Next

    Set ws = GetHistorySheet()
    If ws Is Nothing Then Exit Sub

    Set tbl = GetHistoryTable(ws)
    If tbl Is Nothing Then Exit Sub

    ' Generate unique run ID
    runId = GenerateRunId()

    ' Add new row
    Set newRow = tbl.ListRows.Add
    If newRow Is Nothing Then Exit Sub

    ' Populate row
    With newRow.Range
        .Cells(1, 1).Value = runId                          ' RunId
        .Cells(1, 2).Value = Now                            ' Timestamp
        .Cells(1, 3).Value = cfg.StartDate                  ' RunDate
        .Cells(1, 4).Value = GetSite()                      ' Site
        .Cells(1, 5).Value = cfg.StartDate                  ' SampleDate
        .Cells(1, 6).Value = cfg.Mode                       ' Mode
        .Cells(1, 7).Value = r.TriggerDay                   ' TriggerDay
        .Cells(1, 8).Value = r.TriggerMetric                ' TriggerMetric
        .Cells(1, 9).Value = Schema.HISTORY_STATUS_ACTIVE   ' Status
    End With

    On Error GoTo 0
End Sub

' ==== Query History ===========================================================

' Get the most recent run for current site
Public Function GetLastRun() As Variant
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim site As String
    Dim i As Long

    Set ws = GetHistorySheet()
    If ws Is Nothing Then Exit Function

    Set tbl = GetHistoryTable(ws)
    If tbl Is Nothing Then Exit Function
    If tbl.ListRows.Count = 0 Then Exit Function

    site = GetSite()

    ' Search from bottom (most recent first)
    For i = tbl.ListRows.Count To 1 Step -1
        If tbl.ListRows(i).Range.Cells(1, 4).Value = site Then
            If tbl.ListRows(i).Range.Cells(1, 9).Value = Schema.HISTORY_STATUS_ACTIVE Then
                GetLastRun = tbl.ListRows(i).Range.Value
                Exit Function
            End If
        End If
    Next i
End Function

' Count runs for current site
Public Function CountRuns() As Long
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim site As String
    Dim i As Long

    Set ws = GetHistorySheet()
    If ws Is Nothing Then Exit Function

    Set tbl = GetHistoryTable(ws)
    If tbl Is Nothing Then Exit Function

    site = GetSite()

    For i = 1 To tbl.ListRows.Count
        If tbl.ListRows(i).Range.Cells(1, 4).Value = site Then
            If tbl.ListRows(i).Range.Cells(1, 9).Value = Schema.HISTORY_STATUS_ACTIVE Then
                CountRuns = CountRuns + 1
            End If
        End If
    Next i
End Function

' ==== Rollback ================================================================

' Mark most recent run as rolled back
Public Function RollbackLast() As Boolean
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim site As String
    Dim i As Long

    On Error Resume Next

    Set ws = GetHistorySheet()
    If ws Is Nothing Then Exit Function

    Set tbl = GetHistoryTable(ws)
    If tbl Is Nothing Then Exit Function
    If tbl.ListRows.Count = 0 Then Exit Function

    site = GetSite()

    ' Find and mark most recent active run
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

' ==== Helper Functions ========================================================

Private Function GetHistorySheet() As Worksheet
    On Error Resume Next
    Set GetHistorySheet = ThisWorkbook.Worksheets(Schema.SHEET_HISTORY)
    On Error GoTo 0
End Function

Private Function GetHistoryTable(ByVal ws As Worksheet) As ListObject
    On Error Resume Next
    Set GetHistoryTable = ws.ListObjects(Schema.TABLE_HISTORY)
    On Error GoTo 0
End Function

Private Function GetSite() As String
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(Schema.SHEET_INPUT)
    If Not ws Is Nothing Then
        GetSite = CStr(ws.Range(Schema.NAME_SITE).Value)
    End If
    On Error GoTo 0
End Function

Private Function GenerateRunId() As String
    ' Simple unique ID: timestamp + random
    Randomize
    GenerateRunId = Format$(Now, "yyyymmdd_hhmmss") & "_" & Right$(Format$(Rnd() * 10000, "0000"), 4)
End Function
