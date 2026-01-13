Option Explicit
' SimLog: Persistent daily simulation output storage.
' Dependencies: Core, Schema
'
' All runs are stored with RunId prefix for history/recall.

' ==== Write Functions =======================================================

Public Sub WriteLog(ByRef r As Result, ByRef cfg As Config, ByVal runId As String)
    ' Clears data from StartDate onwards, then appends new snapshots
    Dim ws As Worksheet, tbl As ListObject
    Dim i As Long, j As Long, n As Long, newRow As ListRow

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(Schema.SHEET_LOG)
    On Error GoTo 0
    If ws Is Nothing Then Exit Sub

    Set tbl = ws.ListObjects(Schema.TABLE_LOG_DAILY)
    If tbl Is Nothing Then Exit Sub

    ' Clear existing data from StartDate onwards (keeps historical, replaces future)
    DeleteFromDate cfg.StartDate

    n = UBound(r.Snaps)
    For i = 0 To n
        Set newRow = tbl.ListRows.Add
        With newRow.Range
            .Cells(1, 1) = runId                        ' RunId
            .Cells(1, 2) = cfg.StartDate + i            ' Date
            .Cells(1, 3) = i                            ' Day
            .Cells(1, 4) = r.Snaps(i).Vol              ' Volume
            For j = 1 To Core.METRIC_COUNT
                .Cells(1, 4 + j) = r.Snaps(i).Chem(j)  ' Chemistry
            Next j
        End With
    Next i
End Sub

Public Sub DeleteRun(ByVal runId As String)
    ' Deletes all rows for a specific RunId (for true rollback)
    Dim ws As Worksheet, tbl As ListObject
    Dim i As Long

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(Schema.SHEET_LOG)
    On Error GoTo 0
    If ws Is Nothing Then Exit Sub

    Set tbl = ws.ListObjects(Schema.TABLE_LOG_DAILY)
    If tbl Is Nothing Then Exit Sub
    If tbl.DataBodyRange Is Nothing Then Exit Sub

    ' Delete from bottom up to avoid index issues
    For i = tbl.ListRows.Count To 1 Step -1
        If tbl.DataBodyRange.Cells(i, 1).Value = runId Then
            tbl.ListRows(i).Delete
        End If
    Next i
End Sub

Public Sub ClearAllLogs()
    ' Clears entire log table (use with caution)
    Dim ws As Worksheet, tbl As ListObject

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(Schema.SHEET_LOG)
    On Error GoTo 0
    If ws Is Nothing Then Exit Sub

    Set tbl = ws.ListObjects(Schema.TABLE_LOG_DAILY)
    If Not tbl Is Nothing Then
        If Not tbl.DataBodyRange Is Nothing Then tbl.DataBodyRange.Delete
    End If
End Sub

' ==== Read Functions ========================================================

Public Function GetRunSnapshots(ByVal runId As String) As State()
    ' Returns array of State snapshots for a specific RunId
    Dim ws As Worksheet, tbl As ListObject
    Dim snaps() As State, tempSnaps() As State
    Dim i As Long, j As Long, cnt As Long

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(Schema.SHEET_LOG)
    On Error GoTo 0
    If ws Is Nothing Then Exit Function

    Set tbl = ws.ListObjects(Schema.TABLE_LOG_DAILY)
    If tbl Is Nothing Then Exit Function
    If tbl.DataBodyRange Is Nothing Then Exit Function

    ' Count matching rows
    cnt = 0
    For i = 1 To tbl.ListRows.Count
        If tbl.DataBodyRange.Cells(i, 1).Value = runId Then cnt = cnt + 1
    Next i
    If cnt = 0 Then Exit Function

    ' Extract snapshots
    ReDim snaps(0 To cnt - 1)
    cnt = 0
    For i = 1 To tbl.ListRows.Count
        If tbl.DataBodyRange.Cells(i, 1).Value = runId Then
            snaps(cnt).Vol = tbl.DataBodyRange.Cells(i, 4)
            For j = 1 To Core.METRIC_COUNT
                snaps(cnt).Chem(j) = tbl.DataBodyRange.Cells(i, 4 + j)
            Next j
            cnt = cnt + 1
        End If
    Next i

    GetRunSnapshots = snaps
End Function

Public Function GetAllRunIds() As Variant
    ' Returns array of unique RunIds in the log
    Dim ws As Worksheet, tbl As ListObject
    Dim dict As Object, i As Long, runId As String
    Dim result() As String, cnt As Long

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(Schema.SHEET_LOG)
    On Error GoTo 0
    If ws Is Nothing Then Exit Function

    Set tbl = ws.ListObjects(Schema.TABLE_LOG_DAILY)
    If tbl Is Nothing Then Exit Function
    If tbl.DataBodyRange Is Nothing Then Exit Function

    ' Use dictionary for unique values (DictionaryShim for Mac compatibility)
    Set dict = New DictionaryShim
    For i = 1 To tbl.ListRows.Count
        runId = tbl.DataBodyRange.Cells(i, 1).Value
        If Len(runId) > 0 And Not dict.Exists(runId) Then
            dict.Add runId, True
        End If
    Next i

    If dict.Count = 0 Then Exit Function

    GetAllRunIds = dict.Keys
End Function

Public Function GetRunCount() As Long
    ' Returns count of unique runs in the log
    Dim ids As Variant
    ids = GetAllRunIds()
    If IsArray(ids) Then
        GetRunCount = UBound(ids) - LBound(ids) + 1
    End If
End Function

Public Function GetLatestLogDate() As Date
    ' Returns the most recent date in tblLogDaily (0 if empty)
    Dim ws As Worksheet, tbl As ListObject
    Dim i As Long, d As Date, maxDate As Date

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(Schema.SHEET_LOG)
    On Error GoTo 0
    If ws Is Nothing Then Exit Function

    Set tbl = ws.ListObjects(Schema.TABLE_LOG_DAILY)
    If tbl Is Nothing Then Exit Function
    If tbl.DataBodyRange Is Nothing Then Exit Function

    maxDate = 0
    For i = 1 To tbl.ListRows.Count
        d = tbl.DataBodyRange.Cells(i, 2).Value  ' Date column
        If d > maxDate Then maxDate = d
    Next i
    GetLatestLogDate = maxDate
End Function

' ==== Private Helpers =========================================================

Private Sub DeleteFromDate(ByVal startDate As Date)
    ' Deletes all rows where Date >= startDate
    Dim ws As Worksheet, tbl As ListObject
    Dim i As Long, d As Date

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(Schema.SHEET_LOG)
    On Error GoTo 0
    If ws Is Nothing Then Exit Sub

    Set tbl = ws.ListObjects(Schema.TABLE_LOG_DAILY)
    If tbl Is Nothing Then Exit Sub
    If tbl.DataBodyRange Is Nothing Then Exit Sub

    ' Delete from bottom up to avoid index issues
    For i = tbl.ListRows.Count To 1 Step -1
        d = tbl.DataBodyRange.Cells(i, 2).Value  ' Date column
        If d >= startDate Then
            tbl.ListRows(i).Delete
        End If
    Next i
End Sub
