Option Explicit
' SimLog: Persistent daily simulation output storage.
' Dependencies: Core, Schema, Setup (for EnsureSiteLogTable)
'
' All runs are stored per-site with RunId prefix for history/recall.
' Tables are created on-demand: tblLog_RP1, tblLog_RP2, etc.

' ==== Write Functions =======================================================

Public Sub WriteLog(ByRef r As Result, ByRef cfg As Config, ByVal runId As String, ByVal site As String)
    ' Appends new snapshots to site's log table
    ' Day 0 (sample date) is shaded for identification
    Dim tbl As ListObject
    Dim i As Long, j As Long, n As Long, newRow As ListRow

    Set tbl = GetLogTable(site)
    If tbl Is Nothing Then Exit Sub

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

            ' Shade Day 0 (sample/start date) for identification
            If i = 0 Then
                .Interior.Color = Schema.COLOR_SAMPLE_DATE
            End If
        End With
    Next i
End Sub

Public Sub DeleteRun(ByVal runId As String, ByVal site As String)
    ' Deletes all rows for a specific RunId from site's log table
    Dim tbl As ListObject
    Dim i As Long

    Set tbl = GetLogTable(site)
    If tbl Is Nothing Then Exit Sub
    If tbl.DataBodyRange Is Nothing Then Exit Sub

    ' Delete from bottom up to avoid index issues
    For i = tbl.ListRows.Count To 1 Step -1
        If tbl.DataBodyRange.Cells(i, 1).Value = runId Then
            tbl.ListRows(i).Delete
        End If
    Next i
End Sub

Public Sub ClearSiteLog(ByVal site As String)
    ' Clears entire log table for site
    Dim tbl As ListObject

    Set tbl = GetLogTable(site)
    If tbl Is Nothing Then Exit Sub
    If Not tbl.DataBodyRange Is Nothing Then tbl.DataBodyRange.Delete
End Sub

' ==== Read Functions ========================================================

Public Function GetLatestLogDate(ByVal site As String) As Date
    ' Returns the most recent date in site's log (0 if empty)
    Dim tbl As ListObject
    Dim i As Long, d As Date, maxDate As Date

    Set tbl = GetLogTable(site)
    If tbl Is Nothing Then Exit Function
    If tbl.DataBodyRange Is Nothing Then Exit Function

    maxDate = 0
    For i = 1 To tbl.ListRows.Count
        d = tbl.DataBodyRange.Cells(i, 2).Value  ' Date column
        If d > maxDate Then maxDate = d
    Next i
    GetLatestLogDate = maxDate
End Function

' ==== Table Access ===========================================================

Private Function GetLogTable(ByVal site As String) As ListObject
    ' Returns site's log table, creating it if necessary
    Dim ws As Worksheet, tblName As String

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(Schema.SHEET_LOG)
    On Error GoTo 0
    If ws Is Nothing Then Exit Function

    tblName = Schema.LogTableName(site)

    ' Try to get existing table
    On Error Resume Next
    Set GetLogTable = ws.ListObjects(tblName)
    On Error GoTo 0

    ' Create if doesn't exist
    If GetLogTable Is Nothing Then
        Setup.EnsureSiteLogTable site
        On Error Resume Next
        Set GetLogTable = ws.ListObjects(tblName)
        On Error GoTo 0
    End If
End Function
