Option Explicit
' Helpers: Utility functions for worksheet/table access and styling.
' Dependencies: Schema (constants only)

' ==== Table Column Access =======================================================

Public Function ColIdx(ByVal tbl As ListObject, ByVal colName As String) As Long
    ' Returns column index (1-based) or 0 if not found
    Dim col As ListColumn
    On Error Resume Next
    Set col = tbl.ListColumns(colName)
    If Not col Is Nothing Then ColIdx = col.Index
    On Error GoTo 0
End Function

' ==== Table Naming ==============================================================

Public Function LiveTableName(ByVal site As String) As String
    ' Returns table name for site's live log table (e.g., "tblLive_RP1")
    LiveTableName = Schema.LIVE_TABLE_PREFIX & site
End Function

Public Function HistoryTableName(ByVal site As String) As String
    ' Returns table name for site's history table (e.g., "tblHistory_RP1")
    HistoryTableName = Schema.HISTORY_TABLE_PREFIX & site
End Function

Public Function TelemECColName(ByVal site As String) As String
    ' Returns telemetry EC column name for site, e.g., "EC (RP1)"
    TelemECColName = "EC (" & site & ")"
End Function

Public Function TelemVolColName(ByVal site As String) As String
    ' Returns telemetry Volume column name for site, e.g., "Vol (RP1)"
    TelemVolColName = "Vol (" & site & ")"
End Function

Public Function SeasonLogTableName(ByVal site As String) As String
    ' Returns table name for site's season backtest table (e.g., "tblSeasonLog_RP1")
    SeasonLogTableName = "tblSeasonLog_" & site
End Function

' ==== Worksheet/Table Access ====================================================

Public Function GetSheet(ByVal nm As String) As Worksheet
    ' Returns worksheet by name, or Nothing if not found
    On Error Resume Next
    Set GetSheet = ThisWorkbook.Worksheets(nm)
    On Error GoTo 0
End Function

Public Function GetTable(ByVal sheetName As String, ByVal tableName As String) As ListObject
    ' Returns ListObject by sheet and table name, or Nothing if not found
    Dim ws As Worksheet
    Set ws = GetSheet(sheetName)
    If Not ws Is Nothing Then
        On Error Resume Next
        Set GetTable = ws.ListObjects(tableName)
        On Error GoTo 0
    End If
End Function

Public Function WithTableData(ByVal sheetName As String, ByVal tableName As String) As ListObject
    ' Returns table only if it exists AND has data rows, else Nothing
    ' Replaces the common pattern:
    '   Set tbl = GetTable(...): If tbl Is Nothing Then Exit
    '   If tbl.DataBodyRange Is Nothing Then Exit
    Dim tbl As ListObject
    Set tbl = GetTable(sheetName, tableName)
    If tbl Is Nothing Then Exit Function
    If tbl.DataBodyRange Is Nothing Then Exit Function
    Set WithTableData = tbl
End Function

Public Function HasData(ByVal tbl As ListObject) As Boolean
    ' Returns True if table exists and has data rows
    ' Use: If Not HasData(tbl) Then Exit Function
    If tbl Is Nothing Then Exit Function
    If tbl.DataBodyRange Is Nothing Then Exit Function
    HasData = True
End Function

Public Function MatchesSite(ByVal v As Variant, ByVal site As String) As Boolean
    ' Case-insensitive site comparison
    MatchesSite = (UCase$(Trim$(CStr(v))) = UCase$(Trim$(site)))
End Function

Public Function FindRowByDate(ByVal tbl As ListObject, ByVal targetDate As Date) As Long
    ' Returns row index (1-based) for date in first column, or 0 if not found
    ' Uses O(1) MATCH instead of O(n) loop
    Dim rowIdx As Variant
    If tbl Is Nothing Then Exit Function
    If tbl.DataBodyRange Is Nothing Then Exit Function
    rowIdx = Application.Match(CDbl(targetDate), tbl.ListColumns(1).DataBodyRange, 0)
    If Not IsError(rowIdx) Then FindRowByDate = CLng(rowIdx)
End Function

Public Sub InitIRRowAction(ByVal rowRng As Range, ByVal tbl As ListObject)
    ' Sets action cell value
    Dim actionCol As Long
    actionCol = ColIdx(tbl, Schema.IR_COL_ACTION)
    If actionCol > 0 Then rowRng.Cells(1, actionCol).Value = Schema.ACTION_REMOVE
End Sub

' ==== Named Range Access ========================================================

Public Function GetRng(ByVal ws As Worksheet, ByVal nm As String) As Range
    ' Returns named range on worksheet, or Nothing if not found
    On Error Resume Next
    Set GetRng = ws.Range(nm)
    On Error GoTo 0
End Function

Public Sub WriteToRange(ByVal ws As Worksheet, ByVal nm As String, ByVal v As Variant)
    ' Writes value to named range if it exists
    Dim rng As Range
    Set rng = GetRng(ws, nm)
    If Not rng Is Nothing Then rng.Value = v
End Sub

Public Function ReadFromRange(ByVal ws As Worksheet, ByVal nm As String) As Variant
    ' Reads value from named range, returns Empty if not found
    Dim rng As Range
    Set rng = GetRng(ws, nm)
    If Not rng Is Nothing Then ReadFromRange = rng.Value
End Function

Public Function GetDateVal(ByVal ws As Worksheet, ByVal nm As String) As Date
    ' Returns date value from named range, or 0 if invalid/empty
    Dim rng As Range, v As Variant
    Set rng = GetRng(ws, nm)
    If rng Is Nothing Then Exit Function
    v = rng.Value
    If IsEmpty(v) Then Exit Function
    If IsDate(v) Then
        GetDateVal = CDate(v)
    ElseIf IsNumeric(v) And v > 0 Then
        GetDateVal = CDate(v)  ' Excel serial date number
    End If
End Function

' ==== Serialization Helpers ====================================================

Public Function SerializeRange(ByVal rng As Range, ByVal count As Long) As String
    ' Serializes horizontal range values to pipe-delimited string
    Dim i As Long, parts() As String
    If rng Is Nothing Then Exit Function
    ReDim parts(0 To count - 1)
    For i = 1 To count
        parts(i - 1) = CStr(Val(rng.Cells(1, i).Value))
    Next i
    SerializeRange = Join(parts, "|")
End Function

Public Function SerializeColumn(ByVal rng As Range, ByVal count As Long) As String
    ' Serializes vertical column range values to pipe-delimited string
    Dim i As Long, parts() As String
    If rng Is Nothing Then Exit Function
    ReDim parts(0 To count - 1)
    For i = 1 To count
        parts(i - 1) = CStr(Val(rng.Cells(i, 1).Value))
    Next i
    SerializeColumn = Join(parts, "|")
End Function

Public Sub DeserializeToRange(ByVal str As String, ByVal rng As Range, ByVal count As Long)
    ' Writes pipe-delimited values to horizontal range
    Dim parts() As String, i As Long
    If rng Is Nothing Or Len(str) = 0 Then Exit Sub
    parts = Split(str, "|")
    For i = 1 To count
        If i - 1 <= UBound(parts) Then rng.Cells(1, i).Value = Val(parts(i - 1))
    Next i
End Sub

Public Sub DeserializeToColumn(ByVal str As String, ByVal rng As Range, ByVal count As Long)
    ' Writes pipe-delimited values to vertical column range
    Dim parts() As String, i As Long
    If rng Is Nothing Or Len(str) = 0 Then Exit Sub
    parts = Split(str, "|")
    For i = 1 To count
        If i - 1 <= UBound(parts) Then rng.Cells(i, 1).Value = Val(parts(i - 1))
    Next i
End Sub

Public Function SerializeIRTable(ByVal tbl As ListObject) As String
    ' Serializes IR table to multi-line pipe-delimited string
    ' Format: Source|Flow|Active|SampleDate|EC|F_U|F_Mn|SO4|Mg|Ca|TAN
    Dim row As ListRow, lines() As String, lineIdx As Long
    Dim srcCol As Long, flowCol As Long, activeCol As Long, dateCol As Long
    Dim chemNames As Variant, i As Long, parts() As String

    If tbl Is Nothing Then Exit Function
    If tbl.DataBodyRange Is Nothing Then Exit Function

    srcCol = ColIdx(tbl, Schema.IR_COL_SOURCE)
    flowCol = ColIdx(tbl, Schema.IR_COL_FLOW)
    activeCol = ColIdx(tbl, Schema.IR_COL_ACTIVE)
    dateCol = ColIdx(tbl, Schema.IR_COL_SAMPLE_DATE)
    chemNames = Schema.ChemistryNames()

    ReDim lines(0 To tbl.ListRows.Count - 1)
    lineIdx = 0

    For Each row In tbl.ListRows
        ReDim parts(0 To 3 + Core.METRIC_COUNT)
        parts(0) = CStr(row.Range.Cells(1, srcCol).Value)
        parts(1) = CStr(Val(row.Range.Cells(1, flowCol).Value))
        parts(2) = CStr(row.Range.Cells(1, activeCol).Value)
        parts(3) = Format$(row.Range.Cells(1, dateCol).Value, "yyyy-mm-dd")
        For i = 1 To Core.METRIC_COUNT
            parts(3 + i) = CStr(Val(row.Range.Cells(1, ColIdx(tbl, chemNames(i - 1))).Value))
        Next i
        lines(lineIdx) = Join(parts, "|")
        lineIdx = lineIdx + 1
    Next row

    SerializeIRTable = Join(lines, vbLf)
End Function

Public Sub DeserializeIRTable(ByVal str As String, ByVal tbl As ListObject)
    ' Restores IR table from multi-line pipe-delimited string
    Dim lines() As String, parts() As String, i As Long, j As Long
    Dim srcCol As Long, flowCol As Long, activeCol As Long, dateCol As Long
    Dim chemNames As Variant, row As ListRow

    If tbl Is Nothing Or Len(str) = 0 Then Exit Sub

    ' Clear existing rows
    If Not tbl.DataBodyRange Is Nothing Then tbl.DataBodyRange.Delete

    srcCol = ColIdx(tbl, Schema.IR_COL_SOURCE)
    flowCol = ColIdx(tbl, Schema.IR_COL_FLOW)
    activeCol = ColIdx(tbl, Schema.IR_COL_ACTIVE)
    dateCol = ColIdx(tbl, Schema.IR_COL_SAMPLE_DATE)
    chemNames = Schema.ChemistryNames()

    lines = Split(str, vbLf)

    For i = 0 To UBound(lines)
        If Len(Trim$(lines(i))) > 0 Then
            parts = Split(lines(i), "|")
            Set row = tbl.ListRows.Add
            If srcCol > 0 Then row.Range.Cells(1, srcCol).Value = parts(0)
            If flowCol > 0 And UBound(parts) >= 1 Then row.Range.Cells(1, flowCol).Value = Val(parts(1))
            If activeCol > 0 And UBound(parts) >= 2 Then row.Range.Cells(1, activeCol).Value = parts(2)
            If dateCol > 0 And UBound(parts) >= 3 Then row.Range.Cells(1, dateCol).Value = CDate(parts(3))
            For j = 1 To Core.METRIC_COUNT
                If UBound(parts) >= 3 + j Then
                    row.Range.Cells(1, ColIdx(tbl, chemNames(j - 1))).Value = Val(parts(3 + j))
                End If
            Next j
            InitIRRowAction row.Range, tbl
        End If
    Next i
End Sub
