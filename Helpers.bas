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

' ==== Action Cell Styling =======================================================

Public Sub StyleActionCell(ByVal cell As Range)
    ' Applies blue hyperlink style to action cells
    With cell
        .Font.Color = Schema.COLOR_ACTION_FONT
        .Font.Underline = xlUnderlineStyleSingle
    End With
End Sub

Public Sub InitIRRowAction(ByVal rowRng As Range, ByVal tbl As ListObject)
    ' Sets action cell value and styling only - no other formatting
    Dim actionCol As Long
    actionCol = ColIdx(tbl, Schema.IR_COL_ACTION)
    If actionCol > 0 Then
        rowRng.Cells(1, actionCol).Value = Schema.ACTION_REMOVE
        StyleActionCell rowRng.Cells(1, actionCol)
    End If
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
