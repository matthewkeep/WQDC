Option Explicit
' Telemetry: Access layer for telemetry data (rain, EC, volume).
' Dependencies: Schema
'
' All functions handle missing data gracefully - missing values don't break simulation.

' ==== Single Value Lookups ===================================================

Public Function GetRainForDate(ByVal d As Date) As Double
    ' Returns rainfall (mm) for date, or 0 if not found
    Dim v As Variant
    v = LookupValue(d, 2)
    If IsEmpty(v) Or IsError(v) Then
        GetRainForDate = 0
    Else
        GetRainForDate = CDbl(v)
    End If
End Function

' ==== Range Lookups ==========================================================

Public Function GetHindcastRain(ByVal startDate As Date, ByVal endDate As Date) As Double()
    ' Returns array of daily rainfall for date range (inclusive)
    ' Missing values default to 0
    Dim days As Long, i As Long
    Dim result() As Double

    days = endDate - startDate + 1
    If days < 1 Then
        ReDim result(0 To 0): result(0) = 0
        GetHindcastRain = result
        Exit Function
    End If

    ReDim result(0 To days - 1)
    For i = 0 To days - 1
        result(i) = GetRainForDate(startDate + i)
    Next i
    GetHindcastRain = result
End Function

Public Function GetLatestEC(ByVal beforeDate As Date, ByVal site As String) As Variant
    ' Returns most recent EC value on or before the given date for site
    GetLatestEC = GetLatestTelemValue(beforeDate, Helpers.TelemECColName(site))
End Function

Public Function GetLatestVol(ByVal beforeDate As Date, ByVal site As String) As Variant
    ' Returns most recent Volume value on or before the given date for site
    GetLatestVol = GetLatestTelemValue(beforeDate, Helpers.TelemVolColName(site))
End Function

Private Function GetLatestTelemValue(ByVal beforeDate As Date, ByVal colName As String) As Variant
    ' Returns most recent value on or before the given date for specified column
    ' Returns Empty if no data found
    ' Uses MATCH for O(1) lookup + backward scan for first non-empty
    Dim tbl As ListObject, i As Long
    Dim col As Long, rowIdx As Variant, v As Variant

    Set tbl = GetTelemTable()
    If Not Helpers.HasData(tbl) Then Exit Function

    col = Helpers.ColIdx(tbl, colName)
    If col = 0 Then Exit Function

    ' Use MATCH to find starting position (largest date <= beforeDate)
    rowIdx = Application.Match(CDbl(beforeDate), tbl.ListColumns(1).DataBodyRange, 1)
    If IsError(rowIdx) Then Exit Function

    ' Scan backwards from startIdx to find first non-empty value
    For i = CLng(rowIdx) To 1 Step -1
        v = tbl.DataBodyRange.Cells(i, col).Value
        If Not IsEmpty(v) Then
            GetLatestTelemValue = v
            Exit Function
        End If
    Next i
End Function

' ==== Aggregates =============================================================

Public Function GetTotalRain(ByVal startDate As Date, ByVal endDate As Date) As Double
    ' Returns total rainfall (mm) for date range
    Dim rain() As Double, i As Long, total As Double
    rain = GetHindcastRain(startDate, endDate)
    total = 0
    For i = LBound(rain) To UBound(rain)
        total = total + rain(i)
    Next i
    GetTotalRain = total
End Function

' ==== Private Helpers ========================================================

Private Function LookupValue(ByVal d As Date, ByVal col As Long) As Variant
    ' Looks up value in telemetry table by date and column index
    ' Uses MATCH for O(1) lookup instead of loop scan
    Dim tbl As ListObject, rowIdx As Variant

    Set tbl = GetTelemTable()
    If Not Helpers.HasData(tbl) Then LookupValue = Empty: Exit Function

    rowIdx = Application.Match(CDbl(d), tbl.ListColumns(1).DataBodyRange, 0)
    If IsError(rowIdx) Then
        LookupValue = Empty
    Else
        LookupValue = tbl.DataBodyRange.Cells(rowIdx, col).Value
    End If
End Function

Private Function GetTelemTable() As ListObject
    ' Returns tblTelemetry or Nothing if not found
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(Schema.SHEET_RESULTS)
    If Not ws Is Nothing Then
        Set GetTelemTable = ws.ListObjects(Schema.TABLE_TELEMETRY)
    End If
    On Error GoTo 0
End Function
