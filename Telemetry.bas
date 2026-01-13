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

Public Function GetECForDate(ByVal d As Date, ByVal site As String) As Variant
    ' Returns EC (uS/cm) for date and site, or Empty if not found
    GetECForDate = LookupValueByColName(d, Schema.TelemECColName(site))
End Function

Public Function GetVolForDate(ByVal d As Date, ByVal site As String) As Variant
    ' Returns Volume (ML) for date and site, or Empty if not found
    GetVolForDate = LookupValueByColName(d, Schema.TelemVolColName(site))
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

Public Function GetHindcastEC(ByVal startDate As Date, ByVal endDate As Date, ByVal site As String) As Variant()
    ' Returns array of daily EC for date range (inclusive) for site
    ' Missing values are Empty
    Dim days As Long, i As Long
    Dim result() As Variant

    days = endDate - startDate + 1
    If days < 1 Then
        ReDim result(0 To 0): result(0) = Empty
        GetHindcastEC = result
        Exit Function
    End If

    ReDim result(0 To days - 1)
    For i = 0 To days - 1
        result(i) = GetECForDate(startDate + i, site)
    Next i
    GetHindcastEC = result
End Function

Public Function GetLatestEC(ByVal beforeDate As Date, ByVal site As String) As Variant
    ' Returns most recent EC value on or before the given date for site
    ' Returns Empty if no data found
    Dim tbl As ListObject, i As Long, d As Date, ec As Variant
    Dim bestDate As Date, bestEC As Variant
    Dim ecCol As Long

    Set tbl = GetTelemTable()
    If tbl Is Nothing Then Exit Function
    If tbl.DataBodyRange Is Nothing Then Exit Function

    ecCol = GetColIndex(tbl, Schema.TelemECColName(site))
    If ecCol = 0 Then Exit Function

    bestDate = 0: bestEC = Empty
    For i = 1 To tbl.ListRows.Count
        d = tbl.DataBodyRange.Cells(i, 1).Value
        If d <= beforeDate And d > bestDate Then
            ec = tbl.DataBodyRange.Cells(i, ecCol).Value
            If Not IsEmpty(ec) Then
                bestDate = d
                bestEC = ec
            End If
        End If
    Next i
    GetLatestEC = bestEC
End Function

Public Function GetLatestVol(ByVal beforeDate As Date, ByVal site As String) As Variant
    ' Returns most recent Volume value on or before the given date for site
    ' Returns Empty if no data found
    Dim tbl As ListObject, i As Long, d As Date, v As Variant
    Dim bestDate As Date, bestVol As Variant
    Dim volCol As Long

    Set tbl = GetTelemTable()
    If tbl Is Nothing Then Exit Function
    If tbl.DataBodyRange Is Nothing Then Exit Function

    volCol = GetColIndex(tbl, Schema.TelemVolColName(site))
    If volCol = 0 Then Exit Function

    bestDate = 0: bestVol = Empty
    For i = 1 To tbl.ListRows.Count
        d = tbl.DataBodyRange.Cells(i, 1).Value
        If d <= beforeDate And d > bestDate Then
            v = tbl.DataBodyRange.Cells(i, volCol).Value
            If Not IsEmpty(v) Then
                bestDate = d
                bestVol = v
            End If
        End If
    Next i
    GetLatestVol = bestVol
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
    ' Returns cell value or Empty if not found
    Dim tbl As ListObject, i As Long, rowDate As Date

    Set tbl = GetTelemTable()
    If tbl Is Nothing Then
        LookupValue = Empty
        Exit Function
    End If
    If tbl.DataBodyRange Is Nothing Then
        LookupValue = Empty
        Exit Function
    End If

    For i = 1 To tbl.ListRows.Count
        rowDate = tbl.DataBodyRange.Cells(i, 1).Value
        If rowDate = d Then
            LookupValue = tbl.DataBodyRange.Cells(i, col).Value
            Exit Function
        End If
    Next i
    LookupValue = Empty
End Function

Private Function LookupValueByColName(ByVal d As Date, ByVal colName As String) As Variant
    ' Looks up value in telemetry table by date and column name
    ' Returns cell value or Empty if not found
    Dim tbl As ListObject, colIdx As Long

    Set tbl = GetTelemTable()
    If tbl Is Nothing Then
        LookupValueByColName = Empty
        Exit Function
    End If

    colIdx = GetColIndex(tbl, colName)
    If colIdx = 0 Then
        LookupValueByColName = Empty
        Exit Function
    End If

    LookupValueByColName = LookupValue(d, colIdx)
End Function

Private Function GetColIndex(ByVal tbl As ListObject, ByVal colName As String) As Long
    ' Returns column index (1-based) or 0 if not found
    Dim col As ListColumn
    On Error Resume Next
    Set col = tbl.ListColumns(colName)
    If Not col Is Nothing Then GetColIndex = col.Index
    On Error GoTo 0
End Function

Private Function GetTelemTable() As ListObject
    ' Returns tblTelemetry or Nothing if not found
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(Schema.SHEET_TELEMETRY)
    If Not ws Is Nothing Then
        Set GetTelemTable = ws.ListObjects(Schema.TABLE_TELEMETRY)
    End If
    On Error GoTo 0
End Function
