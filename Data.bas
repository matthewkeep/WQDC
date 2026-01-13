Option Explicit
' Data: Worksheet I/O.
' Dependencies: Types, Schema

Public Function LoadState() As State
    Dim s As State, ws As Worksheet, rng As Range, i As Long
    On Error Resume Next
    Set ws = GetSheet(Schema.SHEET_INPUT)
    If ws Is Nothing Then Exit Function

    s.Vol = Val(GetVal(ws, Schema.NAME_INIT_VOL))

    Set rng = GetRng(ws, Schema.NAME_RES_ROW)
    If Not rng Is Nothing Then
        For i = 1 To AAATypes.METRIC_COUNT
            If i <= rng.Columns.Count Then s.Chem(i) = Val(rng.Cells(1, i).Value)
        Next i
    End If

    Set rng = GetRng(ws, Schema.NAME_HIDDEN_MASS)
    If Not rng Is Nothing Then
        For i = 1 To AAATypes.METRIC_COUNT
            If i <= rng.Rows.Count Then s.Hidden(i) = Val(rng.Cells(i, 1).Value)
        Next i
    End If

    LoadState = s
    On Error GoTo 0
End Function

Public Function LoadConfig() As Config
    Dim cfg As Config, ws As Worksheet, rng As Range, i As Long, mode As String
    On Error Resume Next
    Set ws = GetSheet(Schema.SHEET_INPUT)
    If ws Is Nothing Then Exit Function

    mode = GetVal(ws, Schema.NAME_ENHANCED_MODE)
    cfg.Mode = IIf(UCase$(mode) = "ON", "TwoBucket", "Simple")
    cfg.Days = Schema.DEFAULT_FORECAST_DAYS
    cfg.StartDate = Val(GetVal(ws, Schema.NAME_SAMPLE_DATE))
    cfg.Tau = Val(GetVal(ws, Schema.NAME_TAU))
    cfg.Outflow = Val(GetVal(ws, Schema.NAME_NET_OUT))
    cfg.SurfaceFrac = Val(GetVal(ws, Schema.NAME_SURFACE_FRACTION))
    If cfg.SurfaceFrac = 0 Then cfg.SurfaceFrac = Schema.DEFAULT_SURFACE_FRACTION

    LoadInflowIR ws, cfg

    cfg.TriggerVol = Val(GetVal(ws, Schema.NAME_TRIGGER_VOL))
    Set rng = GetRng(ws, Schema.NAME_LIMIT_ROW)
    If Not rng Is Nothing Then
        For i = 1 To AAATypes.METRIC_COUNT
            If i <= rng.Columns.Count Then cfg.TriggerChem(i) = Val(rng.Cells(1, i).Value)
        Next i
    End If

    LoadConfig = cfg
    On Error GoTo 0
End Function

Private Sub LoadInflowIR(ByVal ws As Worksheet, ByRef cfg As Config)
    Dim tbl As ListObject, row As ListRow
    Dim flowCol As Long, activeCol As Long, chemCol As Long
    Dim flow As Double, i As Long
    On Error Resume Next
    Set tbl = ws.ListObjects(Schema.TABLE_IR)
    On Error GoTo 0
    If tbl Is Nothing Then Exit Sub
    If tbl.ListRows.Count = 0 Then Exit Sub

    flowCol = ColIdx(tbl, Schema.IR_COL_FLOW)
    activeCol = ColIdx(tbl, Schema.IR_COL_ACTIVE)
    chemCol = ColIdx(tbl, AAATypes.MetricName(1))
    If flowCol = 0 Then Exit Sub

    On Error Resume Next
    For Each row In tbl.ListRows
        If IsActive(row.Range.Cells(1, activeCol).Value) Then
            flow = Val(row.Range.Cells(1, flowCol).Value)
            cfg.Inflow = cfg.Inflow + flow
            If chemCol > 0 Then
                For i = 1 To AAATypes.METRIC_COUNT
                    cfg.InflowChem(i) = cfg.InflowChem(i) + flow * Val(row.Range.Cells(1, chemCol + i - 1).Value)
                Next i
            End If
        End If
    Next row
    On Error GoTo 0

    If cfg.Inflow > AAATypes.EPS Then
        For i = 1 To AAATypes.METRIC_COUNT
            cfg.InflowChem(i) = cfg.InflowChem(i) / cfg.Inflow
        Next i
    End If
End Sub

Public Sub SaveResult(ByRef r As Result)
    Dim ws As Worksheet, txt As String, rng As Range, i As Long
    On Error Resume Next
    Set ws = GetSheet(Schema.SHEET_INPUT)
    If ws Is Nothing Then Exit Sub

    If r.TriggerDay = AAATypes.NO_TRIGGER Then
        txt = "No trigger in " & UBound(r.Snaps) & " days"
    Else
        txt = r.TriggerMetric & " day " & r.TriggerDay & " (" & Format$(r.TriggerDate, "dd-mmm") & ")"
    End If
    SetVal ws, Schema.NAME_STD_TRIGGER, txt

    Set rng = GetRng(ws, Schema.NAME_HIDDEN_MASS)
    If Not rng Is Nothing Then
        For i = 1 To AAATypes.METRIC_COUNT
            If i <= rng.Rows.Count Then rng.Cells(i, 1).Value = r.FinalState.Hidden(i)
        Next i
    End If
    On Error GoTo 0
End Sub

' ==== Helpers ================================================================

Private Function GetSheet(ByVal nm As String) As Worksheet
    On Error Resume Next
    Set GetSheet = ThisWorkbook.Worksheets(nm)
    On Error GoTo 0
End Function

Private Function GetRng(ByVal ws As Worksheet, ByVal nm As String) As Range
    On Error Resume Next
    Set GetRng = ws.Range(nm)
    On Error GoTo 0
End Function

Private Function GetVal(ByVal ws As Worksheet, ByVal nm As String) As String
    Dim rng As Range
    Set rng = GetRng(ws, nm)
    If Not rng Is Nothing Then GetVal = CStr(rng.Value)
End Function

Private Sub SetVal(ByVal ws As Worksheet, ByVal nm As String, ByVal v As Variant)
    Dim rng As Range
    Set rng = GetRng(ws, nm)
    If Not rng Is Nothing Then rng.Value = v
End Sub

Private Function ColIdx(ByVal tbl As ListObject, ByVal nm As String) As Long
    Dim col As ListColumn
    On Error Resume Next
    Set col = tbl.ListColumns(nm)
    If Not col Is Nothing Then ColIdx = col.Index
    On Error GoTo 0
End Function

Private Function IsActive(ByVal v As Variant) As Boolean
    Dim s As String
    s = UCase$(Trim$(CStr(v)))
    IsActive = (s = "TRUE" Or s = "YES" Or s = "ON" Or s = "1" Or s = "X")
End Function
