Attribute VB_Name = "Data"
Option Explicit
' Data: Read/write worksheets.
' Purpose: Single module for all worksheet I/O. Load state, load config, save results.
' Dependencies: Types, Schema

' ==== Load State ==============================================================

' Load initial state from Input sheet
Public Function LoadState() As State
    Dim s As State
    Dim ws As Worksheet
    Dim i As Long
    Dim chemRange As Range

    On Error Resume Next

    Set ws = GetSheet(Schema.SHEET_INPUT)
    If ws Is Nothing Then Exit Function

    ' Volume
    s.Vol = Val(GetNamedValue(ws, Schema.NAME_INIT_VOL))

    ' Chemistry concentrations from latest results row
    Set chemRange = TryGetRange(ws, Schema.NAME_RES_ROW)
    If Not chemRange Is Nothing Then
        For i = 1 To Types.METRIC_COUNT
            If i <= chemRange.Columns.Count Then
                s.Chem(i) = Val(chemRange.Cells(1, i).Value)
            End If
        Next i
    End If

    ' Hidden mass (from previous run, if any)
    Set chemRange = TryGetRange(ws, Schema.NAME_HIDDEN_MASS)
    If Not chemRange Is Nothing Then
        For i = 1 To Types.METRIC_COUNT
            If i <= chemRange.Rows.Count Then
                s.Hidden(i) = Val(chemRange.Cells(i, 1).Value)
            End If
        Next i
    End If

    LoadState = s

    On Error GoTo 0
End Function

' ==== Load Config =============================================================

' Load config from Input sheet
Public Function LoadConfig() As Config
    Dim cfg As Config
    Dim ws As Worksheet
    Dim i As Long
    Dim chemRange As Range
    Dim enhMode As String

    On Error Resume Next

    Set ws = GetSheet(Schema.SHEET_INPUT)
    If ws Is Nothing Then Exit Function

    ' Mode: "On" = TwoBucket, else Simple
    enhMode = GetNamedValue(ws, Schema.NAME_ENHANCED_MODE)
    If UCase$(enhMode) = "ON" Then
        cfg.Mode = "TwoBucket"
    Else
        cfg.Mode = "Simple"
    End If

    ' Time
    cfg.Days = Schema.DEFAULT_FORECAST_DAYS
    cfg.StartDate = Val(GetNamedValue(ws, Schema.NAME_SAMPLE_DATE))

    ' Physics
    cfg.Tau = Val(GetNamedValue(ws, Schema.NAME_TAU))
    cfg.Outflow = Val(GetNamedValue(ws, Schema.NAME_NET_OUT))
    cfg.SurfaceFrac = Val(GetNamedValue(ws, Schema.NAME_SURFACE_FRACTION))
    If cfg.SurfaceFrac = 0 Then cfg.SurfaceFrac = Schema.DEFAULT_SURFACE_FRACTION

    ' Inflow from IR table (sum of active sources)
    LoadInflowFromIR ws, cfg

    ' Rain
    LoadRainConfig ws, cfg

    ' Triggers
    cfg.TriggerVol = Val(GetNamedValue(ws, Schema.NAME_TRIGGER_VOL))
    Set chemRange = TryGetRange(ws, Schema.NAME_LIMIT_ROW)
    If Not chemRange Is Nothing Then
        For i = 1 To Types.METRIC_COUNT
            If i <= chemRange.Columns.Count Then
                cfg.TriggerChem(i) = Val(chemRange.Cells(1, i).Value)
            End If
        Next i
    End If

    LoadConfig = cfg

    On Error GoTo 0
End Function

' Load inflow from IR table (sum of active sources)
Private Sub LoadInflowFromIR(ByVal ws As Worksheet, ByRef cfg As Config)
    Dim tbl As ListObject
    Dim row As ListRow
    Dim flowCol As Long, activeCol As Long, chemStartCol As Long
    Dim flow As Double
    Dim i As Long

    On Error Resume Next
    Set tbl = ws.ListObjects(Schema.TABLE_IR)
    On Error GoTo 0
    If tbl Is Nothing Then Exit Sub
    If tbl.ListRows.Count = 0 Then Exit Sub

    flowCol = ColIndex(tbl, Schema.IR_COL_FLOW)
    activeCol = ColIndex(tbl, Schema.IR_COL_ACTIVE)
    chemStartCol = ColIndex(tbl, Types.MetricName(1))

    If flowCol = 0 Then Exit Sub

    ' Sum active sources
    On Error Resume Next
    For Each row In tbl.ListRows
        If IsActive(row.Range.Cells(1, activeCol).Value) Then
            flow = Val(row.Range.Cells(1, flowCol).Value)
            cfg.Inflow = cfg.Inflow + flow

            ' Weight inflow chemistry by flow
            If chemStartCol > 0 Then
                For i = 1 To Types.METRIC_COUNT
                    cfg.InflowChem(i) = cfg.InflowChem(i) + _
                        flow * Val(row.Range.Cells(1, chemStartCol + i - 1).Value)
                Next i
            End If
        End If
    Next row
    On Error GoTo 0

    ' Convert weighted sum to average concentration
    If cfg.Inflow > Types.EPS Then
        For i = 1 To Types.METRIC_COUNT
            cfg.InflowChem(i) = cfg.InflowChem(i) / cfg.Inflow
        Next i
    End If
End Sub

' Load rain config
Private Sub LoadRainConfig(ByVal ws As Worksheet, ByRef cfg As Config)
    Dim rainFactor As Double
    Dim rainMode As String

    rainFactor = Val(GetNamedValue(ws, Schema.NAME_RAIN_FACTOR))
    rainMode = GetNamedValue(ws, Schema.NAME_RAIN_MODE)

    ' For now, use simple average or zero based on mode
    If UCase$(rainMode) = UCase$(Schema.RAIN_MODE_CONSERVATIVE) Then
        cfg.RainVol = 0
    Else
        ' Could calculate from historical rain, for now use factor as proxy
        cfg.RainVol = rainFactor * 0.5  ' Placeholder
    End If
End Sub

' ==== Save Results ============================================================

' Save result to Input sheet (trigger info)
Public Sub SaveResult(ByRef r As Result)
    Dim ws As Worksheet
    Dim triggerText As String

    On Error Resume Next

    Set ws = GetSheet(Schema.SHEET_INPUT)
    If ws Is Nothing Then Exit Sub

    ' Format trigger result
    If r.TriggerDay = Types.NO_TRIGGER Then
        triggerText = "No trigger in " & UBound(r.Snaps) & " days"
    Else
        triggerText = r.TriggerMetric & " on day " & r.TriggerDay & _
                      " (" & Format$(r.TriggerDate, "dd-mmm") & ")"
    End If

    ' Write to trigger output cell
    SetNamedValue ws, Schema.NAME_STD_TRIGGER, triggerText

    ' Save final hidden mass for next run
    SaveHiddenMass ws, r.FinalState

    On Error GoTo 0
End Sub

' Save hidden mass for next run
Private Sub SaveHiddenMass(ByVal ws As Worksheet, ByRef s As State)
    Dim chemRange As Range
    Dim i As Long

    Set chemRange = TryGetRange(ws, Schema.NAME_HIDDEN_MASS)
    If chemRange Is Nothing Then Exit Sub

    For i = 1 To Types.METRIC_COUNT
        If i <= chemRange.Rows.Count Then
            chemRange.Cells(i, 1).Value = s.Hidden(i)
        End If
    Next i
End Sub

' ==== Helper Functions ========================================================

Private Function GetSheet(ByVal name As String) As Worksheet
    On Error Resume Next
    Set GetSheet = ThisWorkbook.Worksheets(name)
    On Error GoTo 0
End Function

Private Function TryGetRange(ByVal ws As Worksheet, ByVal name As String) As Range
    On Error Resume Next
    Set TryGetRange = ws.Range(name)
    On Error GoTo 0
End Function

Private Function GetNamedValue(ByVal ws As Worksheet, ByVal name As String) As String
    Dim rng As Range
    Set rng = TryGetRange(ws, name)
    If Not rng Is Nothing Then
        GetNamedValue = CStr(rng.Value)
    End If
End Function

Private Sub SetNamedValue(ByVal ws As Worksheet, ByVal name As String, ByVal value As Variant)
    Dim rng As Range
    Set rng = TryGetRange(ws, name)
    If Not rng Is Nothing Then
        rng.Value = value
    End If
End Sub

Private Function ColIndex(ByVal tbl As ListObject, ByVal colName As String) As Long
    Dim col As ListColumn
    On Error Resume Next
    Set col = tbl.ListColumns(colName)
    If Not col Is Nothing Then ColIndex = col.Index
    On Error GoTo 0
End Function

Private Function IsActive(ByVal value As Variant) As Boolean
    Dim s As String
    s = UCase$(Trim$(CStr(value)))
    IsActive = (s = "TRUE" Or s = "YES" Or s = "ON" Or s = "1" Or s = "X")
End Function
