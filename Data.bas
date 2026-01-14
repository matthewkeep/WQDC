Option Explicit
' Data: Worksheet I/O.
' Dependencies: Core, Schema, Telemetry

' ==== Site Access ===========================================================

Public Function GetSite() As String
    ' Returns currently selected site from Inputs sheet
    Dim ws As Worksheet
    Set ws = Helpers.GetSheet(Schema.SHEET_INPUT)
    If Not ws Is Nothing Then
        On Error Resume Next
        GetSite = Trim$(CStr(ws.Range(Schema.NAME_SITE).Value))
        On Error GoTo 0
    End If
End Function

Public Function GetEnhancedMode() As String
    ' Returns Enhanced Mode setting (On/Off)
    Dim ws As Worksheet
    Set ws = Helpers.GetSheet(Schema.SHEET_INPUT)
    If Not ws Is Nothing Then
        On Error Resume Next
        GetEnhancedMode = Trim$(CStr(ws.Range(Schema.NAME_ENHANCED_MODE).Value))
        On Error GoTo 0
    End If
End Function

' ==== State Loading =========================================================

Public Function LoadState() As State
    Dim s As State, ws As Worksheet, rng As Range, i As Long
    On Error Resume Next
    Set ws = Helpers.GetSheet(Schema.SHEET_INPUT)
    If ws Is Nothing Then Exit Function

    s.Vol = Val(Helpers.ReadFromRange(ws, Schema.NAME_INIT_VOL))

    Set rng = Helpers.GetRng(ws, Schema.NAME_RES_ROW)
    If Not rng Is Nothing Then
        For i = 1 To Core.METRIC_COUNT
            If i <= rng.Columns.Count Then s.Chem(i) = Val(rng.Cells(1, i).Value)
        Next i
    End If

    Set rng = Helpers.GetRng(ws, Schema.NAME_HIDDEN_MASS)
    If Not rng Is Nothing Then
        For i = 1 To Core.METRIC_COUNT
            If i <= rng.Rows.Count Then s.Hidden(i) = Val(rng.Cells(i, 1).Value)
        Next i
    End If

    LoadState = s
    On Error GoTo 0
End Function

Public Function SnapState(ByRef s As State, ByVal site As String) As State
    ' Calibrates state by snapping visible layer to latest telemetry
    ' Hidden layer unchanged - trust model's physics-based estimate
    ' This follows data assimilation best practice: direct insertion for
    ' observable states, model continuity for unobservable states
    Dim snapped As State, latestVol As Variant, latestEC As Variant
    snapped = Core.CopyState(s)

    latestVol = Telemetry.GetLatestVol(Date, site)
    latestEC = Telemetry.GetLatestEC(Date, site)

    ' Snap visible layer only
    If Not IsEmpty(latestVol) Then snapped.Vol = CDbl(latestVol)
    If Not IsEmpty(latestEC) Then snapped.Chem(1) = CDbl(latestEC)

    SnapState = snapped
End Function

Public Function GetTelemCalEnabled() As Boolean
    ' Returns True if telemetry calibration is enabled
    Dim ws As Worksheet
    Set ws = Helpers.GetSheet(Schema.SHEET_INPUT)
    If Not ws Is Nothing Then
        On Error Resume Next
        GetTelemCalEnabled = (UCase$(Trim$(ws.Range(Schema.NAME_TELEM_CAL).Value)) = "ON")
        On Error GoTo 0
    End If
End Function

Public Function GetPredMode() As String
    ' Returns "STANDARD" or "ENHANCED" based on J5 toggle cell
    Dim ws As Worksheet
    Set ws = Helpers.GetSheet(Schema.SHEET_INPUT)
    If Not ws Is Nothing Then
        On Error Resume Next
        GetPredMode = UCase$(Trim$(ws.Range(Schema.NAME_PRED_MODE).Value))
        On Error GoTo 0
    End If
    If GetPredMode <> "ENHANCED" Then GetPredMode = "STANDARD"  ' Default to Standard
End Function

Public Function LoadConfig(ByVal site As String, ByVal runType As String) As Config
    ' Loads config for Standard or Enhanced run
    ' Standard: Simple mode, no rainfall adjustment, no telemetry calibration
    ' Enhanced: Uses configured Mixing Model, Rainfall Mode, Telemetry Cal
    Dim cfg As Config, ws As Worksheet, rng As Range, i As Long
    Dim mixingModel As String, rainfallMode As String, telemCal As String
    On Error Resume Next
    Set ws = Helpers.GetSheet(Schema.SHEET_INPUT)
    If ws Is Nothing Then Exit Function

    ' Common config
    cfg.Site = site
    cfg.Days = Schema.DEFAULT_FORECAST_DAYS
    cfg.StartDate = Helpers.GetDateVal(ws, Schema.NAME_SAMPLE_DATE)
    cfg.Tau = Val(Helpers.ReadFromRange(ws, Schema.NAME_TAU))
    If cfg.Tau < Core.EPS Then cfg.Tau = 7  ' Default 7 days for TwoBucket mixing
    cfg.Outflow = Val(Helpers.ReadFromRange(ws, Schema.NAME_OUTPUT))
    cfg.SurfaceFrac = Val(Helpers.ReadFromRange(ws, Schema.NAME_SURFACE_FRACTION))
    If cfg.SurfaceFrac < Core.EPS Then cfg.SurfaceFrac = Schema.DEFAULT_SURFACE_FRACTION

    ' Load inflows
    LoadInflowIR ws, cfg

    ' Load triggers
    cfg.TriggerVol = Val(Helpers.ReadFromRange(ws, Schema.NAME_TRIGGER_VOL))
    Set rng = Helpers.GetRng(ws, Schema.NAME_LIMIT_ROW)
    If Not rng Is Nothing Then
        For i = 1 To Core.METRIC_COUNT
            If i <= rng.Columns.Count Then cfg.TriggerChem(i) = Val(rng.Cells(1, i).Value)
        Next i
    End If

    ' Mode-specific settings
    If UCase$(runType) = "ENHANCED" Then
        ' Enhanced: read configured options
        mixingModel = CStr(Helpers.ReadFromRange(ws, Schema.NAME_MIXING_MODEL))
        rainfallMode = CStr(Helpers.ReadFromRange(ws, Schema.NAME_RAINFALL_MODE))
        telemCal = CStr(Helpers.ReadFromRange(ws, Schema.NAME_TELEM_CAL))

        ' Set mixing model
        If UCase$(mixingModel) = UCase$(Schema.MIXING_TWOBUCKET) Then
            cfg.Mode = "TwoBucket"
        Else
            cfg.Mode = "Simple"
        End If

        ' Set rainfall mode (applied per-day in Sim.Run)
        cfg.RainfallMode = rainfallMode

        ' Load rain factor (mm to ML conversion)
        cfg.RainFactor = Val(Helpers.ReadFromRange(ws, Schema.NAME_RAIN_FACTOR))
        If cfg.RainFactor = 0 Then cfg.RainFactor = 1  ' Default 1:1
    Else
        ' Standard: Simple mode, no rainfall, no calibration
        cfg.Mode = "Simple"
        cfg.RainfallMode = Schema.RAINFALL_OFF
    End If

    LoadConfig = cfg
    On Error GoTo 0
End Function

Private Sub LoadInflowIR(ByVal ws As Worksheet, ByRef cfg As Config)
    Dim tbl As ListObject, row As ListRow
    Dim flowCol As Long, activeCol As Long, chemCol As Long
    Dim flow As Double, i As Long
    Dim chemNames As Variant
    On Error Resume Next
    Set tbl = ws.ListObjects(Schema.TABLE_IR)
    On Error GoTo 0
    If tbl Is Nothing Then Exit Sub
    If tbl.ListRows.Count = 0 Then Exit Sub

    chemNames = Schema.ChemistryNames()
    flowCol = Helpers.ColIdx(tbl, Schema.IR_COL_FLOW)
    activeCol = Helpers.ColIdx(tbl, Schema.IR_COL_ACTIVE)
    chemCol = Helpers.ColIdx(tbl, chemNames(0))  ' First chemistry column (e.g., "EC (uS/cm)")
    If flowCol = 0 Then Exit Sub

    On Error Resume Next
    For Each row In tbl.ListRows
        If IsActive(row.Range.Cells(1, activeCol).Value) Then
            flow = Val(row.Range.Cells(1, flowCol).Value)
            cfg.Inflow = cfg.Inflow + flow
            If chemCol > 0 Then
                For i = 1 To Core.METRIC_COUNT
                    cfg.InflowChem(i) = cfg.InflowChem(i) + flow * Val(row.Range.Cells(1, chemCol + i - 1).Value)
                Next i
            End If
        End If
    Next row
    On Error GoTo 0

    If cfg.Inflow > Core.EPS Then
        For i = 1 To Core.METRIC_COUNT
            cfg.InflowChem(i) = cfg.InflowChem(i) / cfg.Inflow
        Next i
    End If
End Sub

Public Sub SaveResult(ByRef r As Result, ByVal runType As String)
    ' Saves result to appropriate output cell based on runType
    ' Days output is relative to Run Date (today), not Sample Date
    Dim ws As Worksheet, rng As Range, i As Long
    Dim predState As State, targetName As String
    Dim sampleDate As Date, runDate As Date, dayOffset As Long
    On Error Resume Next
    Set ws = Helpers.GetSheet(Schema.SHEET_INPUT)
    If ws Is Nothing Then Exit Sub

    ' Use state at trigger day for predictions (or final if no trigger)
    If r.TriggerDay <> Core.NO_TRIGGER Then
        predState = r.Snaps(r.TriggerDay)
    Else
        predState = r.FinalState
    End If

    ' Calculate days from Run Date (today), not Sample Date
    sampleDate = Helpers.GetDateVal(ws, Schema.NAME_SAMPLE_DATE)
    runDate = Helpers.GetDateVal(ws, Schema.NAME_RUN_DATE)
    dayOffset = CLng(runDate - sampleDate)

    Dim days As Long
    If r.TriggerDay = Core.NO_TRIGGER Then
        days = UBound(r.Snaps) - dayOffset
    Else
        days = r.TriggerDay - dayOffset
    End If

    ' Write to appropriate trigger output cell
    If UCase$(runType) = "ENHANCED" Then
        targetName = Schema.NAME_ENH_TRIGGER
    Else
        targetName = Schema.NAME_STD_TRIGGER
    End If
    Helpers.WriteToRange ws, targetName, days

    ' Write to Pred_Row based on Pred_Mode toggle (J5: Standard/Enhanced)
    Dim predMode As String
    predMode = GetPredMode()

    If (UCase$(runType) = "STANDARD" And predMode = "STANDARD") Or _
       (UCase$(runType) = "ENHANCED" And predMode = "ENHANCED") Then
        Helpers.WriteToRange ws, Schema.NAME_RESULT_VOL, predState.Vol
        Set rng = Helpers.GetRng(ws, Schema.NAME_PRED_ROW)
        If Not rng Is Nothing Then
            For i = 1 To Core.METRIC_COUNT
                If i <= rng.Columns.Count Then rng.Cells(1, i).Value = predState.Chem(i)
            Next i
        End If
        ' Format triggered cell if trigger occurred
        If r.TriggerDay <> Core.NO_TRIGGER Then
            FormatTriggeredCell ws, r.TriggerMetric
        End If
    End If

    ' Hidden mass NOT written to Inputs - would corrupt sample-date state with end-of-run values.
    ' Continuity handled via tblLive (SimLog stores Hidden per-date).
    ' Inputs updates when Sample Date changes (Events.OnInputsChange → LoadHiddenForDate).
    On Error GoTo 0
End Sub

' ==== Private Helpers ========================================================

Private Function IsActive(ByVal v As Variant) As Boolean
    Dim s As String
    s = UCase$(Trim$(CStr(v)))
    IsActive = (s = "TRUE" Or s = "YES" Or s = "ON" Or s = "1" Or s = "X")
End Function

Private Sub FormatTriggeredCell(ByVal ws As Worksheet, ByVal metricName As String)
    ' Applies red bold formatting to the triggered metric in Pred_Row
    Dim rng As Range, colIdx As Long, i As Long

    ' Clear previous trigger formatting
    Set rng = Helpers.GetRng(ws, Schema.NAME_PRED_ROW)
    If Not rng Is Nothing Then
        rng.Font.Bold = False
        rng.Font.Color = vbBlack
    End If
    Set rng = Helpers.GetRng(ws, Schema.NAME_RESULT_VOL)
    If Not rng Is Nothing Then
        rng.Font.Bold = False
        rng.Font.Color = vbBlack
    End If

    ' Find and format triggered metric
    If metricName = "Volume" Then
        Set rng = Helpers.GetRng(ws, Schema.NAME_RESULT_VOL)
    Else
        For i = 1 To Core.METRIC_COUNT
            If Core.MetricName(i) = metricName Then colIdx = i: Exit For
        Next i
        If colIdx > 0 Then
            Set rng = Helpers.GetRng(ws, Schema.NAME_PRED_ROW)
            If Not rng Is Nothing Then Set rng = rng.Cells(1, colIdx)
        End If
    End If

    If Not rng Is Nothing Then
        rng.Font.Bold = True
        rng.Font.Color = Schema.COLOR_TRIGGER_FONT
    End If
End Sub

' ==== Log-Based State Loading ===============================================

Public Function LoadHiddenFromLog(ByVal site As String, ByVal targetDate As Date) As State
    ' Loads hidden layer state from tblLive at targetDate for TwoBucket continuity
    ' Returns State with Hidden(1-7) populated, other fields zeroed
    ' If no data found, returns empty state (Hidden = 0)
    Dim tbl As ListObject, ws As Worksheet
    Dim rowIdx As Long, i As Long, hidCol As Long
    Dim s As State

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(Schema.SHEET_LOG)
    On Error GoTo 0
    If ws Is Nothing Then Exit Function

    ' Get live table for site
    On Error Resume Next
    Set tbl = ws.ListObjects(Helpers.LiveTableName(site))
    On Error GoTo 0
    If tbl Is Nothing Then Exit Function
    If tbl.DataBodyRange Is Nothing Then Exit Function

    ' Find row for target date
    rowIdx = FindLogRowByDate(tbl, targetDate)
    If rowIdx = 0 Then Exit Function

    ' Read hidden layer values
    For i = 1 To Core.METRIC_COUNT
        hidCol = Helpers.ColIdx(tbl, Schema.EnhHidColName(i))
        If hidCol > 0 Then
            s.Hidden(i) = Val(tbl.DataBodyRange.Cells(rowIdx, hidCol).Value)
        End If
    Next i

    LoadHiddenFromLog = s
End Function

Public Sub LoadHiddenForDate(ByVal site As String, ByVal targetDate As Date)
    ' Loads hidden mass from log at date and displays in Inputs sheet
    ' Called when user changes Sample Date for TwoBucket feedback
    Dim s As State, ws As Worksheet, rng As Range, i As Long

    If Len(site) = 0 Or targetDate = 0 Then Exit Sub

    s = LoadHiddenFromLog(site, targetDate)
    If Core.IsHiddenEmpty(s) Then Exit Sub  ' No data found

    Set ws = Helpers.GetSheet(Schema.SHEET_INPUT)
    If ws Is Nothing Then Exit Sub

    Set rng = Helpers.GetRng(ws, Schema.NAME_HIDDEN_MASS)
    If rng Is Nothing Then Exit Sub

    For i = 1 To Core.METRIC_COUNT
        If i <= rng.Rows.Count Then rng.Cells(i, 1).Value = s.Hidden(i)
    Next i
End Sub

Public Function HasLogDataForDate(ByVal site As String, ByVal targetDate As Date) As Boolean
    ' Returns True if tblLive has data for the specified date
    Dim tbl As ListObject, ws As Worksheet

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(Schema.SHEET_LOG)
    On Error GoTo 0
    If ws Is Nothing Then Exit Function

    On Error Resume Next
    Set tbl = ws.ListObjects(Helpers.LiveTableName(site))
    On Error GoTo 0
    If tbl Is Nothing Then Exit Function
    If tbl.DataBodyRange Is Nothing Then Exit Function

    HasLogDataForDate = (FindLogRowByDate(tbl, targetDate) > 0)
End Function

Private Function FindLogRowByDate(ByVal tbl As ListObject, ByVal targetDate As Date) As Long
    ' Returns row index (1-based) for date in log table, or 0 if not found
    ' Delegates to shared O(1) utility
    FindLogRowByDate = Helpers.FindRowByDate(tbl, targetDate)
End Function
