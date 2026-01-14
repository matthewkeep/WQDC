Option Explicit
' Backtest: Season replay for prediction validation.
' Dependencies: Core, Schema, Data, Sim, Telemetry, Setup
'
' Runs both Standard and Enhanced modes for A/B comparison:
' - Standard: Simple mixing, independent runs (no state carryover)
' - Enhanced: Uses configured settings, hidden layer carries forward

' ==== Entry Point ==============================================================

Public Sub RunSeason()
    ' Backtests all RR samples for current site using both Standard and Enhanced
    ' Simulates weekly operational workflow with progressive hidden layer
    Dim site As String, samples As Variant
    Dim i As Long, n As Long, predictDay As Long, cm As XlCalculation
    Dim sStd As State, sEnh As State, cfgStd As Config, cfgEnh As Config
    Dim rStd As Result, rEnh As Result
    Dim results() As Variant
    Dim enhancedMode As Boolean, telemCalEnabled As Boolean

    site = Data.GetSite()
    If Len(site) = 0 Then
        MsgBox "No site selected.", vbExclamation, "Backtest"
        Exit Sub
    End If

    samples = GetAllSamples(site)
    If Not IsArray(samples) Then
        MsgBox "No samples found for " & site & " in Results table.", vbExclamation, "Backtest"
        Exit Sub
    End If

    n = UBound(samples, 1)
    If n < 2 Then
        MsgBox "Need at least 2 samples to backtest (found " & n & ").", vbExclamation, "Backtest"
        Exit Sub
    End If

    ' Check current Enhanced settings
    enhancedMode = (UCase$(Data.GetEnhancedMode()) = "ON")
    telemCalEnabled = Data.GetTelemCalEnabled()

    On Error GoTo Cleanup
    cm = Application.Calculation
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual

    ' Ensure season log table exists and clear it
    Setup.EnsureSeasonLogTable site
    ClearSeasonLog site

    ' Results: RunDate, SampleDate, ActualEC, ActualVol, StdPredEC, StdErrEC, StdPredVol, StdErrVol, EnhPredEC, EnhErrEC, EnhPredVol, EnhErrVol
    ReDim results(1 To n - 1, 1 To 12)

    ' Initialize Enhanced hidden state at equilibrium from first sample
    sEnh = LoadStateAtDate(site, samples(1, 1))
    sEnh = Core.InitHiddenAtEquilibrium(sEnh)

    For i = 1 To n - 1
        ' === Standard Run (independent each time) ===
        sStd = LoadStateAtDate(site, samples(i, 1))
        cfgStd = LoadConfigForBacktest(site, samples(i, 1), "Standard")
        cfgStd.StartDate = samples(i, 1) + 7
        rStd = Sim.Run(sStd, cfgStd)

        ' === Enhanced Run (carries hidden state forward) ===
        If enhancedMode Then
            ' Load visible state from observed data
            sEnh = LoadStateAtDate(site, samples(i, 1))

            ' Apply telemetry calibration if enabled (snap visible layer)
            If telemCalEnabled Then
                sEnh = SnapVisibleLayer(sEnh, site, samples(i, 1))
            End If

            ' Preserve hidden layer from previous run (or equilibrium if first)
            If i > 1 Then
                ' Use hidden state from end of previous Enhanced run
                sEnh = CarryHiddenFromPrevious(sEnh, rEnh.FinalState)
            End If

            cfgEnh = LoadConfigForBacktest(site, samples(i, 1), "Enhanced")
            cfgEnh.StartDate = samples(i, 1) + 7
            rEnh = Sim.Run(sEnh, cfgEnh)
        End If

        ' Calculate prediction day (when next sample occurs)
        predictDay = samples(i + 1, 1) - (samples(i, 1) + 7)
        If predictDay < 0 Then predictDay = 0
        If predictDay > UBound(rStd.Snaps) Then predictDay = UBound(rStd.Snaps)

        ' Record results
        results(i, 1) = Date                                      ' RunDate
        results(i, 2) = samples(i, 1)                             ' SampleDate
        results(i, 3) = samples(i + 1, 2)                         ' ActualEC (next sample)
        results(i, 4) = samples(i + 1, 3)                         ' ActualVol (next sample)

        ' Standard predictions
        results(i, 5) = rStd.Snaps(predictDay).Chem(1)            ' StdPredEC
        results(i, 6) = results(i, 5) - results(i, 3)             ' StdErrEC
        results(i, 7) = rStd.Snaps(predictDay).Vol                ' StdPredVol
        results(i, 8) = results(i, 7) - results(i, 4)             ' StdErrVol

        ' Enhanced predictions (if enabled)
        If enhancedMode Then
            results(i, 9) = rEnh.Snaps(predictDay).Chem(1)        ' EnhPredEC
            results(i, 10) = results(i, 9) - results(i, 3)        ' EnhErrEC
            results(i, 11) = rEnh.Snaps(predictDay).Vol           ' EnhPredVol
            results(i, 12) = results(i, 11) - results(i, 4)       ' EnhErrVol
        Else
            results(i, 9) = Empty: results(i, 10) = Empty
            results(i, 11) = Empty: results(i, 12) = Empty
        End If
    Next i

    WriteSeasonLog site, results

    Application.Calculation = cm
    Application.ScreenUpdating = True
    Application.EnableEvents = True

    Dim msg As String
    msg = "Backtest complete: " & (n - 1) & " samples processed." & vbNewLine & vbNewLine
    msg = msg & "Standard: Simple mode, independent runs" & vbNewLine
    If enhancedMode Then
        msg = msg & "Enhanced: " & GetInputVal(Schema.NAME_MIXING_MODEL) & " mode, progressive hidden layer"
        If telemCalEnabled Then msg = msg & ", telemetry calibration"
    Else
        msg = msg & "Enhanced: Off (enable to compare)"
    End If
    msg = msg & vbNewLine & vbNewLine & "Results in SeasonLog table on Log sheet."
    MsgBox msg, vbInformation, "Backtest"
    Exit Sub

Cleanup:
    Application.Calculation = cm
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    If Err.Number <> 0 Then
        MsgBox "Error: " & Err.Description, vbExclamation, "Backtest"
    End If
End Sub

' ==== Hidden Layer Management ==================================================

Private Function CarryHiddenFromPrevious(ByRef current As State, ByRef previous As State) As State
    ' Preserves hidden layer from previous run's final state
    Dim result As State, i As Long
    result = Core.CopyState(current)
    For i = 1 To Core.METRIC_COUNT
        result.Hidden(i) = previous.Hidden(i)
    Next i
    result.HidVol = previous.HidVol
    CarryHiddenFromPrevious = result
End Function

Private Function SnapVisibleLayer(ByRef s As State, ByVal site As String, ByVal sampleDate As Date) As State
    ' Snaps visible layer to telemetry values (hidden unchanged)
    Dim snapped As State, latestVol As Variant, latestEC As Variant
    snapped = Core.CopyState(s)

    latestVol = Telemetry.GetLatestVol(sampleDate, site)
    latestEC = Telemetry.GetLatestEC(sampleDate, site)

    If Not IsEmpty(latestVol) Then snapped.Vol = CDbl(latestVol)
    If Not IsEmpty(latestEC) Then snapped.Chem(1) = CDbl(latestEC)

    SnapVisibleLayer = snapped
End Function

' ==== Sample Data Access =======================================================

Private Function GetAllSamples(ByVal site As String) As Variant
    ' Returns 2D array of (SampleDate, EC, Vol) sorted by date ascending
    Dim tbl As ListObject, row As ListRow
    Dim sampleDate As Date, ec As Double, vol As Variant
    Dim dict As Object, i As Long, cnt As Long
    Dim dates() As Date, ecs() As Double, vols() As Double
    Dim result() As Variant

    Set tbl = GetResultsTable()
    If tbl Is Nothing Then Exit Function
    If tbl.ListRows.Count = 0 Then Exit Function

    ' Collect all samples for this site
    Set dict = New DictionaryShim
    For Each row In tbl.ListRows
        If Schema.MatchesSite(row.Range.Cells(1, 1).Value, site) Then
            On Error Resume Next
            sampleDate = CDate(row.Range.Cells(1, 2).Value)
            ec = Val(row.Range.Cells(1, Schema.ColIdx(tbl, Schema.ChemistryNames()(0))).Value)
            On Error GoTo 0

            If sampleDate > 0 And Not dict.Exists(CLng(sampleDate)) Then
                dict.Add CLng(sampleDate), Array(sampleDate, ec)
            End If
        End If
    Next row

    If dict.Count < 2 Then Exit Function

    cnt = dict.Count
    ReDim dates(1 To cnt)
    ReDim ecs(1 To cnt)
    ReDim vols(1 To cnt)

    i = 1
    Dim k As Variant, arr As Variant
    For Each k In dict.Keys
        arr = dict(k)
        dates(i) = arr(0)
        ecs(i) = arr(1)
        vol = Telemetry.GetLatestVol(arr(0), site)
        If IsEmpty(vol) Then vols(i) = 0 Else vols(i) = CDbl(vol)
        i = i + 1
    Next k

    SortByDate dates, ecs, vols

    ReDim result(1 To cnt, 1 To 3)
    For i = 1 To cnt
        result(i, 1) = dates(i)
        result(i, 2) = ecs(i)
        result(i, 3) = vols(i)
    Next i

    GetAllSamples = result
End Function

Private Sub SortByDate(ByRef dates() As Date, ByRef ecs() As Double, ByRef vols() As Double)
    Dim i As Long, j As Long, n As Long
    Dim tmpDate As Date, tmpEc As Double, tmpVol As Double

    n = UBound(dates)
    For i = 1 To n - 1
        For j = i + 1 To n
            If dates(j) < dates(i) Then
                tmpDate = dates(i): dates(i) = dates(j): dates(j) = tmpDate
                tmpEc = ecs(i): ecs(i) = ecs(j): ecs(j) = tmpEc
                tmpVol = vols(i): vols(i) = vols(j): vols(j) = tmpVol
            End If
        Next j
    Next i
End Sub

Private Function LoadStateAtDate(ByVal site As String, ByVal sampleDate As Date) As State
    ' Loads state from tblResults at specific date
    Dim s As State, tbl As ListObject, row As ListRow
    Dim rowDate As Date, vol As Variant
    Dim chem As Variant, i As Long

    Set tbl = GetResultsTable()
    If tbl Is Nothing Then Exit Function

    chem = Schema.ChemistryNames()

    For Each row In tbl.ListRows
        If Schema.MatchesSite(row.Range.Cells(1, 1).Value, site) Then
            On Error Resume Next
            rowDate = CDate(row.Range.Cells(1, 2).Value)
            On Error GoTo 0

            If rowDate = sampleDate Then
                For i = 0 To Core.METRIC_COUNT - 1
                    s.Chem(i + 1) = Val(row.Range.Cells(1, Schema.ColIdx(tbl, chem(i))).Value)
                Next i

                vol = Telemetry.GetLatestVol(sampleDate, site)
                If Not IsEmpty(vol) Then s.Vol = CDbl(vol)

                Exit For
            End If
        End If
    Next row

    LoadStateAtDate = s
End Function

Private Function LoadConfigForBacktest(ByVal site As String, ByVal beforeDate As Date, ByVal runType As String) As Config
    ' Loads config with IR chemistry from before the given date
    Dim cfg As Config, tblCat As ListObject, tblRes As ListObject
    Dim catRow As ListRow, irSite As String, flow As Double
    Dim labData As Variant, chem As Variant, i As Long
    Dim mixingModel As String, rainfallMode As String

    cfg.Site = site
    cfg.Days = Schema.DEFAULT_FORECAST_DAYS
    cfg.Tau = Val(GetInputVal(Schema.NAME_TAU))
    cfg.Outflow = Val(GetInputVal(Schema.NAME_OUTPUT))
    cfg.SurfaceFrac = Val(GetInputVal(Schema.NAME_SURFACE_FRACTION))
    If cfg.SurfaceFrac = 0 Then cfg.SurfaceFrac = Schema.DEFAULT_SURFACE_FRACTION

    ' Load triggers
    cfg.TriggerVol = Val(GetInputVal(Schema.NAME_TRIGGER_VOL))
    Dim rng As Range
    On Error Resume Next
    Set rng = ThisWorkbook.Worksheets(Schema.SHEET_INPUT).Range(Schema.NAME_LIMIT_ROW)
    If Not rng Is Nothing Then
        For i = 1 To Core.METRIC_COUNT
            If i <= rng.Columns.Count Then cfg.TriggerChem(i) = Val(rng.Cells(1, i).Value)
        Next i
    End If
    On Error GoTo 0

    ' Mode-specific settings
    If UCase$(runType) = "ENHANCED" Then
        mixingModel = GetInputVal(Schema.NAME_MIXING_MODEL)
        rainfallMode = GetInputVal(Schema.NAME_RAINFALL_MODE)

        If UCase$(mixingModel) = UCase$(Schema.MIXING_TWOBUCKET) Then
            cfg.Mode = "TwoBucket"
        Else
            cfg.Mode = "Simple"
        End If
        cfg.RainfallMode = rainfallMode
    Else
        cfg.Mode = "Simple"
        cfg.RainfallMode = Schema.RAINFALL_OFF
    End If

    ' Load IR flows and chemistry
    Set tblCat = GetCatalogTable()
    Set tblRes = GetResultsTable()
    If tblCat Is Nothing Or tblRes Is Nothing Then
        LoadConfigForBacktest = cfg
        Exit Function
    End If

    chem = Schema.ChemistryNames()

    For Each catRow In tblCat.ListRows
        If Schema.MatchesSite(catRow.Range.Cells(1, 1).Value, site) Then
            irSite = Trim$(catRow.Range.Cells(1, 2).Value)
            flow = Val(catRow.Range.Cells(1, 3).Value)

            labData = GetLabDataBeforeDate(irSite, beforeDate, tblRes, chem)

            If Not IsEmpty(labData) Then
                cfg.Inflow = cfg.Inflow + flow
                For i = 1 To Core.METRIC_COUNT
                    cfg.InflowChem(i) = cfg.InflowChem(i) + flow * labData(i)
                Next i
            End If
        End If
    Next catRow

    If cfg.Inflow > Core.EPS Then
        For i = 1 To Core.METRIC_COUNT
            cfg.InflowChem(i) = cfg.InflowChem(i) / cfg.Inflow
        Next i
    End If

    LoadConfigForBacktest = cfg
End Function

Private Function GetLabDataBeforeDate(ByVal irSite As String, ByVal beforeDate As Date, _
                                      ByVal tbl As ListObject, ByVal chem As Variant) As Variant
    Dim row As ListRow, rowDate As Date
    Dim latestDate As Date, latestRow As ListRow
    Dim result() As Double, i As Long

    latestDate = 0
    For Each row In tbl.ListRows
        If Schema.MatchesSite(row.Range.Cells(1, 1).Value, irSite) Then
            On Error Resume Next
            rowDate = CDate(row.Range.Cells(1, 2).Value)
            On Error GoTo 0

            If rowDate > 0 And rowDate < beforeDate And rowDate > latestDate Then
                latestDate = rowDate
                Set latestRow = row
            End If
        End If
    Next row

    If latestRow Is Nothing Then Exit Function

    ReDim result(1 To Core.METRIC_COUNT)
    For i = 0 To Core.METRIC_COUNT - 1
        result(i + 1) = Val(latestRow.Range.Cells(1, Schema.ColIdx(tbl, chem(i))).Value)
    Next i

    GetLabDataBeforeDate = result
End Function

' ==== Season Log Output ========================================================

Private Sub WriteSeasonLog(ByVal site As String, ByRef results() As Variant)
    Dim tbl As ListObject, i As Long, n As Long, newRow As ListRow

    Set tbl = GetSeasonLogTable(site)
    If tbl Is Nothing Then Exit Sub

    n = UBound(results, 1)
    For i = 1 To n
        Set newRow = tbl.ListRows.Add
        With newRow.Range
            .Cells(1, 1) = results(i, 1)   ' RunDate
            .Cells(1, 2) = results(i, 2)   ' SampleDate
            .Cells(1, 3) = results(i, 3)   ' ActualEC
            .Cells(1, 4) = results(i, 4)   ' ActualVol
            .Cells(1, 5) = results(i, 5)   ' StdPredEC
            .Cells(1, 6) = results(i, 6)   ' StdErrEC
            .Cells(1, 7) = results(i, 7)   ' StdPredVol
            .Cells(1, 8) = results(i, 8)   ' StdErrVol
            .Cells(1, 9) = results(i, 9)   ' EnhPredEC
            .Cells(1, 10) = results(i, 10) ' EnhErrEC
            .Cells(1, 11) = results(i, 11) ' EnhPredVol
            .Cells(1, 12) = results(i, 12) ' EnhErrVol
        End With
    Next i
End Sub

Private Sub ClearSeasonLog(ByVal site As String)
    Dim tbl As ListObject
    Set tbl = GetSeasonLogTable(site)
    If tbl Is Nothing Then Exit Sub
    If Not tbl.DataBodyRange Is Nothing Then tbl.DataBodyRange.Delete
End Sub

' ==== Table Access =============================================================

Private Function GetResultsTable() As ListObject
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(Schema.SHEET_RESULTS)
    If Not ws Is Nothing Then Set GetResultsTable = ws.ListObjects(Schema.TABLE_RESULTS)
    On Error GoTo 0
End Function

Private Function GetCatalogTable() As ListObject
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(Schema.SHEET_CONFIG)
    If Not ws Is Nothing Then Set GetCatalogTable = ws.ListObjects(Schema.TABLE_CATALOG)
    On Error GoTo 0
End Function

Private Function GetSeasonLogTable(ByVal site As String) As ListObject
    Dim ws As Worksheet, tblName As String
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(Schema.SHEET_LOG)
    tblName = Schema.SeasonLogTableName(site)
    If Not ws Is Nothing Then Set GetSeasonLogTable = ws.ListObjects(tblName)
    On Error GoTo 0
End Function

' ==== Helpers ==================================================================

Private Function GetInputVal(ByVal nm As String) As String
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(Schema.SHEET_INPUT)
    If Not ws Is Nothing Then GetInputVal = CStr(ws.Range(nm).Value)
    On Error GoTo 0
End Function

