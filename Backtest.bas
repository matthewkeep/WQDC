Option Explicit
' Backtest: Season replay for prediction validation.
' Dependencies: Core, Schema, Data, Sim, SimLog, History, Loader, Telemetry, Setup
'
' Runs both Standard and Enhanced modes for A/B comparison:
' - Standard: Simple mixing, independent runs (no state carryover)
' - Enhanced: Uses configured settings, hidden layer carries forward
'
' Writes to tblLive (charts auto-update) and tblHistory (audit trail).
' SeasonLog kept for error metrics analysis.

' ==== Entry Point ==============================================================

Public Sub RunSeason()
    ' Backtests all RR samples for current site using both Standard and Enhanced
    ' Simulates weekly operational workflow: Run Date = Sample Date + 7
    ' Writes to Live/History tables so charts update automatically
    Dim site As String, samples As Variant
    Dim i As Long, n As Long, predictDay As Long, cm As XlCalculation
    Dim sStd As State, sEnh As State, cfgStd As Config, cfgEnh As Config
    Dim rStd As Result, rEnh As Result
    Dim results() As Variant
    Dim enhancedMode As Boolean, telemCalEnabled As Boolean
    Dim wsInput As Worksheet, runId As String, runSeq As Long
    Dim sampleDate As Date, runDate As Date, dayOffset As Long

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

    Set wsInput = Helpers.GetSheet(Schema.SHEET_INPUT)
    If wsInput Is Nothing Then Exit Sub

    On Error GoTo Cleanup
    cm = Application.Calculation
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual

    ' Ensure tables exist and clear season log
    Setup.EnsureSiteLiveTable site
    Setup.EnsureSiteHistoryTable site
    Setup.EnsureSeasonLogTable site
    ClearSeasonLog site

    ' Clear Enhanced columns if running Standard-only (removes stale data from previous Enhanced runs)
    If Not enhancedMode Then SimLog.ClearEnhancedColumns site

    ' Get starting run sequence from history count
    runSeq = History.CountRuns(site)

    ' Results: RunDate, SampleDate, ActualEC, ActualVol, StdPredEC, StdErrEC, StdPredVol, StdErrVol, EnhPredEC, EnhErrEC, EnhPredVol, EnhErrVol
    ReDim results(1 To n - 1, 1 To 12)

    For i = 1 To n - 1
        sampleDate = samples(i, 1)
        runDate = sampleDate + 7  ' Simulates lab delay

        ' === Set dates and refresh data ===
        SetBacktestDates wsInput, sampleDate, runDate
        Loader.LoadLatestInternal site, sampleDate

        ' === Generate RunId ===
        runSeq = runSeq + 1
        runId = site & "_" & Format$(runSeq, "000")

        ' === Load state and config from Inputs sheet ===
        sStd = Data.LoadState()

        ' === Standard Run (Simple mode, no rainfall) ===
        cfgStd = Data.LoadConfig(site, "Standard")
        cfgStd.StartDate = sampleDate
        Helpers.ExtendForecastToRunDate cfgStd, runDate
        rStd = Sim.Run(sStd, cfgStd)
        SimLog.WriteLog rStd, cfgStd, runId, site, "Standard"

        ' === Enhanced Run (if enabled) ===
        If enhancedMode Then
            ' Start with visible layer from observed data
            sEnh = Data.LoadState()

            ' Apply telemetry calibration if enabled
            If telemCalEnabled Then
                sEnh = SnapVisibleLayer(sEnh, site, sampleDate)
            End If

            ' Initialize hidden layer (must happen AFTER LoadState)
            If i = 1 Then
                ' First run: initialize at equilibrium
                sEnh = Core.InitHiddenAtEquilibrium(sEnh)
            Else
                ' Subsequent runs: carry hidden from previous run at actual time offset
                ' (not from day 100 - use the day matching real elapsed time)
                dayOffset = samples(i, 1) - samples(i - 1, 1)
                If dayOffset < 0 Then dayOffset = 0
                If dayOffset > UBound(rEnh.Snaps) Then dayOffset = UBound(rEnh.Snaps)
                sEnh = CarryHiddenFromPrevious(sEnh, rEnh.Snaps(dayOffset))
            End If

            cfgEnh = Data.LoadConfig(site, "Enhanced")
            cfgEnh.StartDate = sampleDate
            Helpers.ExtendForecastToRunDate cfgEnh, runDate
            rEnh = Sim.Run(sEnh, cfgEnh)
            SimLog.WriteLog rEnh, cfgEnh, runId, site, "Enhanced"
        End If

        ' === Record to History ===
        History.RecordRun sStd, cfgStd, rStd, cfgEnh, rEnh, enhancedMode, telemCalEnabled, runId, site

        ' === Calculate error metrics for SeasonLog ===
        ' predictDay is snap index (days from sampleDate/StartDate, not runDate)
        predictDay = CLng(samples(i + 1, 1) - sampleDate)
        If predictDay < 0 Then predictDay = 0
        If predictDay > UBound(rStd.Snaps) Then predictDay = UBound(rStd.Snaps)

        results(i, 1) = runDate                                      ' RunDate
        results(i, 2) = sampleDate                                   ' SampleDate
        results(i, 3) = samples(i + 1, 2)                            ' ActualEC (next sample)
        results(i, 4) = samples(i + 1, 3)                            ' ActualVol (next sample)

        ' Standard predictions
        results(i, 5) = rStd.Snaps(predictDay).Chem(mEC)             ' StdPredEC
        results(i, 6) = results(i, 5) - results(i, 3)                ' StdErrEC
        results(i, 7) = rStd.Snaps(predictDay).Vol                   ' StdPredVol
        results(i, 8) = results(i, 7) - results(i, 4)                ' StdErrVol

        ' Enhanced predictions (if enabled)
        If enhancedMode Then
            results(i, 9) = rEnh.Snaps(predictDay).Chem(mEC)         ' EnhPredEC
            results(i, 10) = results(i, 9) - results(i, 3)           ' EnhErrEC
            results(i, 11) = rEnh.Snaps(predictDay).Vol              ' EnhPredVol
            results(i, 12) = results(i, 11) - results(i, 4)          ' EnhErrVol
        Else
            results(i, 9) = Empty: results(i, 10) = Empty
            results(i, 11) = Empty: results(i, 12) = Empty
        End If
    Next i

    WriteSeasonLog site, results

    ' Generate/update charts
    WQOC.GenerateCharts site, cfgStd, enhancedMode

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
    msg = msg & vbNewLine & vbNewLine & "Results written to:" & vbNewLine
    msg = msg & "- tblLive_" & site & " (charts)" & vbNewLine
    msg = msg & "- tblHistory_" & site & " (audit)" & vbNewLine
    msg = msg & "- tblSeasonLog_" & site & " (errors)"
    MsgBox msg, vbInformation, "Backtest"
    Exit Sub

Cleanup:
    Application.Calculation = cm
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    If Err.Number <> 0 Then
        Error.TraceErr "Backtest.RunSeason"
        MsgBox "Error: " & Err.Description, vbExclamation, "Backtest"
    End If
End Sub

' ==== Hidden Layer Management ==================================================

Private Function CarryHiddenFromPrevious(ByRef current As State, ByRef previous As State) As State
    ' Copies hidden layer from previous state to current (visible unchanged)
    Dim result As State, i As Long
    result = Core.CopyState(current)
    For i = 1 To Core.METRIC_COUNT
        result.Hidden(i) = previous.Hidden(i)
    Next i
    CarryHiddenFromPrevious = result
End Function

Private Function SnapVisibleLayer(ByRef s As State, ByVal site As String, ByVal sampleDate As Date) As State
    ' Snaps visible layer to telemetry values (hidden unchanged)
    Dim snapped As State, latestVol As Variant, latestEC As Variant
    snapped = Core.CopyState(s)

    latestVol = Telemetry.GetLatestVol(sampleDate, site)
    latestEC = Telemetry.GetLatestEC(sampleDate, site)

    If Not IsEmpty(latestVol) Then snapped.Vol = CDbl(latestVol)
    If Not IsEmpty(latestEC) Then snapped.Chem(mEC) = CDbl(latestEC)

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
        If Helpers.MatchesSite(row.Range.Cells(1, 1).Value, site) Then
            On Error Resume Next
            sampleDate = CDate(row.Range.Cells(1, 2).Value)
            ec = Val(row.Range.Cells(1, Helpers.ColIdx(tbl, Schema.ChemistryNames()(0))).Value)
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
        arr = dict.Item(k)
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

Private Function GetSeasonLogTable(ByVal site As String) As ListObject
    Set GetSeasonLogTable = Helpers.GetSiteTable(Schema.SHEET_LOG, Schema.SEASONLOG_TABLE_PREFIX, site)
End Function

' ==== Helpers ==================================================================

Private Sub SetBacktestDates(ByVal ws As Worksheet, ByVal sampleDate As Date, ByVal runDate As Date)
    ' Sets Sample Date and Run Date in Inputs sheet for backtest iteration
    On Error Resume Next
    ws.Range(Schema.NAME_SAMPLE_DATE).Value = sampleDate
    ws.Range(Schema.NAME_RUN_DATE).Value = runDate
    On Error GoTo 0
End Sub

Private Function GetInputVal(ByVal nm As String) As String
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(Schema.SHEET_INPUT)
    If Not ws Is Nothing Then GetInputVal = CStr(ws.Range(nm).Value)
    On Error GoTo 0
End Function
