Option Explicit
' WQOC: Entry point for Water Quality Optimisation Calculator.
' Dependencies: Core, Data, Sim, History, SimLog, Schema, Setup, Validate

Public Sub Run()
    ' Main entry point - runs Standard, optionally Enhanced, generates charts
    Dim s As State, logState As State, cfgStd As Config, cfgEnh As Config
    Dim rStd As Result, rEnh As Result
    Dim runIdStd As String, runIdEnh As String
    Dim site As String, cm As XlCalculation, latestDate As Date
    Dim enhancedMode As Boolean, i As Long

    ' Pre-flight validation
    If Not Validate.Check() Then
        MsgBox "Structure validation failed. Run Validate.Report for details.", vbExclamation, "WQOC"
        Exit Sub
    End If

    On Error GoTo Cleanup

    cm = Application.Calculation
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual

    ' Get current site and ensure tables exist
    site = Data.GetSite()
    If Len(site) = 0 Then
        MsgBox "No site selected.", vbExclamation, "WQOC"
        GoTo Cleanup
    End If
    Setup.EnsureSiteTables site

    ' Load state and Standard config
    s = Data.LoadState()
    cfgStd = Data.LoadConfig(site, "Standard")

    ' Validate state
    If s.Vol < Core.EPS Then
        MsgBox "Initial volume must be greater than 0.", vbExclamation, "WQOC"
        GoTo Cleanup
    End If

    ' Validate config
    If cfgStd.Days < 1 Then
        MsgBox "Forecast days must be at least 1.", vbExclamation, "WQOC"
        GoTo Cleanup
    End If

    ' Safety check: warn if running from a date earlier than existing log data
    latestDate = SimLog.GetLatestLogDate(site)
    If latestDate > 0 And cfgStd.StartDate < latestDate Then
        Application.ScreenUpdating = True
        If MsgBox("Start date (" & Format$(cfgStd.StartDate, "d/mm/yy") & ") is before existing log data (" & _
                  Format$(latestDate, "d/mm/yy") & "). This will overwrite future forecasts." & vbNewLine & vbNewLine & _
                  "Continue?", vbYesNo + vbQuestion, "WQOC") = vbNo Then
            GoTo Cleanup
        End If
        Application.ScreenUpdating = False
    End If

    ' Run Standard simulation
    rStd = Sim.Run(s, cfgStd)
    runIdStd = MakeRunId("STD", site)
    SimLog.WriteLog rStd, cfgStd, runIdStd, site
    History.RecordRun cfgStd, rStd, runIdStd, site
    Data.SaveResult rStd, "Standard"

    ' Check if Enhanced mode is enabled
    enhancedMode = (UCase$(Data.GetEnhancedMode()) = "ON")

    ' Run Enhanced if enabled
    If enhancedMode Then
        cfgEnh = Data.LoadConfig(site, "Enhanced")

        ' Apply telemetry calibration (snap to latest observed values) if enabled
        If Data.GetTelemCalEnabled() Then
            s = Data.SnapState(s, site)
        End If

        ' Load hidden layer from log for TwoBucket continuity
        ' Priority: 1) Log at sample date, 2) Inputs sheet, 3) Initialize at equilibrium
        If cfgEnh.Mode = "TwoBucket" Then
            logState = Data.LoadHiddenFromLog(site, cfgEnh.StartDate)
            If Not Core.IsHiddenEmpty(logState) Then
                ' Found hidden state in log - use it
                For i = 1 To Core.METRIC_COUNT
                    s.Hidden(i) = logState.Hidden(i)
                Next i
            ElseIf Core.IsHiddenEmpty(s) Then
                ' No log data and Inputs sheet empty - initialize at equilibrium
                s = Core.InitHiddenAtEquilibrium(s)
            End If
            ' Else: use hidden state from Inputs sheet (LoadState already loaded it)
        End If

        rEnh = Sim.Run(s, cfgEnh)
        runIdEnh = MakeRunId("ENH", site)
        SimLog.WriteLog rEnh, cfgEnh, runIdEnh, site
        History.RecordRun cfgEnh, rEnh, runIdEnh, site
        Data.SaveResult rEnh, "Enhanced"
    End If

    ' Generate charts for site
    GenerateCharts site, cfgStd, rStd, rEnh, enhancedMode

    Application.Calculation = cm
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Exit Sub

Cleanup:
    Application.Calculation = cm
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    If Err.Number <> 0 Then
        Error.TraceErr "WQOC.Run"
        MsgBox "Error: " & Err.Description, vbExclamation, "WQOC"
    End If
End Sub

Public Sub Rollback()
    Dim site As String
    site = Data.GetSite()
    If Len(site) = 0 Then
        MsgBox "No site selected.", vbExclamation, "WQOC"
        Exit Sub
    End If
    If History.RollbackLast(site) Then
        MsgBox "Last run rolled back.", vbInformation, "WQOC"
    Else
        MsgBox "No run to rollback.", vbExclamation, "WQOC"
    End If
End Sub

Public Sub ShowCnt()
    Dim site As String
    site = Data.GetSite()
    If Len(site) = 0 Then
        MsgBox "No site selected.", vbExclamation, "WQOC"
        Exit Sub
    End If
    MsgBox "Runs for " & site & ": " & History.CountRuns(site), vbInformation, "WQOC"
End Sub

Private Function MakeRunId(ByVal prefix As String, ByVal site As String) As String
    ' Creates run ID: STD-RP1-20260114-001 or ENH-RP1-20260114-001
    Dim baseId As String, seq As Long
    baseId = prefix & "-" & site & "-" & Format$(Now, "yyyymmdd")
    seq = History.CountRuns(site) + 1
    MakeRunId = baseId & "-" & Format$(seq, "000")
End Function

Public Function GetTrigDay() As Long
    Dim s As State, cfg As Config, r As Result, site As String
    site = Data.GetSite()
    If Len(site) = 0 Then Exit Function
    s = Data.LoadState()
    cfg = Data.LoadConfig(site, "Standard")
    r = Sim.Run(s, cfg)
    GetTrigDay = r.TriggerDay
End Function

' ==== Chart Generation =======================================================

Private Sub GenerateCharts(ByVal site As String, ByRef cfg As Config, ByRef rStd As Result, ByRef rEnh As Result, ByVal hasEnhanced As Boolean)
    ' Generates 7 dual-axis charts (one per chemistry metric)
    ' Each chart: Chemistry on left Y-axis, Volume on right Y-axis
    ' Styling: Std=Blue, Enh=Teal, Chemistry=Solid, Volume=Dashed
    Dim wsChart As Worksheet, wsLog As Worksheet
    Dim tbl As ListObject
    Dim cht As ChartObject
    Dim n As Long, i As Long, chemIdx As Long
    Dim dates() As Date, volStd() As Double, volEnh() As Double
    Dim chemStd() As Double, chemEnh() As Double
    Dim trigArr() As Double
    Dim stdVolCol As Long, enhVolCol As Long
    Dim stdChemCol As Long, enhChemCol As Long
    Dim hasEnhData As Boolean
    Dim chartTop As Double
    Dim chemName As String, chemUnit As String

    On Error Resume Next
    Set wsChart = ThisWorkbook.Worksheets(Schema.SHEET_CHART)
    Set wsLog = ThisWorkbook.Worksheets(Schema.SHEET_LOG)
    On Error GoTo 0
    If wsChart Is Nothing Then Error.Trace "GenerateCharts", "Chart sheet missing": Exit Sub
    If wsLog Is Nothing Then Error.Trace "GenerateCharts", "Log sheet missing": Exit Sub

    ' Get live table for site
    On Error Resume Next
    Set tbl = wsLog.ListObjects(Helpers.LiveTableName(site))
    On Error GoTo 0
    If tbl Is Nothing Then Error.Trace "GenerateCharts", "No live table for " & site: Exit Sub
    If tbl.DataBodyRange Is Nothing Then Error.Trace "GenerateCharts", "Empty live table": Exit Sub

    n = tbl.ListRows.Count
    If n < 1 Then Error.Trace "GenerateCharts", "No data rows": Exit Sub

    ' Get volume column indices
    stdVolCol = Helpers.ColIdx(tbl, Schema.LIVE_COL_STD_VOL)
    enhVolCol = Helpers.ColIdx(tbl, Schema.LIVE_COL_ENH_VOL)

    ' Build date and volume arrays (shared across all charts)
    ReDim dates(1 To n)
    ReDim volStd(1 To n)
    ReDim volEnh(1 To n)

    For i = 1 To n
        dates(i) = tbl.DataBodyRange.Cells(i, 1).Value
        If stdVolCol > 0 Then volStd(i) = Val(tbl.DataBodyRange.Cells(i, stdVolCol).Value)
        If enhVolCol > 0 Then
            volEnh(i) = Val(tbl.DataBodyRange.Cells(i, enhVolCol).Value)
            If volEnh(i) > 0 Then hasEnhData = True
        End If
    Next i

    ' Clear existing charts
    For Each cht In wsChart.ChartObjects: cht.Delete: Next cht

    ' Generate one chart per chemistry metric
    chartTop = Schema.CHART_TOP_START
    For chemIdx = 1 To Schema.ChemistryCount()
        ' Get chemistry column indices
        stdChemCol = Helpers.ColIdx(tbl, Schema.StdChemColName(chemIdx))
        enhChemCol = Helpers.ColIdx(tbl, Schema.EnhChemColName(chemIdx))

        ' Build chemistry arrays
        ReDim chemStd(1 To n)
        ReDim chemEnh(1 To n)
        For i = 1 To n
            If stdChemCol > 0 Then chemStd(i) = Val(tbl.DataBodyRange.Cells(i, stdChemCol).Value)
            If enhChemCol > 0 Then chemEnh(i) = Val(tbl.DataBodyRange.Cells(i, enhChemCol).Value)
        Next i

        ' Get chemistry name and unit for axis labels
        chemName = Schema.ChemShortName(chemIdx)
        chemUnit = Schema.ChemistryNames()(chemIdx - 1)  ' Full name with unit

        ' Create dual-axis chart
        Set cht = wsChart.ChartObjects.Add(Schema.CHART_LEFT_POS, chartTop, _
                                           Schema.CHART_WIDTH, Schema.CHART_HEIGHT_METRIC)
        With cht.Chart
            .ChartType = xlLine

            ' === LEFT Y-AXIS: Chemistry (Solid lines) ===
            ' Standard chemistry series
            With .SeriesCollection.NewSeries
                .Name = "Std " & chemName
                .XValues = dates
                .Values = chemStd
                .Format.Line.ForeColor.RGB = Schema.COLOR_STD_LINE
                .Format.Line.Weight = 2
            End With
            ' Enhanced chemistry series (if data exists)
            If hasEnhData Then
                With .SeriesCollection.NewSeries
                    .Name = "Enh " & chemName
                    .XValues = dates
                    .Values = chemEnh
                    .Format.Line.ForeColor.RGB = Schema.COLOR_ENH_LINE
                    .Format.Line.Weight = 2
                End With
            End If
            ' Chemistry trigger threshold (left axis)
            If cfg.TriggerChem(chemIdx) > 0 Then
                ReDim trigArr(1 To n)
                For i = 1 To n: trigArr(i) = cfg.TriggerChem(chemIdx): Next i
                With .SeriesCollection.NewSeries
                    .Name = chemName & " Trigger"
                    .XValues = dates
                    .Values = trigArr
                    .Format.Line.ForeColor.RGB = Schema.COLOR_TRIGGER_LINE
                    .Format.Line.DashStyle = msoLineDashDot
                    .Format.Line.Weight = 1.5
                End With
            End If

            ' === RIGHT Y-AXIS: Volume (Dashed lines) ===
            ' Standard volume series
            With .SeriesCollection.NewSeries
                .Name = "Std Vol"
                .XValues = dates
                .Values = volStd
                .Format.Line.ForeColor.RGB = Schema.COLOR_STD_LINE
                .Format.Line.DashStyle = msoLineDash
                .Format.Line.Weight = 2
                .AxisGroup = xlSecondary
            End With
            ' Enhanced volume series (if data exists)
            If hasEnhData Then
                With .SeriesCollection.NewSeries
                    .Name = "Enh Vol"
                    .XValues = dates
                    .Values = volEnh
                    .Format.Line.ForeColor.RGB = Schema.COLOR_ENH_LINE
                    .Format.Line.DashStyle = msoLineDash
                    .Format.Line.Weight = 2
                    .AxisGroup = xlSecondary
                End With
            End If
            ' Volume trigger threshold (right axis)
            If cfg.TriggerVol > 0 Then
                ReDim trigArr(1 To n)
                For i = 1 To n: trigArr(i) = cfg.TriggerVol: Next i
                With .SeriesCollection.NewSeries
                    .Name = "Vol Trigger"
                    .XValues = dates
                    .Values = trigArr
                    .Format.Line.ForeColor.RGB = Schema.COLOR_TRIGGER_LINE
                    .Format.Line.DashStyle = msoLineDashDot
                    .Format.Line.Weight = 1.5
                    .AxisGroup = xlSecondary
                End With
            End If

            ' Chart formatting
            .HasTitle = True: .ChartTitle.Text = site & " - " & chemName & " + Volume"
            .Axes(xlCategory).TickLabels.NumberFormat = "d/mm/yy"
            .Axes(xlValue, xlPrimary).HasTitle = True
            .Axes(xlValue, xlPrimary).AxisTitle.Text = chemUnit
            .Axes(xlValue, xlSecondary).HasTitle = True
            .Axes(xlValue, xlSecondary).AxisTitle.Text = "Volume (ML)"
            .HasLegend = True
            .Legend.Position = xlLegendPositionBottom
        End With

        chartTop = chartTop + Schema.CHART_HEIGHT_METRIC + Schema.CHART_SPACING
    Next chemIdx
End Sub

' ==== Quick Tests ============================================================

Public Sub TestCore()
    Dim s As State, cfg As Config, r As Result
    s.Vol = 100: s.Chem(mEC) = 200
    cfg.Mode = "Simple": cfg.Days = 50: cfg.Inflow = 2: cfg.Outflow = 1: cfg.TriggerVol = 150
    r = Sim.Run(s, cfg)

    If r.TriggerDay = Core.NO_TRIGGER Then
        Debug.Print "No trigger. Final vol: " & r.FinalState.Vol & " ML"
    Else
        Debug.Print "TRIGGER day " & r.TriggerDay & ": " & r.TriggerMetric
        Debug.Print "  Final vol: " & r.FinalState.Vol & " ML"
    End If
End Sub

Public Sub TestTwoBucket()
    Dim s As State, cfg As Config, r As Result
    s.Vol = 100: s.Chem(mEC) = 200: s.Hidden(mEC) = 5000
    cfg.Mode = "TwoBucket": cfg.Days = 30: cfg.Tau = 7
    cfg.Inflow = 2: cfg.Outflow = 1: cfg.TriggerChem(mEC) = 300
    r = Sim.Run(s, cfg)

    Debug.Print "Two-bucket: Start EC=" & s.Chem(mEC) & " End EC=" & r.FinalState.Chem(mEC)
    If r.TriggerDay <> Core.NO_TRIGGER Then
        Debug.Print "  TRIGGER day " & r.TriggerDay & ": " & r.TriggerMetric
    Else
        Debug.Print "  No trigger in " & cfg.Days & " days"
    End If
End Sub

