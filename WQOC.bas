Option Explicit
' WQOC: Entry point for Water Quality Optimisation Calculator.
' Dependencies: Core, Data, Sim, History, SimLog, Schema, Setup, Validate

' ==== Entry Points ============================================================

Public Sub Run()
    ' Main entry point - runs Standard, optionally Enhanced, generates charts
    ' Creates new history entry
    RunCore "", True
End Sub

Public Sub Replay()
    ' Regenerates simulation output using existing runId from history
    ' Used after rollback - does not create new history entry
    Dim site As String, runId As String
    site = Data.GetSite()
    If Len(site) = 0 Then Exit Sub
    runId = History.GetCurrentRunId(site)
    If Len(runId) = 0 Then Exit Sub
    RunCore runId, False
End Sub

' ==== Private Implementation ==================================================

Private Sub RunCore(ByVal existingRunId As String, ByVal recordHistory As Boolean)
    ' Core simulation logic shared by Run and Replay
    Dim s As State, logState As State, cfgStd As Config, cfgEnh As Config
    Dim rStd As Result, rEnh As Result
    Dim runId As String
    Dim site As String, cm As XlCalculation
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

    ' Use existing runId (replay) or create new one
    If Len(existingRunId) > 0 Then
        runId = existingRunId
    Else
        runId = MakeRunId(site)
    End If

    ' Load state and Standard config
    s = Data.LoadState()
    cfgStd = Data.LoadConfig(site, "Standard")

    ' Validate: Run date must not be before sample date
    Dim runDate As Date, wsInput As Worksheet
    Set wsInput = Helpers.GetSheet(Schema.SHEET_INPUT)
    If Not wsInput Is Nothing Then
        runDate = Helpers.GetDateVal(wsInput, Schema.NAME_RUN_DATE)
        If runDate > 0 And cfgStd.StartDate > 0 And runDate < cfgStd.StartDate Then
            MsgBox "Run date (" & Format$(runDate, "d-mmm-yyyy") & ") cannot be before sample date (" & _
                   Format$(cfgStd.StartDate, "d-mmm-yyyy") & ").", vbExclamation, "WQOC"
            GoTo Cleanup
        End If
    End If

    ' Run Standard simulation
    rStd = Sim.Run(s, cfgStd)
    SimLog.WriteLog rStd, cfgStd, "STD-" & runId, site
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
            If logState.Hidden(mEC) > Core.EPS Then
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
        SimLog.WriteLog rEnh, cfgEnh, "ENH-" & runId, site
        Data.SaveResult rEnh, "Enhanced"
    End If

    ' Record history entry only for new runs (not replay)
    If recordHistory Then
        History.RecordRun cfgStd, rStd, cfgEnh, rEnh, enhancedMode, runId, site
    End If

    ' Generate charts for site
    GenerateCharts site, cfgStd, enhancedMode

    Application.Calculation = cm
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Exit Sub

Cleanup:
    Application.Calculation = cm
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    If Err.Number <> 0 Then
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

Private Function MakeRunId(ByVal site As String) As String
    ' Creates run ID: RP1-20260114-001 (one per run, captures both Std and Enh)
    Dim baseId As String, seq As Long
    baseId = site & "-" & Format$(Now, "yyyymmdd")
    seq = History.CountRuns(site) + 1
    MakeRunId = baseId & "-" & Format$(seq, "000")
End Function

' ==== Chart Generation =======================================================

Private Sub GenerateCharts(ByVal site As String, ByRef cfg As Config, ByVal hasEnhanced As Boolean)
    ' Generates 7 charts (one per chemistry metric) bound to table columns
    ' EC chart: Dual-axis with Volume; other charts: single-analyte only
    ' Charts created once per site, auto-update when table data changes
    Dim wsChart As Worksheet, tbl As ListObject
    Dim cht As ChartObject, chemIdx As Long
    Dim chartLeft As Double, chartTop As Double

    Set wsChart = Helpers.GetSheet(Schema.SHEET_CHART)
    If wsChart Is Nothing Then Exit Sub

    Set tbl = Helpers.GetTable(Schema.SHEET_LOG, Helpers.LiveTableName(site))
    If tbl Is Nothing Or tbl.DataBodyRange Is Nothing Then Exit Sub

    chartLeft = GetSiteChartLeft(wsChart, site)
    chartTop = Schema.CHART_TOP_START

    For chemIdx = 1 To Schema.ChemistryCount()
        Set cht = GetOrCreateChart(wsChart, site, chemIdx, chartLeft, chartTop)
        If ChartNeedsSeries(cht) Then
            BuildChartSeries cht.Chart, tbl, site, chemIdx, cfg, hasEnhanced
        Else
            UpdateChartRanges cht.Chart, tbl, chemIdx, hasEnhanced
        End If
        chartTop = chartTop + Schema.CHART_HEIGHT + Schema.CHART_SPACING
    Next chemIdx
End Sub

Private Function GetOrCreateChart(ByVal ws As Worksheet, ByVal site As String, _
                                  ByVal chemIdx As Long, ByVal left As Double, _
                                  ByVal top As Double) As ChartObject
    Dim chartName As String
    chartName = "cht_" & site & "_" & Schema.ChemShortName(chemIdx)

    On Error Resume Next
    Set GetOrCreateChart = ws.ChartObjects(chartName)
    On Error GoTo 0

    If GetOrCreateChart Is Nothing Then
        Set GetOrCreateChart = ws.ChartObjects.Add(left, top, _
                                   Schema.CHART_WIDTH, Schema.CHART_HEIGHT)
        GetOrCreateChart.Name = chartName
    End If
End Function

Private Function ChartNeedsSeries(ByVal cht As ChartObject) As Boolean
    ChartNeedsSeries = (cht.Chart.SeriesCollection.Count = 0)
End Function

Private Function GetSiteChartLeft(ByVal ws As Worksheet, ByVal site As String) As Double
    Dim cht As ChartObject, prefix As String
    Dim maxRight As Double, siteLeft As Double

    prefix = "cht_" & site & "_"
    siteLeft = -1
    maxRight = 0

    For Each cht In ws.ChartObjects
        If Left$(cht.Name, Len(prefix)) = prefix Then siteLeft = cht.left
        If cht.left + cht.Width > maxRight Then maxRight = cht.left + cht.Width
    Next cht

    If siteLeft >= 0 Then
        GetSiteChartLeft = siteLeft
    ElseIf maxRight > 0 Then
        GetSiteChartLeft = maxRight + Schema.CHART_SPACING
    Else
        GetSiteChartLeft = Schema.CHART_LEFT_POS
    End If
End Function

' ==== Chart Series Creation ==================================================

Private Sub BuildChartSeries(ByVal cht As Chart, ByVal tbl As ListObject, _
                             ByVal site As String, ByVal chemIdx As Long, _
                             ByRef cfg As Config, ByVal hasEnhanced As Boolean)
    ' Creates all series for a chart from table columns
    Dim chemName As String, chemUnit As String, includeVol As Boolean
    Dim dateRng As Range

    chemName = Schema.ChemShortName(chemIdx)
    chemUnit = Schema.ChemistryNames()(chemIdx - 1)
    includeVol = (chemIdx = 1)  ' EC only
    Set dateRng = GetColRange(tbl, Schema.LIVE_COL_DATE)

    cht.ChartType = xlLine

    ' Chemistry series (left Y-axis)
    AddDataSeries cht, "Std " & chemName, dateRng, _
                  GetColRange(tbl, Schema.StdChemColName(chemIdx)), _
                  Schema.COLOR_STD_LINE, False, xlPrimary
    If hasEnhanced Then
        AddDataSeries cht, "Enh " & chemName, dateRng, _
                      GetColRange(tbl, Schema.EnhChemColName(chemIdx)), _
                      Schema.COLOR_ENH_LINE, False, xlPrimary
    End If
    If cfg.TriggerChem(chemIdx) > 0 Then
        AddTriggerLine cht, chemName & " Trigger", dateRng, cfg.TriggerChem(chemIdx), xlPrimary
    End If

    ' Volume series (right Y-axis, EC only)
    If includeVol Then
        AddDataSeries cht, "Std Vol", dateRng, _
                      GetColRange(tbl, Schema.LIVE_COL_STD_VOL), _
                      Schema.COLOR_STD_LINE, True, xlSecondary
        If hasEnhanced Then
            AddDataSeries cht, "Enh Vol", dateRng, _
                          GetColRange(tbl, Schema.LIVE_COL_ENH_VOL), _
                          Schema.COLOR_ENH_LINE, True, xlSecondary
        End If
        If cfg.TriggerVol > 0 Then
            AddTriggerLine cht, "Vol Trigger", dateRng, cfg.TriggerVol, xlSecondary
        End If
    End If

    FormatChart cht, site, chemName, chemUnit, includeVol
End Sub

Private Sub AddDataSeries(ByVal cht As Chart, ByVal seriesName As String, _
                          ByVal xRng As Range, ByVal yRng As Range, _
                          ByVal lineColor As Long, ByVal dashed As Boolean, _
                          ByVal axisGroup As XlAxisGroup)
    If yRng Is Nothing Then Exit Sub
    With cht.SeriesCollection.NewSeries
        .Name = seriesName
        .XValues = xRng
        .Values = yRng
        .Format.Line.ForeColor.RGB = lineColor
        .Format.Line.Weight = Schema.CHART_LINE_WEIGHT
        If dashed Then .Format.Line.DashStyle = msoLineDash
        .AxisGroup = axisGroup
    End With
End Sub

Private Sub AddTriggerLine(ByVal cht As Chart, ByVal seriesName As String, _
                           ByVal dateRng As Range, ByVal triggerVal As Double, _
                           ByVal axisGroup As XlAxisGroup)
    Dim trigArr() As Double, i As Long, n As Long
    n = dateRng.Rows.Count
    ReDim trigArr(1 To n)
    For i = 1 To n: trigArr(i) = triggerVal: Next i

    With cht.SeriesCollection.NewSeries
        .Name = seriesName
        .XValues = dateRng
        .Values = trigArr
        .Format.Line.ForeColor.RGB = Schema.COLOR_TRIGGER_LINE
        .Format.Line.DashStyle = msoLineDashDot
        .Format.Line.Weight = Schema.CHART_TRIGGER_WEIGHT
        .AxisGroup = axisGroup
    End With
End Sub

Private Sub FormatChart(ByVal cht As Chart, ByVal site As String, _
                        ByVal chemName As String, ByVal chemUnit As String, _
                        ByVal includeVol As Boolean)
    With cht
        .HasTitle = True
        .ChartTitle.Text = site & " - " & chemName & IIf(includeVol, " + Volume", "")
        .Axes(xlCategory).TickLabels.NumberFormat = "d/mm/yy"
        .Axes(xlValue, xlPrimary).HasTitle = True
        .Axes(xlValue, xlPrimary).AxisTitle.Text = chemUnit
        If includeVol Then
            .Axes(xlValue, xlSecondary).HasTitle = True
            .Axes(xlValue, xlSecondary).AxisTitle.Text = "Volume (ML)"
        End If
        .HasLegend = True
        .Legend.Position = xlLegendPositionBottom
    End With
End Sub

' ==== Chart Update ===========================================================

Private Sub UpdateChartRanges(ByVal cht As Chart, ByVal tbl As ListObject, _
                              ByVal chemIdx As Long, ByVal hasEnhanced As Boolean)
    ' Updates existing chart series to current table ranges
    Dim ser As Series, nm As String
    Dim dateRng As Range, stdChemRng As Range, enhChemRng As Range
    Dim stdVolRng As Range, enhVolRng As Range

    Set dateRng = GetColRange(tbl, Schema.LIVE_COL_DATE)
    Set stdChemRng = GetColRange(tbl, Schema.StdChemColName(chemIdx))
    Set enhChemRng = GetColRange(tbl, Schema.EnhChemColName(chemIdx))
    Set stdVolRng = GetColRange(tbl, Schema.LIVE_COL_STD_VOL)
    Set enhVolRng = GetColRange(tbl, Schema.LIVE_COL_ENH_VOL)

    On Error Resume Next
    For Each ser In cht.SeriesCollection
        ser.XValues = dateRng
        nm = ser.Name
        Select Case True
            Case nm Like "Std *" And InStr(nm, "Vol") > 0: ser.Values = stdVolRng
            Case nm Like "Std *": ser.Values = stdChemRng
            Case nm Like "Enh *" And hasEnhanced And InStr(nm, "Vol") > 0: ser.Values = enhVolRng
            Case nm Like "Enh *" And hasEnhanced: ser.Values = enhChemRng
        End Select
    Next ser
    On Error GoTo 0
End Sub

Private Function GetColRange(ByVal tbl As ListObject, ByVal colName As String) As Range
    Dim col As ListColumn
    On Error Resume Next
    Set col = tbl.ListColumns(colName)
    On Error GoTo 0
    If Not col Is Nothing Then Set GetColRange = col.DataBodyRange
End Function

