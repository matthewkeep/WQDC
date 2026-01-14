Option Explicit
' WQOC: Entry point for Water Quality Optimisation Calculator.
' Dependencies: Core, Data, Sim, History, SimLog, Schema, Setup

Public Sub Run()
    ' Main entry point - runs Standard, optionally Enhanced, generates charts
    Dim s As State, cfgStd As Config, cfgEnh As Config
    Dim rStd As Result, rEnh As Result
    Dim runIdStd As String, runIdEnh As String
    Dim site As String, cm As XlCalculation, latestDate As Date
    Dim enhancedMode As Boolean
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

    ' Safety check: warn if running from a date earlier than existing log data
    latestDate = SimLog.GetLatestLogDate(site)
    If latestDate > 0 And cfgStd.StartDate < latestDate Then
        Application.ScreenUpdating = True
        If MsgBox("Start date (" & Format$(cfgStd.StartDate, "dd-mmm") & ") is before existing log data (" & _
                  Format$(latestDate, "dd-mmm") & "). This will overwrite future forecasts." & vbNewLine & vbNewLine & _
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

    ' Show results (Standard result shown, Enhanced noted if run)
    If enhancedMode Then
        ShowResDual rStd, rEnh
    Else
        ShowRes rStd
    End If
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

Private Sub ShowRes(ByRef r As Result)
    Dim msg As String
    If r.TriggerDay = Core.NO_TRIGGER Then
        msg = "No trigger in " & UBound(r.Snaps) & " days." & vbNewLine & _
              "Final volume: " & Format$(r.FinalState.Vol, "0.0") & " ML"
    Else
        msg = "TRIGGER REACHED" & vbNewLine & vbNewLine & _
              "Metric: " & r.TriggerMetric & vbNewLine & _
              "Day: " & r.TriggerDay & vbNewLine & _
              "Date: " & Format$(r.TriggerDate, "dd-mmm-yyyy")
    End If
    MsgBox msg, vbInformation, "WQOC Result"
End Sub

Private Sub ShowResDual(ByRef rStd As Result, ByRef rEnh As Result)
    Dim msg As String
    msg = "STANDARD MODE:" & vbNewLine
    If rStd.TriggerDay = Core.NO_TRIGGER Then
        msg = msg & "  No trigger in " & UBound(rStd.Snaps) & " days" & vbNewLine
    Else
        msg = msg & "  Trigger: " & rStd.TriggerMetric & " on day " & rStd.TriggerDay & vbNewLine
    End If

    msg = msg & vbNewLine & "ENHANCED MODE:" & vbNewLine
    If rEnh.TriggerDay = Core.NO_TRIGGER Then
        msg = msg & "  No trigger in " & UBound(rEnh.Snaps) & " days"
    Else
        msg = msg & "  Trigger: " & rEnh.TriggerMetric & " on day " & rEnh.TriggerDay
    End If

    MsgBox msg, vbInformation, "WQOC Result"
End Sub

' ==== Chart Generation =======================================================

Private Sub GenerateCharts(ByVal site As String, ByRef cfg As Config, ByRef rStd As Result, ByRef rEnh As Result, ByVal hasEnhanced As Boolean)
    ' Draws charts from simulation results - Standard solid, Enhanced dashed
    Dim wsChart As Worksheet
    Dim cht As ChartObject
    Dim n As Long, i As Long
    Dim dates() As Date, volStd() As Double, volEnh() As Double
    Dim ecStd() As Double, ecEnh() As Double
    Dim trigArr() As Double

    On Error Resume Next
    Set wsChart = ThisWorkbook.Worksheets(Schema.SHEET_CHART)
    On Error GoTo 0
    If wsChart Is Nothing Then Exit Sub

    ' Build arrays from Standard result
    n = UBound(rStd.Snaps) + 1
    If n < 1 Then Exit Sub

    ReDim dates(1 To n)
    ReDim volStd(1 To n)
    ReDim ecStd(1 To n)
    For i = 0 To n - 1
        dates(i + 1) = cfg.StartDate + i
        volStd(i + 1) = rStd.Snaps(i).Vol
        ecStd(i + 1) = rStd.Snaps(i).Chem(1)
    Next i

    ' Build Enhanced arrays if available
    If hasEnhanced Then
        ReDim volEnh(1 To n)
        ReDim ecEnh(1 To n)
        For i = 0 To n - 1
            volEnh(i + 1) = rEnh.Snaps(i).Vol
            ecEnh(i + 1) = rEnh.Snaps(i).Chem(1)
        Next i
    End If

    ' Clear existing charts
    For Each cht In wsChart.ChartObjects: cht.Delete: Next cht

    ' Volume chart
    Set cht = wsChart.ChartObjects.Add(Schema.CHART_LEFT_POS, Schema.CHART_TOP_START, _
                                       Schema.CHART_WIDTH, Schema.CHART_HEIGHT_VOLUME)
    With cht.Chart
        .ChartType = xlLine
        ' Standard series
        With .SeriesCollection.NewSeries
            .Name = "Std Volume"
            .XValues = dates
            .Values = volStd
            .Format.Line.ForeColor.RGB = Schema.COLOR_STD_LINE
            .Format.Line.Weight = 2
        End With
        ' Enhanced series
        If hasEnhanced Then
            With .SeriesCollection.NewSeries
                .Name = "Enh Volume"
                .XValues = dates
                .Values = volEnh
                .Format.Line.ForeColor.RGB = Schema.COLOR_ENH_LINE
                .Format.Line.DashStyle = msoLineDash
                .Format.Line.Weight = 2
            End With
        End If
        ' Trigger threshold
        If cfg.TriggerVol > 0 Then
            ReDim trigArr(1 To n)
            For i = 1 To n: trigArr(i) = cfg.TriggerVol: Next i
            With .SeriesCollection.NewSeries
                .Name = "Trigger"
                .XValues = dates
                .Values = trigArr
                .Format.Line.ForeColor.RGB = Schema.COLOR_TRIGGER_LINE
                .Format.Line.DashStyle = msoLineDash
                .Format.Line.Weight = 1.5
            End With
        End If
        .HasTitle = True: .ChartTitle.Text = site & " - Volume"
        .Axes(xlCategory).HasTitle = True: .Axes(xlCategory).AxisTitle.Text = "Date"
        .Axes(xlCategory).TickLabels.NumberFormat = "dd-mmm"
        .Axes(xlValue).HasTitle = True: .Axes(xlValue).AxisTitle.Text = "ML"
    End With

    ' EC chart
    Set cht = wsChart.ChartObjects.Add(Schema.CHART_LEFT_POS, _
        Schema.CHART_TOP_START + Schema.CHART_HEIGHT_VOLUME + Schema.CHART_SPACING, _
        Schema.CHART_WIDTH, Schema.CHART_HEIGHT_METRIC)
    With cht.Chart
        .ChartType = xlLine
        ' Standard series
        With .SeriesCollection.NewSeries
            .Name = "Std EC"
            .XValues = dates
            .Values = ecStd
            .Format.Line.ForeColor.RGB = Schema.COLOR_STD_LINE
            .Format.Line.Weight = 2
        End With
        ' Enhanced series
        If hasEnhanced Then
            With .SeriesCollection.NewSeries
                .Name = "Enh EC"
                .XValues = dates
                .Values = ecEnh
                .Format.Line.ForeColor.RGB = Schema.COLOR_ENH_LINE
                .Format.Line.DashStyle = msoLineDash
                .Format.Line.Weight = 2
            End With
        End If
        ' Trigger threshold
        If cfg.TriggerChem(1) > 0 Then
            ReDim trigArr(1 To n)
            For i = 1 To n: trigArr(i) = cfg.TriggerChem(1): Next i
            With .SeriesCollection.NewSeries
                .Name = "Trigger"
                .XValues = dates
                .Values = trigArr
                .Format.Line.ForeColor.RGB = Schema.COLOR_TRIGGER_LINE
                .Format.Line.DashStyle = msoLineDash
                .Format.Line.Weight = 1.5
            End With
        End If
        .HasTitle = True: .ChartTitle.Text = site & " - EC"
        .Axes(xlCategory).HasTitle = True: .Axes(xlCategory).AxisTitle.Text = "Date"
        .Axes(xlCategory).TickLabels.NumberFormat = "dd-mmm"
        .Axes(xlValue).HasTitle = True: .Axes(xlValue).AxisTitle.Text = "EC (uS/cm)"
    End With
End Sub

' ==== Quick Tests ============================================================

Public Sub TestCore()
    Dim s As State, cfg As Config, r As Result
    s.Vol = 100: s.Chem(1) = 200
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
    s.Vol = 100: s.HidVol = 50: s.Chem(1) = 200: s.Hidden(1) = 5000
    cfg.Mode = "TwoBucket": cfg.Days = 30: cfg.Tau = 7
    cfg.Inflow = 2: cfg.Outflow = 1: cfg.TriggerChem(1) = 300
    r = Sim.Run(s, cfg)

    Debug.Print "Two-bucket: Start EC=" & s.Chem(1) & " End EC=" & r.FinalState.Chem(1)
    If r.TriggerDay <> Core.NO_TRIGGER Then
        Debug.Print "  TRIGGER day " & r.TriggerDay & ": " & r.TriggerMetric
    Else
        Debug.Print "  No trigger in " & cfg.Days & " days"
    End If
End Sub
