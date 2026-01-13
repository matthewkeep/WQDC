Option Explicit
' WQOC: Entry point for Water Quality Optimisation Calculator.
' Dependencies: Core, Data, Sim, History, SimLog, Schema

Public Sub Run()
    Dim s As State, cfg As Config, r As Result, cm As XlCalculation
    Dim runId As String
    On Error GoTo Cleanup

    cm = Application.Calculation
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual

    s = Data.LoadState()
    cfg = Data.LoadConfig()
    r = Sim.Run(s, cfg)
    Data.SaveResult r
    runId = Format$(Now, "yyyymmdd_hhmmss")
    History.RecordRun cfg, r, runId
    SimLog.WriteLog r, cfg, runId
    GenerateCharts r, cfg

    Application.Calculation = cm
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    ShowRes r
    Exit Sub

Cleanup:
    Application.Calculation = cm
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    MsgBox "Error: " & Err.Description, vbExclamation, "WQOC"
End Sub

Public Sub Rollback()
    If History.RollbackLast() Then
        MsgBox "Last run rolled back.", vbInformation, "WQOC"
    Else
        MsgBox "No run to rollback.", vbExclamation, "WQOC"
    End If
End Sub

Public Sub ShowCnt()
    MsgBox "Runs for this site: " & History.CountRuns(), vbInformation, "WQOC"
End Sub

Public Function GetTrigDay() As Long
    Dim s As State, cfg As Config, r As Result
    s = Data.LoadState()
    cfg = Data.LoadConfig()
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

' ==== Chart Generation =======================================================

Private Sub GenerateCharts(ByRef r As Result, ByRef cfg As Config)
    Dim ws As Worksheet, cht As ChartObject, i As Long, n As Long

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(Schema.SHEET_CHART)
    On Error GoTo 0
    If ws Is Nothing Then Exit Sub

    n = UBound(r.Snaps)

    ' Clear existing charts and data
    For Each cht In ws.ChartObjects: cht.Delete: Next cht
    ws.Cells.Clear

    ' Write data with dates and trigger thresholds
    ws.Range("A1") = "Date"
    ws.Range("B1") = "Volume (ML)"
    ws.Range("C1") = "Vol Trigger"
    ws.Range("D1") = "EC"
    ws.Range("E1") = "EC Trigger"

    For i = 0 To n
        ws.Cells(i + 2, 1) = cfg.StartDate + i           ' Actual date
        ws.Cells(i + 2, 2) = r.Snaps(i).Vol             ' Volume
        ws.Cells(i + 2, 3) = cfg.TriggerVol             ' Volume trigger (horizontal line)
        ws.Cells(i + 2, 4) = r.Snaps(i).Chem(1)         ' EC
        ws.Cells(i + 2, 5) = cfg.TriggerChem(1)         ' EC trigger (horizontal line)
    Next i

    ' Format date column
    ws.Range("A2:A" & n + 2).NumberFormat = "dd-mmm-yy"

    ' Volume chart with trigger threshold
    Set cht = ws.ChartObjects.Add(Schema.CHART_LEFT_POS, Schema.CHART_TOP_START, _
                                   Schema.CHART_WIDTH, Schema.CHART_HEIGHT_VOLUME)
    With cht.Chart
        .ChartType = xlLine
        ' Volume series
        .SetSourceData ws.Range("B1:B" & n + 2)
        .SeriesCollection(1).XValues = ws.Range("A2:A" & n + 2)
        .SeriesCollection(1).Name = "Volume"
        ' Trigger threshold series
        If cfg.TriggerVol > 0 Then
            With .SeriesCollection.NewSeries
                .Name = "Trigger"
                .XValues = ws.Range("A2:A" & n + 2)
                .Values = ws.Range("C2:C" & n + 2)
                .Format.Line.ForeColor.RGB = RGB(192, 0, 0)
                .Format.Line.DashStyle = msoLineDash
                .Format.Line.Weight = 1.5
            End With
        End If
        .HasTitle = True: .ChartTitle.Text = "Volume Over Time"
        .Axes(xlCategory).HasTitle = True: .Axes(xlCategory).AxisTitle.Text = "Date"
        .Axes(xlCategory).TickLabels.NumberFormat = "dd-mmm"
        .Axes(xlValue).HasTitle = True: .Axes(xlValue).AxisTitle.Text = "ML"
    End With

    ' EC chart with trigger threshold
    Set cht = ws.ChartObjects.Add(Schema.CHART_LEFT_POS, _
        Schema.CHART_TOP_START + Schema.CHART_HEIGHT_VOLUME + Schema.CHART_SPACING, _
        Schema.CHART_WIDTH, Schema.CHART_HEIGHT_METRIC)
    With cht.Chart
        .ChartType = xlLine
        ' EC series
        .SetSourceData ws.Range("D1:D" & n + 2)
        .SeriesCollection(1).XValues = ws.Range("A2:A" & n + 2)
        .SeriesCollection(1).Name = "EC"
        ' Trigger threshold series
        If cfg.TriggerChem(1) > 0 Then
            With .SeriesCollection.NewSeries
                .Name = "Trigger"
                .XValues = ws.Range("A2:A" & n + 2)
                .Values = ws.Range("E2:E" & n + 2)
                .Format.Line.ForeColor.RGB = RGB(192, 0, 0)
                .Format.Line.DashStyle = msoLineDash
                .Format.Line.Weight = 1.5
            End With
        End If
        .HasTitle = True: .ChartTitle.Text = "EC Over Time"
        .Axes(xlCategory).HasTitle = True: .Axes(xlCategory).AxisTitle.Text = "Date"
        .Axes(xlCategory).TickLabels.NumberFormat = "dd-mmm"
        .Axes(xlValue).HasTitle = True: .Axes(xlValue).AxisTitle.Text = "EC"
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
