Option Explicit
' WQOC: Entry point for Water Quality Optimisation Calculator.
' Dependencies: Core, Data, Sim, History, Schema

Public Sub Run()
    Dim s As State, cfg As Config, r As Result, cm As XlCalculation
    On Error GoTo Cleanup

    cm = Application.Calculation
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual

    s = Data.LoadState()
    cfg = Data.LoadConfig()
    r = Sim.Run(s, cfg)
    Data.SaveResult r
    History.RecordRun cfg, r
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
    Dim volData() As Double, ecData() As Double, days() As Long

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(Schema.SHEET_CHART)
    On Error GoTo 0
    If ws Is Nothing Then Exit Sub

    n = UBound(r.Snaps)
    ReDim volData(0 To n): ReDim ecData(0 To n): ReDim days(0 To n)

    For i = 0 To n
        days(i) = i
        volData(i) = r.Snaps(i).Vol
        ecData(i) = r.Snaps(i).Chem(1)
    Next i

    ' Clear existing charts
    For Each cht In ws.ChartObjects: cht.Delete: Next cht
    ws.Cells.Clear

    ' Write data to sheet for chart source
    ws.Range("A1") = "Day": ws.Range("B1") = "Volume (ML)": ws.Range("C1") = "EC"
    For i = 0 To n
        ws.Cells(i + 2, 1) = days(i)
        ws.Cells(i + 2, 2) = volData(i)
        ws.Cells(i + 2, 3) = ecData(i)
    Next i

    ' Volume chart
    Set cht = ws.ChartObjects.Add(Schema.CHART_LEFT_POS, Schema.CHART_TOP_START, _
                                   Schema.CHART_WIDTH, Schema.CHART_HEIGHT_VOLUME)
    With cht.Chart
        .ChartType = xlLine
        .SetSourceData ws.Range("A1:B" & n + 2)
        .HasTitle = True: .ChartTitle.Text = "Volume Over Time"
        .Axes(xlCategory).HasTitle = True: .Axes(xlCategory).AxisTitle.Text = "Day"
        .Axes(xlValue).HasTitle = True: .Axes(xlValue).AxisTitle.Text = "ML"
        If r.TriggerDay <> Core.NO_TRIGGER Then AddTriggerLine cht.Chart, r.TriggerDay
    End With

    ' EC chart
    Set cht = ws.ChartObjects.Add(Schema.CHART_LEFT_POS, _
        Schema.CHART_TOP_START + Schema.CHART_HEIGHT_VOLUME + Schema.CHART_SPACING, _
        Schema.CHART_WIDTH, Schema.CHART_HEIGHT_METRIC)
    With cht.Chart
        .ChartType = xlLine
        .SetSourceData Union(ws.Range("A1").Resize(n + 2, 1), ws.Range("C1").Resize(n + 2, 1))
        .SeriesCollection(1).Name = "EC"
        .HasTitle = True: .ChartTitle.Text = "EC Over Time"
        .Axes(xlCategory).HasTitle = True: .Axes(xlCategory).AxisTitle.Text = "Day"
        .Axes(xlValue).HasTitle = True: .Axes(xlValue).AxisTitle.Text = "EC"
        If r.TriggerDay <> Core.NO_TRIGGER Then AddTriggerLine cht.Chart, r.TriggerDay
    End With
End Sub

Private Sub AddTriggerLine(ByRef cht As Chart, ByVal trigDay As Long)
    ' Add vertical line at trigger day (approximation using a series)
    Dim s As Series
    On Error Resume Next
    Set s = cht.SeriesCollection.NewSeries
    If Not s Is Nothing Then
        s.Name = "Trigger"
        s.XValues = Array(trigDay, trigDay)
        s.Values = Array(cht.Axes(xlValue).MinimumScale, cht.Axes(xlValue).MaximumScale)
        s.ChartType = xlLine
        s.Format.Line.ForeColor.RGB = RGB(192, 0, 0)
        s.Format.Line.Weight = 2
    End If
    On Error GoTo 0
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
