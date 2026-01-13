Option Explicit
' WQOC: Entry point for Water Quality Optimisation Calculator.
' Dependencies: Core, Data, Sim, History, SimLog, Schema

Public Sub Run()
    Dim s As State, cfg As Config, r As Result, cm As XlCalculation
    Dim runId As String, latestDate As Date
    On Error GoTo Cleanup

    cm = Application.Calculation
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual

    s = Data.LoadState()
    cfg = Data.LoadConfig()

    ' Safety check: warn if running from a date earlier than existing log data
    latestDate = SimLog.GetLatestLogDate()
    If latestDate > 0 And cfg.StartDate < latestDate Then
        Application.ScreenUpdating = True
        If MsgBox("Start date (" & Format$(cfg.StartDate, "dd-mmm") & ") is before existing log data (" & _
                  Format$(latestDate, "dd-mmm") & "). This will overwrite future forecasts." & vbNewLine & vbNewLine & _
                  "Continue?", vbYesNo + vbQuestion, "WQOC") = vbNo Then
            Application.Calculation = cm
            Application.EnableEvents = True
            Exit Sub
        End If
        Application.ScreenUpdating = False
    End If

    r = Sim.Run(s, cfg)
    Data.SaveResult r
    runId = Format$(Now, "yyyymmdd_hhmmss")
    History.RecordRun cfg, r, runId
    SimLog.WriteLog r, cfg, runId
    GenerateCharts cfg

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

Private Sub GenerateCharts(ByRef cfg As Config)
    ' Draws charts from tblLogDaily (Log sheet) - shows cumulative run history
    Dim wsChart As Worksheet, wsLog As Worksheet, tbl As ListObject
    Dim cht As ChartObject, n As Long, i As Long
    Dim dateCol As Long, volCol As Long, ecCol As Long
    Dim trigArr() As Double

    On Error Resume Next
    Set wsChart = ThisWorkbook.Worksheets(Schema.SHEET_CHART)
    Set wsLog = ThisWorkbook.Worksheets(Schema.SHEET_LOG)
    On Error GoTo 0
    If wsChart Is Nothing Or wsLog Is Nothing Then Exit Sub

    Set tbl = wsLog.ListObjects(Schema.TABLE_LOG_DAILY)
    If tbl Is Nothing Then Exit Sub
    If tbl.DataBodyRange Is Nothing Then Exit Sub

    n = tbl.ListRows.Count
    If n = 0 Then Exit Sub

    ' Column indices in tblLogDaily: 1=RunId, 2=Date, 3=Day, 4=Volume, 5+=Chemistry
    dateCol = 2
    volCol = 4
    ecCol = 5  ' First chemistry metric (EC)

    ' Clear existing charts only (keep log data)
    For Each cht In wsChart.ChartObjects: cht.Delete: Next cht

    ' Volume chart - draws directly from Log table
    Set cht = wsChart.ChartObjects.Add(Schema.CHART_LEFT_POS, Schema.CHART_TOP_START, _
                                       Schema.CHART_WIDTH, Schema.CHART_HEIGHT_VOLUME)
    With cht.Chart
        .ChartType = xlLine
        ' Volume series from log
        With .SeriesCollection.NewSeries
            .Name = "Volume"
            .XValues = tbl.DataBodyRange.Columns(dateCol)
            .Values = tbl.DataBodyRange.Columns(volCol)
        End With
        ' Trigger threshold (horizontal line)
        If cfg.TriggerVol > 0 Then
            ReDim trigArr(1 To n)
            For i = 1 To n: trigArr(i) = cfg.TriggerVol: Next i
            With .SeriesCollection.NewSeries
                .Name = "Trigger"
                .XValues = tbl.DataBodyRange.Columns(dateCol)
                .Values = trigArr
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

    ' EC chart - draws directly from Log table
    Set cht = wsChart.ChartObjects.Add(Schema.CHART_LEFT_POS, _
        Schema.CHART_TOP_START + Schema.CHART_HEIGHT_VOLUME + Schema.CHART_SPACING, _
        Schema.CHART_WIDTH, Schema.CHART_HEIGHT_METRIC)
    With cht.Chart
        .ChartType = xlLine
        ' EC series from log
        With .SeriesCollection.NewSeries
            .Name = "EC"
            .XValues = tbl.DataBodyRange.Columns(dateCol)
            .Values = tbl.DataBodyRange.Columns(ecCol)
        End With
        ' Trigger threshold
        If cfg.TriggerChem(1) > 0 Then
            ReDim trigArr(1 To n)
            For i = 1 To n: trigArr(i) = cfg.TriggerChem(1): Next i
            With .SeriesCollection.NewSeries
                .Name = "Trigger"
                .XValues = tbl.DataBodyRange.Columns(dateCol)
                .Values = trigArr
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
