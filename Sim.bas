Option Explicit
' Sim: Core simulation loop.
' Dependencies: Core, Modes, Telemetry, Schema

Public Function Run(ByRef init As State, ByRef cfg As Config) As Result
    Dim r As Result, s As State, d As Long
    Dim currentDate As Date

    s = Core.CopyState(init)
    r.TriggerDay = Core.NO_TRIGGER
    ReDim r.Snaps(0 To cfg.Days)
    r.Snaps(0) = s

    For d = 1 To cfg.Days
        currentDate = cfg.StartDate + d

        ' Apply per-day rainfall if enabled
        cfg.RainVol = GetRainForDay(currentDate, cfg.StartDate, cfg.RainfallMode)

        s = Modes.Step(s, cfg)
        r.Snaps(d) = s
        If r.TriggerDay = Core.NO_TRIGGER Then
            If ChkTriggers(s, cfg, r.TriggerMetric) Then
                r.TriggerDay = d
                r.TriggerDate = currentDate
            End If
        End If
    Next d

    r.FinalState = s
    Run = r
End Function

Private Function GetRainForDay(ByVal d As Date, ByVal startDate As Date, ByVal mode As String) As Double
    ' Returns rainfall volume for a given day based on mode
    ' Off: no rain, Hindcast: actual past rain, Hindcast+Forecast: extrapolate average
    Dim avgRain As Double, hindcastDays As Long

    If UCase$(mode) = UCase$(Schema.RAINFALL_OFF) Or Len(mode) = 0 Then
        GetRainForDay = 0
        Exit Function
    End If

    If d <= Date Then
        ' Hindcast period: use actual telemetry
        GetRainForDay = Telemetry.GetRainForDate(d)
    ElseIf UCase$(mode) = UCase$(Schema.RAINFALL_FULL) Then
        ' Forecast with typical: average daily rain from hindcast period
        hindcastDays = Date - startDate
        If hindcastDays > 0 Then
            avgRain = Telemetry.GetTotalRain(startDate, Date) / hindcastDays
        Else
            avgRain = 0
        End If
        GetRainForDay = avgRain
    Else
        ' Hindcast only: no rain for forecast days
        GetRainForDay = 0
    End If
End Function

Private Function ChkTriggers(ByRef s As State, ByRef cfg As Config, ByRef metric As String) As Boolean
    Dim i As Long

    ' Volume trigger
    If cfg.TriggerVol > 0 And s.Vol >= cfg.TriggerVol Then
        metric = "Volume": ChkTriggers = True: Exit Function
    End If

    ' Chemistry triggers
    For i = 1 To Core.METRIC_COUNT
        If cfg.TriggerChem(i) > 0 And s.Chem(i) >= cfg.TriggerChem(i) Then
            metric = Core.MetricName(i): ChkTriggers = True: Exit Function
        End If
    Next i
End Function
