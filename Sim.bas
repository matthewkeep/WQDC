Option Explicit
' Sim: Core simulation loop.
' Dependencies: Types, Modes

Public Function Run(ByRef init As State, ByRef cfg As Config) As Result
    Dim r As Result, s As State, d As Long

    s = Types.CopyState(init)
    r.TriggerDay = Types.NO_TRIGGER
    ReDim r.Snaps(0 To cfg.Days)
    r.Snaps(0) = s

    For d = 1 To cfg.Days
        s = Modes.Step(s, cfg)
        r.Snaps(d) = s
        If r.TriggerDay = Types.NO_TRIGGER Then
            If ChkTriggers(s, cfg, r.TriggerMetric) Then
                r.TriggerDay = d
                r.TriggerDate = cfg.StartDate + d
            End If
        End If
    Next d

    r.FinalState = s
    Run = r
End Function

Private Function ChkTriggers(ByRef s As State, ByRef cfg As Config, ByRef metric As String) As Boolean
    Dim i As Long

    ' Volume trigger
    If cfg.TriggerVol > 0 And s.Vol >= cfg.TriggerVol Then
        metric = "Volume": ChkTriggers = True: Exit Function
    End If

    ' Chemistry triggers
    For i = 1 To Types.METRIC_COUNT
        If cfg.TriggerChem(i) > 0 And s.Chem(i) >= cfg.TriggerChem(i) Then
            metric = Types.MetricName(i): ChkTriggers = True: Exit Function
        End If
    Next i
End Function
