Attribute VB_Name = "Sim"
Option Explicit
' Sim: Core simulation loop.
' Purpose: Run simulation, check triggers, collect snapshots. That's it.
' Dependencies: Types, Modes

' ==== Main Entry Point ========================================================

' Run simulation from initial state with given config
' Returns Result with trigger info and daily snapshots
Public Function Run(ByRef init As State, ByRef cfg As Config) As Result
    Dim r As Result
    Dim s As State
    Dim day As Long
    Dim triggered As Boolean

    ' Initialize
    s = Types.CopyState(init)
    r.TriggerDay = Types.NO_TRIGGER
    ReDim r.Snaps(0 To cfg.Days)
    r.Snaps(0) = s

    ' Simulate each day
    For day = 1 To cfg.Days
        ' Step using configured mode
        s = Modes.Step(s, cfg)

        ' Snapshot
        r.Snaps(day) = s

        ' Check triggers (only if not already triggered)
        If r.TriggerDay = Types.NO_TRIGGER Then
            triggered = CheckTriggers(s, cfg, r.TriggerMetric)
            If triggered Then
                r.TriggerDay = day
                r.TriggerDate = cfg.StartDate + day
            End If
        End If
    Next day

    r.FinalState = s
    Run = r
End Function

' ==== Trigger Detection =======================================================

' Check if state exceeds any trigger threshold
' Returns True if triggered, sets metricName to which one
Private Function CheckTriggers(ByRef s As State, ByRef cfg As Config, _
                                ByRef metricName As String) As Boolean
    Dim i As Long

    ' Volume trigger
    If cfg.TriggerVol > 0 Then
        If s.Vol >= cfg.TriggerVol Then
            metricName = "Volume"
            CheckTriggers = True
            Exit Function
        End If
    End If

    ' Chemistry triggers
    For i = 1 To Types.METRIC_COUNT
        If cfg.TriggerChem(i) > 0 Then
            If s.Chem(i) >= cfg.TriggerChem(i) Then
                metricName = Types.MetricName(i)
                CheckTriggers = True
                Exit Function
            End If
        End If
    Next i

    CheckTriggers = False
End Function

' ==== Convenience Functions ===================================================

' Quick run: create config, run, return trigger day
' Returns -1 if no trigger, or the day number if triggered
Public Function QuickRun(ByRef init As State, ByVal mode As String, _
                         ByVal days As Long, ByVal triggerVol As Double) As Long
    Dim cfg As Config
    Dim r As Result

    cfg.Mode = mode
    cfg.Days = days
    cfg.TriggerVol = triggerVol

    r = Run(init, cfg)
    QuickRun = r.TriggerDay
End Function
