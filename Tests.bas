Option Explicit
' Tests: Smoke tests for core modules.
' Dependencies: Types, Modes, Sim

Public Sub RunSmokeSuite()
    Dim p As Long, f As Long
    Debug.Print "": Debug.Print "=== WQOC Smoke Tests ==="

    Run "Types: count=7", "TstMetricCount", p, f
    Run "Types: names", "TstMetricNames", p, f
    Run "Types: copy", "TstCopyState", p, f
    Run "Modes: vol balance", "TstSimpleVol", p, f
    Run "Modes: mass balance", "TstSimpleMass", p, f
    Run "Modes: twobucket", "TstTwoBucket", p, f
    Run "Sim: vol trigger", "TstVolTrigger", p, f
    Run "Sim: chem trigger", "TstChemTrigger", p, f
    Run "Sim: snapshots", "TstFullRun", p, f
    Run "Sim: no trigger", "TstNoTrigger", p, f

    Debug.Print "": Debug.Print "Results: " & p & " pass, " & f & " fail"
End Sub

Private Sub Run(ByVal nm As String, ByVal fn As String, ByRef p As Long, ByRef f As Long)
    Dim ok As Boolean
    On Error Resume Next
    ok = Application.Run(fn)
    If Err.Number <> 0 Then
        Debug.Print "FAIL: " & nm & " (" & Err.Description & ")"
        f = f + 1
    ElseIf Not ok Then
        Debug.Print "FAIL: " & nm
        f = f + 1
    Else
        Debug.Print "PASS: " & nm
        p = p + 1
    End If
    On Error GoTo 0
End Sub

' ==== Type Tests =============================================================

Private Function TstMetricCount() As Boolean
    TstMetricCount = (AAATypes.METRIC_COUNT = 7)
End Function

Private Function TstMetricNames() As Boolean
    TstMetricNames = (AAATypes.MetricName(1) = "EC") And (AAATypes.MetricName(7) = "TAN")
End Function

Private Function TstCopyState() As Boolean
    Dim s As State, c As State
    s.Vol = 100: s.Chem(1) = 200: s.Hidden(1) = 5000
    c = AAATypes.CopyState(s)
    s.Vol = 50
    TstCopyState = (c.Vol = 100) And (c.Chem(1) = 200)
End Function

' ==== Mode Tests =============================================================

Private Function TstSimpleVol() As Boolean
    Dim s As State, cfg As Config, n As State
    s.Vol = 100
    cfg.Mode = "Simple": cfg.Inflow = 2: cfg.Outflow = 1: cfg.RainVol = 0.5
    n = Modes.StepSimple(s, cfg)
    TstSimpleVol = (Abs(n.Vol - 101.5) < 0.01)
End Function

Private Function TstSimpleMass() As Boolean
    Dim s As State, cfg As Config, n As State
    s.Vol = 100: s.Chem(1) = 100
    cfg.Mode = "Simple": cfg.Inflow = 2: cfg.Outflow = 1: cfg.InflowChem(1) = 500
    n = Modes.StepSimple(s, cfg)
    TstSimpleMass = (Abs(n.Chem(1) - 107.9) < 0.5)
End Function

Private Function TstTwoBucket() As Boolean
    Dim s As State, cfg As Config, n As State
    s.Vol = 100: s.Chem(1) = 100: s.Hidden(1) = 10000: s.HidVol = 50
    cfg.Mode = "TwoBucket": cfg.Tau = 7: cfg.SurfaceFrac = 0.8
    n = Modes.StepTwoBucket(s, cfg)
    TstTwoBucket = (n.Chem(1) > s.Chem(1)) And (n.Hidden(1) < s.Hidden(1))
End Function

' ==== Sim Tests ==============================================================

Private Function TstVolTrigger() As Boolean
    Dim s As State, cfg As Config, r As Result
    s.Vol = 95
    cfg.Mode = "Simple": cfg.Days = 10: cfg.Inflow = 2: cfg.TriggerVol = 100
    r = Sim.Run(s, cfg)
    TstVolTrigger = (r.TriggerDay = 3) And (r.TriggerMetric = "Volume")
End Function

Private Function TstChemTrigger() As Boolean
    Dim s As State, cfg As Config, r As Result
    s.Vol = 100: s.Chem(1) = 90
    cfg.Mode = "Simple": cfg.Days = 10: cfg.Inflow = 1: cfg.Outflow = 1
    cfg.InflowChem(1) = 200: cfg.TriggerChem(1) = 100
    r = Sim.Run(s, cfg)
    TstChemTrigger = (r.TriggerDay > 0) And (r.TriggerMetric = "EC")
End Function

Private Function TstFullRun() As Boolean
    Dim s As State, cfg As Config, r As Result
    s.Vol = 100
    cfg.Mode = "Simple": cfg.Days = 5: cfg.Inflow = 1: cfg.Outflow = 1
    r = Sim.Run(s, cfg)
    TstFullRun = (UBound(r.Snaps) = 5) And (Abs(r.Snaps(5).Vol - 100) < 0.01)
End Function

Private Function TstNoTrigger() As Boolean
    Dim s As State, cfg As Config, r As Result
    s.Vol = 50
    cfg.Mode = "Simple": cfg.Days = 10: cfg.Inflow = 1: cfg.Outflow = 1: cfg.TriggerVol = 200
    r = Sim.Run(s, cfg)
    TstNoTrigger = (r.TriggerDay = AAATypes.NO_TRIGGER)
End Function

' ==== Manual Tests ===========================================================

Public Sub TestWQOC(): WQOC.TestCore: End Sub
Public Sub TestTwoBucket(): WQOC.TestTwoBucket: End Sub
