Attribute VB_Name = "Tests"
Option Explicit
' Tests: Smoke tests for the new architecture.
' Purpose: Verify Types, Modes, Sim work correctly without worksheet I/O.
' Dependencies: Types, Modes, Sim

' ==== Smoke Test Suite =======================================================
' Run in Immediate Window: Tests.RunSmokeSuite

Public Sub RunSmokeSuite()
    Dim passed As Long
    Dim failed As Long

    Debug.Print "=== WQDC Smoke Tests ==="
    Debug.Print ""

    ' Type tests
    RunTest "Types: metric count is 7", "TestMetricCount", passed, failed
    RunTest "Types: metric names exist", "TestMetricNames", passed, failed
    RunTest "Types: CopyState works", "TestCopyState", passed, failed

    ' Mode tests
    RunTest "Modes: simple volume balance", "TestSimpleVolume", passed, failed
    RunTest "Modes: simple mass balance", "TestSimpleMass", passed, failed
    RunTest "Modes: two-bucket mixing", "TestTwoBucketMixing", passed, failed

    ' Simulation tests
    RunTest "Sim: volume trigger detection", "TestVolumeTrigger", passed, failed
    RunTest "Sim: chemistry trigger detection", "TestChemTrigger", passed, failed
    RunTest "Sim: full run produces snapshots", "TestFullRun", passed, failed
    RunTest "Sim: no trigger case", "TestNoTrigger", passed, failed

    Debug.Print ""
    Debug.Print "=== Results: " & passed & " passed, " & failed & " failed ==="
End Sub

Private Sub RunTest(ByVal name As String, ByVal testFunc As String, ByRef p As Long, ByRef f As Long)
    Dim result As Boolean
    On Error Resume Next
    result = Application.Run(testFunc)
    If Err.Number <> 0 Then
        Debug.Print "FAIL: " & name & " (" & Err.Description & ")"
        f = f + 1
    ElseIf Not result Then
        Debug.Print "FAIL: " & name
        f = f + 1
    Else
        Debug.Print "PASS: " & name
        p = p + 1
    End If
    On Error GoTo 0
End Sub

' ==== Type Tests =============================================================

Private Function TestMetricCount() As Boolean
    TestMetricCount = (Types.METRIC_COUNT = 7)
End Function

Private Function TestMetricNames() As Boolean
    ' Verify first and last metric names
    TestMetricNames = (Types.MetricName(1) = "EC") And (Types.MetricName(7) = "TAN")
End Function

Private Function TestCopyState() As Boolean
    Dim s As State, c As State

    s.Vol = 100
    s.Chem(1) = 200
    s.Hidden(1) = 5000

    c = Types.CopyState(s)

    ' Modify original, verify copy is independent
    s.Vol = 50

    TestCopyState = (c.Vol = 100) And (c.Chem(1) = 200) And (c.Hidden(1) = 5000)
End Function

' ==== Mode Tests =============================================================

Private Function TestSimpleVolume() As Boolean
    ' Test: volume balance in simple mode
    ' 100 ML + 2 inflow + 0.5 rain - 1 outflow = 101.5 ML
    Dim s As State, cfg As Config, n As State

    s.Vol = 100
    cfg.Mode = "Simple"
    cfg.Inflow = 2
    cfg.Outflow = 1
    cfg.RainVol = 0.5

    n = Modes.StepSimple(s, cfg)

    TestSimpleVolume = (Abs(n.Vol - 101.5) < 0.01)
End Function

Private Function TestSimpleMass() As Boolean
    ' Test: mass balance in simple mode
    ' Start: 100 ML at 100 uS/cm = 10000 mass
    ' Add: 2 ML at 500 uS/cm = 1000 mass
    ' Remove: 1 ML at current conc (outflow removes at mixed conc)
    ' Net volume: 100 + 2 - 1 = 101 ML
    Dim s As State, cfg As Config, n As State

    s.Vol = 100
    s.Chem(1) = 100  ' EC = 100 uS/cm
    cfg.Mode = "Simple"
    cfg.Inflow = 2
    cfg.Outflow = 1
    cfg.InflowChem(1) = 500  ' Inflow EC = 500 uS/cm

    n = Modes.StepSimple(s, cfg)

    ' Expected: mass increases from inflow, diluted by volume change
    ' Initial mass: 100 * 100 = 10000
    ' Mass out (at initial conc): 1 * 100 = 100
    ' Mass in: 2 * 500 = 1000
    ' Final mass: 10000 - 100 + 1000 = 10900
    ' Final vol: 101
    ' Final conc: 10900 / 101 = 107.9
    TestSimpleMass = (Abs(n.Chem(1) - 107.9) < 0.5)
End Function

Private Function TestTwoBucketMixing() As Boolean
    ' Test: two-bucket mode mixes hidden mass into visible layer
    Dim s As State, cfg As Config, n As State

    s.Vol = 100
    s.Chem(1) = 100       ' Visible EC = 100
    s.Hidden(1) = 10000   ' Hidden mass = 10000 (will mix up)
    s.HidVol = 50         ' Hidden volume
    cfg.Mode = "TwoBucket"
    cfg.Tau = 7           ' 7-day mixing time
    cfg.SurfaceFrac = 0.8
    cfg.Inflow = 0
    cfg.Outflow = 0

    n = Modes.StepTwoBucket(s, cfg)

    ' After one day, some hidden mass should move to visible layer
    ' Visible concentration should increase
    TestTwoBucketMixing = (n.Chem(1) > s.Chem(1)) And (n.Hidden(1) < s.Hidden(1))
End Function

' ==== Simulation Tests =======================================================

Private Function TestVolumeTrigger() As Boolean
    ' Test: trigger fires when volume exceeds threshold
    Dim s As State, cfg As Config, r As Result

    s.Vol = 95
    cfg.Mode = "Simple"
    cfg.Days = 10
    cfg.Inflow = 2
    cfg.Outflow = 0
    cfg.TriggerVol = 100

    r = Sim.Run(s, cfg)

    ' Should trigger around day 3 (95 + 2*3 = 101 >= 100)
    TestVolumeTrigger = (r.TriggerDay = 3) And (r.TriggerMetric = "Volume")
End Function

Private Function TestChemTrigger() As Boolean
    ' Test: trigger fires when chemistry exceeds threshold
    Dim s As State, cfg As Config, r As Result

    s.Vol = 100
    s.Chem(1) = 90  ' EC starting at 90
    cfg.Mode = "Simple"
    cfg.Days = 10
    cfg.Inflow = 1
    cfg.Outflow = 1  ' Net neutral volume
    cfg.InflowChem(1) = 200  ' High EC inflow
    cfg.TriggerChem(1) = 100  ' EC trigger at 100

    r = Sim.Run(s, cfg)

    ' Should trigger when EC reaches 100
    TestChemTrigger = (r.TriggerDay > 0) And (r.TriggerMetric = "EC")
End Function

Private Function TestFullRun() As Boolean
    ' Test: full simulation produces expected snapshots
    Dim s As State, cfg As Config, r As Result

    s.Vol = 100
    cfg.Mode = "Simple"
    cfg.Days = 5
    cfg.Inflow = 1
    cfg.Outflow = 1

    r = Sim.Run(s, cfg)

    ' Should have 6 snapshots (day 0 through 5), volume stable at 100
    TestFullRun = (UBound(r.Snaps) = 5) And (Abs(r.Snaps(5).Vol - 100) < 0.01)
End Function

Private Function TestNoTrigger() As Boolean
    ' Test: no trigger when thresholds not exceeded
    Dim s As State, cfg As Config, r As Result

    s.Vol = 50
    cfg.Mode = "Simple"
    cfg.Days = 10
    cfg.Inflow = 1
    cfg.Outflow = 1
    cfg.TriggerVol = 200  ' Very high, won't be reached

    r = Sim.Run(s, cfg)

    ' Should not trigger
    TestNoTrigger = (r.TriggerDay = Types.NO_TRIGGER)
End Function

' ==== Quick Manual Tests =====================================================

' Run these from Immediate Window for quick verification:
'   Tests.TestWQDC
'   Tests.TestTwoBucket

Public Sub TestWQDC()
    ' Test the main simulation flow
    WQDC.TestCore
End Sub

Public Sub TestTwoBucket()
    ' Test two-bucket mode
    WQDC.TestTwoBucket
End Sub
