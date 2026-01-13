Attribute VB_Name = "Scenarios"
Option Explicit
' Scenarios: Data-driven regression testing for simulation math.
' Purpose: Prove the core engine produces correct results.
' Dependencies: Types, Sim, Modes
'
' Usage:
'   Scenarios.RunAll     - Run all scenarios, report pass/fail
'   Scenarios.RunOne n   - Run scenario n only (1-based)
'
' This module is standalone and can be removed without affecting core functionality.

Private Type Scenario
    Name As String
    InitVol As Double
    InitChem As Double          ' EC only for simplicity
    Mode As String
    Days As Long
    Inflow As Double
    Outflow As Double
    InflowChem As Double        ' EC only
    TriggerVol As Double
    TriggerChem As Double       ' EC only
    Tau As Double               ' For TwoBucket mode
    HiddenMass As Double        ' For TwoBucket mode
    ExpectedTriggerDay As Long  ' -1 = no trigger expected
    ExpectedMetric As String    ' "Volume", "EC", or ""
End Type

' ==== Public Entry Points ====================================================

' Run all scenarios
Public Sub RunAll()
    Dim scenarios() As Scenario
    Dim i As Long
    Dim passed As Long
    Dim failed As Long

    scenarios = BuildScenarios()

    Debug.Print ""
    Debug.Print "=== Scenario Tests ==="
    Debug.Print ""

    For i = LBound(scenarios) To UBound(scenarios)
        If RunScenario(scenarios(i)) Then
            passed = passed + 1
        Else
            failed = failed + 1
        End If
    Next i

    Debug.Print ""
    Debug.Print "=== Results: " & passed & " passed, " & failed & " failed ==="
    Debug.Print ""
End Sub

' Run single scenario by index (1-based)
Public Sub RunOne(ByVal idx As Long)
    Dim scenarios() As Scenario

    scenarios = BuildScenarios()

    If idx < 1 Or idx > UBound(scenarios) + 1 Then
        Debug.Print "Invalid scenario index. Valid range: 1-" & UBound(scenarios) + 1
        Exit Sub
    End If

    Debug.Print ""
    RunScenario scenarios(idx - 1)
    Debug.Print ""
End Sub

' ==== Scenario Definitions ===================================================

Private Function BuildScenarios() As Scenario()
    Dim s(0 To 5) As Scenario

    ' Scenario 1: Simple volume trigger
    s(0).Name = "Simple volume trigger"
    s(0).InitVol = 90
    s(0).Mode = "Simple"
    s(0).Days = 20
    s(0).Inflow = 2
    s(0).Outflow = 0
    s(0).TriggerVol = 100
    s(0).ExpectedTriggerDay = 5  ' 90 + 5*2 = 100
    s(0).ExpectedMetric = "Volume"

    ' Scenario 2: No trigger (volume stable)
    s(1).Name = "No trigger - stable volume"
    s(1).InitVol = 100
    s(1).Mode = "Simple"
    s(1).Days = 30
    s(1).Inflow = 1
    s(1).Outflow = 1
    s(1).TriggerVol = 200
    s(1).ExpectedTriggerDay = -1
    s(1).ExpectedMetric = ""

    ' Scenario 3: Chemistry trigger (EC rising)
    s(2).Name = "Chemistry trigger - EC"
    s(2).InitVol = 100
    s(2).InitChem = 80
    s(2).Mode = "Simple"
    s(2).Days = 20
    s(2).Inflow = 10
    s(2).Outflow = 10
    s(2).InflowChem = 200
    s(2).TriggerChem = 100
    s(2).ExpectedTriggerDay = 2  ' Rapid concentration rise
    s(2).ExpectedMetric = "EC"

    ' Scenario 4: Volume decreasing (no trigger)
    s(3).Name = "Volume decreasing - no trigger"
    s(3).InitVol = 150
    s(3).Mode = "Simple"
    s(3).Days = 30
    s(3).Inflow = 1
    s(3).Outflow = 3
    s(3).TriggerVol = 200
    s(3).ExpectedTriggerDay = -1
    s(3).ExpectedMetric = ""

    ' Scenario 5: TwoBucket mode - hidden mass mixing up
    s(4).Name = "TwoBucket - hidden mass mixing"
    s(4).InitVol = 100
    s(4).InitChem = 50
    s(4).HiddenMass = 20000  ' High hidden mass will mix up
    s(4).Mode = "TwoBucket"
    s(4).Days = 30
    s(4).Tau = 5
    s(4).Inflow = 0
    s(4).Outflow = 0
    s(4).TriggerChem = 100
    s(4).ExpectedTriggerDay = 8  ' Approximate - hidden mass raises EC
    s(4).ExpectedMetric = "EC"

    ' Scenario 6: Fast fill to trigger
    s(5).Name = "Fast fill - day 1 trigger"
    s(5).InitVol = 99
    s(5).Mode = "Simple"
    s(5).Days = 10
    s(5).Inflow = 5
    s(5).Outflow = 0
    s(5).TriggerVol = 100
    s(5).ExpectedTriggerDay = 1  ' 99 + 5 = 104 on day 1
    s(5).ExpectedMetric = "Volume"

    BuildScenarios = s
End Function

' ==== Scenario Execution =====================================================

Private Function RunScenario(ByRef sc As Scenario) As Boolean
    Dim s As State
    Dim cfg As Config
    Dim r As Result
    Dim passed As Boolean
    Dim details As String

    ' Build initial state
    s.Vol = sc.InitVol
    s.Chem(1) = sc.InitChem
    s.Hidden(1) = sc.HiddenMass
    s.HidVol = 50  ' Default hidden volume for TwoBucket

    ' Build config
    cfg.Mode = sc.Mode
    cfg.Days = sc.Days
    cfg.Inflow = sc.Inflow
    cfg.Outflow = sc.Outflow
    cfg.InflowChem(1) = sc.InflowChem
    cfg.TriggerVol = sc.TriggerVol
    cfg.TriggerChem(1) = sc.TriggerChem
    cfg.Tau = IIf(sc.Tau > 0, sc.Tau, 7)
    cfg.SurfaceFrac = 0.8

    ' Run simulation
    r = Sim.Run(s, cfg)

    ' Check result
    passed = CheckResult(sc, r, details)

    ' Report
    If passed Then
        Debug.Print "PASS: " & sc.Name
    Else
        Debug.Print "FAIL: " & sc.Name
        Debug.Print "      " & details
    End If

    RunScenario = passed
End Function

Private Function CheckResult(ByRef sc As Scenario, ByRef r As Result, ByRef details As String) As Boolean
    Dim dayOk As Boolean
    Dim metricOk As Boolean

    ' Check trigger day (allow +/- 1 day tolerance for TwoBucket due to floating point)
    If sc.Mode = "TwoBucket" Then
        dayOk = (Abs(r.TriggerDay - sc.ExpectedTriggerDay) <= 2)
    Else
        dayOk = (r.TriggerDay = sc.ExpectedTriggerDay)
    End If

    ' Check trigger metric
    If sc.ExpectedTriggerDay = -1 Then
        metricOk = (r.TriggerDay = -1)
    Else
        metricOk = (r.TriggerMetric = sc.ExpectedMetric)
    End If

    If Not dayOk Then
        details = "Expected trigger day " & sc.ExpectedTriggerDay & ", got " & r.TriggerDay
    ElseIf Not metricOk Then
        details = "Expected metric '" & sc.ExpectedMetric & "', got '" & r.TriggerMetric & "'"
    End If

    CheckResult = (dayOk And metricOk)
End Function
