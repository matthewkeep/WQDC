Attribute VB_Name = "Scenarios"
Option Explicit
' Scenarios: Regression tests for simulation math.
' Dependencies: Types, Sim, Modes

Private Type Scen
    Nm As String
    Vol As Double
    Chem As Double
    Mode As String
    Days As Long
    Qin As Double
    Qout As Double
    ChemIn As Double
    TrigVol As Double
    TrigChem As Double
    Tau As Double
    HidMass As Double
    ExpDay As Long      ' -1 = no trigger
    ExpMetric As String
End Type

Public Sub RunAll()
    Dim sc() As Scen
    Dim i As Long, p As Long, f As Long

    sc = BuildScenarios()
    Debug.Print "": Debug.Print "=== Scenarios ==="

    For i = LBound(sc) To UBound(sc)
        If RunOne(sc(i)) Then p = p + 1 Else f = f + 1
    Next i

    Debug.Print "": Debug.Print "Results: " & p & " pass, " & f & " fail"
End Sub

Private Function BuildScenarios() As Scen()
    Dim s(0 To 5) As Scen

    ' 1: Simple volume trigger
    s(0).Nm = "Vol trigger": s(0).Vol = 90: s(0).Mode = "Simple"
    s(0).Days = 20: s(0).Qin = 2: s(0).TrigVol = 100
    s(0).ExpDay = 5: s(0).ExpMetric = "Volume"

    ' 2: No trigger - stable
    s(1).Nm = "No trigger": s(1).Vol = 100: s(1).Mode = "Simple"
    s(1).Days = 30: s(1).Qin = 1: s(1).Qout = 1: s(1).TrigVol = 200
    s(1).ExpDay = -1

    ' 3: Chemistry trigger
    s(2).Nm = "Chem trigger": s(2).Vol = 100: s(2).Chem = 80: s(2).Mode = "Simple"
    s(2).Days = 20: s(2).Qin = 10: s(2).Qout = 10: s(2).ChemIn = 200: s(2).TrigChem = 100
    s(2).ExpDay = 2: s(2).ExpMetric = "EC"

    ' 4: Volume decreasing
    s(3).Nm = "Vol decreasing": s(3).Vol = 150: s(3).Mode = "Simple"
    s(3).Days = 30: s(3).Qin = 1: s(3).Qout = 3: s(3).TrigVol = 200
    s(3).ExpDay = -1

    ' 5: TwoBucket mixing
    s(4).Nm = "TwoBucket mix": s(4).Vol = 100: s(4).Chem = 50: s(4).HidMass = 20000
    s(4).Mode = "TwoBucket": s(4).Days = 30: s(4).Tau = 5: s(4).TrigChem = 100
    s(4).ExpDay = 8: s(4).ExpMetric = "EC"

    ' 6: Fast fill
    s(5).Nm = "Fast fill": s(5).Vol = 99: s(5).Mode = "Simple"
    s(5).Days = 10: s(5).Qin = 5: s(5).TrigVol = 100
    s(5).ExpDay = 1: s(5).ExpMetric = "Volume"

    BuildScenarios = s
End Function

Private Function RunOne(ByRef sc As Scen) As Boolean
    Dim st As State, cfg As Config, r As Result
    Dim ok As Boolean

    st.Vol = sc.Vol: st.Chem(1) = sc.Chem: st.Hidden(1) = sc.HidMass: st.HidVol = 50
    cfg.Mode = sc.Mode: cfg.Days = sc.Days
    cfg.Inflow = sc.Qin: cfg.Outflow = sc.Qout: cfg.InflowChem(1) = sc.ChemIn
    cfg.TriggerVol = sc.TrigVol: cfg.TriggerChem(1) = sc.TrigChem
    cfg.Tau = IIf(sc.Tau > 0, sc.Tau, 7): cfg.SurfaceFrac = 0.8

    r = Sim.Run(st, cfg)

    ' Check result (TwoBucket gets +/-2 day tolerance)
    If sc.Mode = "TwoBucket" Then
        ok = (Abs(r.TriggerDay - sc.ExpDay) <= 2)
    Else
        ok = (r.TriggerDay = sc.ExpDay)
    End If
    If sc.ExpDay >= 0 Then ok = ok And (r.TriggerMetric = sc.ExpMetric)

    If ok Then
        Debug.Print "PASS: " & sc.Nm
    Else
        Debug.Print "FAIL: " & sc.Nm & " (exp " & sc.ExpDay & ", got " & r.TriggerDay & ")"
    End If
    RunOne = ok
End Function
