Attribute VB_Name = "Modes"
Option Explicit
' Modes: Pluggable simulation step functions.
' Purpose: Each mode is ONE function. Adding a new mode = adding one function.
' Dependencies: Types

' ==== Mode Dispatcher =========================================================

' Step one day using the configured mode
Public Function Step(ByRef s As State, ByRef cfg As Config) As State
    Select Case UCase$(cfg.Mode)
        Case "SIMPLE"
            Step = StepSimple(s, cfg)
        Case "TWOBUCKET"
            Step = StepTwoBucket(s, cfg)
        Case Else
            ' Default to simple
            Step = StepSimple(s, cfg)
    End Select
End Function

' ==== Mode: Simple ============================================================
' Well-mixed reservoir. All inflow mixes instantly.
' This is the baseline - easy to understand, easy to verify.

Public Function StepSimple(ByRef s As State, ByRef cfg As Config) As State
    Dim n As State
    Dim i As Long
    Dim prevVol As Double
    Dim massOut As Double
    Dim massIn As Double

    n = Types.CopyState(s)
    prevVol = s.Vol

    ' Volume balance: in + rain - out
    n.Vol = prevVol + cfg.Inflow + cfg.RainVol - cfg.Outflow
    If n.Vol < 0 Then n.Vol = 0

    ' Mass balance for each metric
    For i = 1 To Types.METRIC_COUNT
        ' Mass out = outflow * concentration
        If prevVol > Types.EPS Then
            massOut = cfg.Outflow * s.Chem(i)
        Else
            massOut = 0
        End If

        ' Mass in = inflow * inflow concentration
        massIn = cfg.Inflow * cfg.InflowChem(i)

        ' Current mass = volume * concentration
        ' New mass = current - out + in
        ' New concentration = new mass / new volume
        If n.Vol > Types.EPS Then
            n.Chem(i) = (prevVol * s.Chem(i) - massOut + massIn) / n.Vol
        Else
            n.Chem(i) = 0
        End If

        If n.Chem(i) < 0 Then n.Chem(i) = 0
    Next i

    StepSimple = n
End Function

' ==== Mode: TwoBucket =========================================================
' Stratified reservoir with visible (surface) and hidden (deep) layers.
' Inflow enters hidden bucket, mixes upward via alpha = 1 - exp(-1/tau).

Public Function StepTwoBucket(ByRef s As State, ByRef cfg As Config) As State
    Dim n As State
    Dim i As Long
    Dim alpha As Double
    Dim visVol As Double, hidVol As Double
    Dim outVis As Double, outHid As Double
    Dim visRain As Double, hidRain As Double
    Dim massOutVis As Double, massOutHid As Double
    Dim massIn As Double, delta As Double
    Dim visMass As Double, hidMass As Double

    n = Types.CopyState(s)

    ' Mixing coefficient
    If cfg.Tau > 0 Then
        alpha = 1 - Exp(-1 / cfg.Tau)
    Else
        alpha = 1 ' Instant mixing if tau=0
    End If

    ' Current volumes
    visVol = s.Vol
    hidVol = s.HidVol

    ' Outflow: visible first, then hidden
    outVis = Min(cfg.Outflow, visVol)
    outHid = Min(cfg.Outflow - outVis, hidVol)

    ' Rain split
    visRain = cfg.RainVol * cfg.SurfaceFrac
    hidRain = cfg.RainVol * (1 - cfg.SurfaceFrac)

    ' Volume update
    n.Vol = visVol - outVis + visRain
    n.HidVol = hidVol - outHid + hidRain + cfg.Inflow

    If n.Vol < 0 Then n.Vol = 0
    If n.HidVol < 0 Then n.HidVol = 0

    ' Mass balance for each metric
    For i = 1 To Types.METRIC_COUNT
        ' Current masses
        visMass = visVol * s.Chem(i)
        hidMass = s.Hidden(i)

        ' Mass out from visible
        If visVol > Types.EPS Then
            massOutVis = outVis * s.Chem(i)
        Else
            massOutVis = 0
        End If
        If massOutVis > visMass Then massOutVis = visMass

        ' Mass out from hidden
        If hidVol > Types.EPS And outHid > 0 Then
            massOutHid = outHid * (hidMass / hidVol)
        Else
            massOutHid = 0
        End If
        If massOutHid > hidMass Then massOutHid = hidMass

        ' Mass in (goes to hidden)
        massIn = cfg.Inflow * cfg.InflowChem(i)

        ' Update hidden mass (receives inflow, loses to mixing)
        hidMass = hidMass - massOutHid + massIn
        If hidMass < 0 Then hidMass = 0

        ' Mixing transfer: hidden -> visible
        delta = alpha * hidMass
        hidMass = hidMass - delta

        ' Update visible mass (receives rain contribution + mixing)
        visMass = visMass - massOutVis + delta
        If visMass < 0 Then visMass = 0

        ' Store results
        n.Hidden(i) = hidMass
        If n.Vol > Types.EPS Then
            n.Chem(i) = visMass / n.Vol
        Else
            n.Chem(i) = 0
        End If
    Next i

    StepTwoBucket = n
End Function

' ==== Add Your Mode Here ======================================================
' To add a new mode:
' 1. Create: Public Function StepMyMode(s As State, cfg As Config) As State
' 2. Add case to Step() dispatcher above
' 3. Done!

' Example template:
' Public Function StepExperimental(ByRef s As State, ByRef cfg As Config) As State
'     Dim n As State
'     n = Types.CopyState(s)
'     ' ... your physics here ...
'     StepExperimental = n
' End Function

' ==== Helper ==================================================================

Private Function Min(ByVal a As Double, ByVal b As Double) As Double
    If a < b Then Min = a Else Min = b
End Function
