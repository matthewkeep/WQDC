Option Explicit
' Modes: Simulation step functions.
' Dependencies: Core, Schema

' ==== Public Dispatcher =======================================================

Public Function Step(ByRef s As State, ByRef cfg As Config, ByVal rainVol As Double) As State
    Select Case UCase$(cfg.Mode)
        Case UCase$(Schema.MIXING_SIMPLE): Step = StepSimple(s, cfg, rainVol)
        Case UCase$(Schema.MIXING_TWOBUCKET): Step = StepTwoBucket(s, cfg, rainVol)
        Case Else: Step = StepSimple(s, cfg, rainVol)
    End Select
End Function

' ==== Simple Mode ============================================================

Public Function StepSimple(ByRef s As State, ByRef cfg As Config, ByVal rainVol As Double) As State
    Dim n As State, i As Long, pVol As Double, mOut As Double, mIn As Double

    On Error GoTo Fail

    n = Core.CopyState(s)
    pVol = s.Vol

    ' Volume: in + rain - out
    n.Vol = pVol + cfg.Inflow + rainVol - cfg.Outflow
    If n.Vol < 0 Then n.Vol = 0

    ' Mass balance per metric
    For i = 1 To Core.METRIC_COUNT
        If pVol > Core.EPS Then mOut = cfg.Outflow * s.Chem(i) Else mOut = 0
        mIn = cfg.Inflow * cfg.InflowChem(i)
        If n.Vol > Core.EPS Then
            n.Chem(i) = (pVol * s.Chem(i) - mOut + mIn) / n.Vol
        Else
            n.Chem(i) = 0
        End If
    Next i

    StepSimple = n
    Exit Function

Fail:
    Error.TraceErr "Modes.StepSimple"
    StepSimple = s  ' Return unchanged state on error
End Function

' ==== TwoBucket Mode =========================================================
' Two-layer stratified mixing model:
'   - Visible layer: surface water (sampled, released, volume-tracked)
'   - Hidden layer: deep water storing unmixed mass
'   - SurfaceFrac: fraction of inflow mass entering visible layer (rest to hidden)
'   - Tau: mixing time constant controlling exchange rate between layers
'
' Volume is fully conserved (all inflow enters system).
' Stratification only affects chemistry distribution, not total volume.

Public Function StepTwoBucket(ByRef s As State, ByRef cfg As Config, ByVal rainVol As Double) As State
    Dim n As State, i As Long
    Dim pVol As Double, newVol As Double, preOutVol As Double
    Dim alpha As Double, sf As Double
    Dim visMass As Double, hidMass As Double
    Dim inflowMass As Double, mixUp As Double, mixDn As Double

    On Error GoTo Fail

    n = Core.CopyState(s)
    pVol = s.Vol

    ' Mixing parameters
    alpha = IIf(cfg.Tau > Core.EPS, 1 - Exp(-1 / cfg.Tau), 0.1)
    sf = IIf(cfg.SurfaceFrac > 0, cfg.SurfaceFrac, 0.8)

    ' Volume: full water balance (all inflow enters, same as Simple)
    preOutVol = pVol + cfg.Inflow + rainVol
    newVol = preOutVol - cfg.Outflow
    If newVol < 0 Then newVol = 0
    n.Vol = newVol

    ' Chemistry: two-layer mass balance
    For i = 1 To Core.METRIC_COUNT
        visMass = pVol * s.Chem(i)
        hidMass = s.Hidden(i)
        inflowMass = cfg.Inflow * cfg.InflowChem(i)

        ' Step 1: Layer exchange (alpha fraction mixes each direction)
        mixUp = alpha * hidMass
        mixDn = alpha * visMass
        visMass = visMass - mixDn + mixUp
        hidMass = hidMass - mixUp + mixDn

        ' Step 2: Inflow mass splits between layers (volume fully enters visible)
        visMass = visMass + inflowMass * sf
        hidMass = hidMass + inflowMass * (1 - sf)

        ' Step 3: Rain adds volume but no mass (dilution effect in final calc)

        ' Step 4: Outflow removes mass at pre-outflow concentration
        If preOutVol > Core.EPS Then
            visMass = visMass - cfg.Outflow * (visMass / preOutVol)
        End If

        ' Update state
        If newVol > Core.EPS Then n.Chem(i) = visMass / newVol Else n.Chem(i) = 0
        n.Hidden(i) = hidMass
    Next i

    StepTwoBucket = n
    Exit Function

Fail:
    Error.TraceErr "Modes.StepTwoBucket"
    StepTwoBucket = s
End Function
