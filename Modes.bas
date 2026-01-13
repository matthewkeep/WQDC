Option Explicit
' Modes: Simulation step functions.
' Dependencies: Types

Public Function Step(ByRef s As State, ByRef cfg As Config) As State
    Select Case UCase$(cfg.Mode)
        Case "SIMPLE": Step = StepSimple(s, cfg)
        Case "TWOBUCKET": Step = StepTwoBucket(s, cfg)
        Case Else: Step = StepSimple(s, cfg)
    End Select
End Function

' ==== Simple Mode ============================================================

Public Function StepSimple(ByRef s As State, ByRef cfg As Config) As State
    Dim n As State, i As Long, pVol As Double, mOut As Double, mIn As Double

    n = _Types.CopyState(s)
    pVol = s.Vol

    ' Volume: in + rain - out
    n.Vol = pVol + cfg.Inflow + cfg.RainVol - cfg.Outflow
    If n.Vol < 0 Then n.Vol = 0

    ' Mass balance per metric
    For i = 1 To _Types.METRIC_COUNT
        If pVol > _Types.EPS Then mOut = cfg.Outflow * s.Chem(i) Else mOut = 0
        mIn = cfg.Inflow * cfg.InflowChem(i)
        If n.Vol > _Types.EPS Then
            n.Chem(i) = (pVol * s.Chem(i) - mOut + mIn) / n.Vol
        Else
            n.Chem(i) = 0
        End If
    Next i

    StepSimple = n
End Function

' ==== TwoBucket Mode =========================================================

Public Function StepTwoBucket(ByRef s As State, ByRef cfg As Config) As State
    Dim n As State, i As Long
    Dim pVol As Double, alpha As Double, sf As Double
    Dim visMass As Double, hidMass As Double, mixUp As Double, mixDn As Double

    n = _Types.CopyState(s)
    pVol = s.Vol
    alpha = IIf(cfg.Tau > _Types.EPS, 1 - Exp(-1 / cfg.Tau), 0.1)
    sf = IIf(cfg.SurfaceFrac > 0, cfg.SurfaceFrac, 0.8)

    ' Volume: surface layer gets inflow/rain, loses outflow
    n.Vol = pVol + cfg.Inflow * sf + cfg.RainVol - cfg.Outflow
    If n.Vol < 0 Then n.Vol = 0

    ' Chemistry: mix between visible and hidden layers
    For i = 1 To _Types.METRIC_COUNT
        visMass = s.Vol * s.Chem(i)
        hidMass = s.Hidden(i)

        ' Exchange: alpha fraction mixes each way
        mixUp = alpha * hidMass
        mixDn = alpha * visMass

        visMass = visMass - mixDn + mixUp
        hidMass = hidMass - mixUp + mixDn

        ' Inflow adds to visible
        visMass = visMass + cfg.Inflow * cfg.InflowChem(i) * sf

        ' Outflow removes from visible
        If pVol > _Types.EPS Then
            visMass = visMass - cfg.Outflow * (visMass / pVol)
        End If

        ' Update concentrations
        If n.Vol > _Types.EPS Then n.Chem(i) = visMass / n.Vol Else n.Chem(i) = 0
        n.Hidden(i) = hidMass
    Next i

    StepTwoBucket = n
End Function
