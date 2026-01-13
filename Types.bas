Attribute VB_Name = "Types"
Option Explicit
' Types: Core type definitions for WQOC.
' Purpose: Minimal types - State, Config, Result. That's it.
' Dependencies: None

' ==== Constants ===============================================================

Public Const METRIC_COUNT As Long = 7       ' EC, F_U, F_Mn, SO4, Mg, Ca, TAN
Public Const NO_TRIGGER As Long = -1        ' Sentinel: no trigger occurred
Public Const EPS As Double = 0.000001       ' Epsilon for safe division

' ==== Core Types ==============================================================

' State: Where the reservoir is RIGHT NOW
' This is all you need to know to continue simulating
Public Type State
    Vol As Double                   ' Total volume (ML)
    Chem(1 To 7) As Double          ' Concentrations by metric
    Hidden(1 To 7) As Double        ' Hidden mass (for two-bucket mode)
    HidVol As Double                ' Hidden volume (for two-bucket mode)
End Type

' Config: How to run the simulation
' Set once, never changes during a run
Public Type Config
    ' Mode selection
    Mode As String                  ' "Simple", "TwoBucket", etc.

    ' Time
    Days As Long                    ' Forecast days
    StartDate As Date               ' Sample date

    ' Physics
    Tau As Double                   ' Mixing time constant
    Inflow As Double                ' Daily inflow (ML/d)
    Outflow As Double               ' Daily outflow (ML/d)
    RainVol As Double               ' Daily rain volume (ML/d)
    SurfaceFrac As Double           ' Surface fraction for rain split

    ' Inflow chemistry (mass per day = Inflow * Conc)
    InflowChem(1 To 7) As Double    ' Inflow concentrations

    ' Triggers
    TriggerVol As Double            ' Volume trigger (ML)
    TriggerChem(1 To 7) As Double   ' Chemistry triggers
End Type

' Result: What happened
Public Type Result
    TriggerDay As Long              ' Day trigger occurred (-1 if none)
    TriggerDate As Date             ' Date trigger occurred
    TriggerMetric As String         ' Which metric triggered ("Volume" or metric name)
    Snaps() As State                ' Daily snapshots for output/charts
    FinalState As State             ' End state (for chaining runs)
End Type

' ==== Metric Names ============================================================

Private mMetricNames As Variant

Public Function MetricName(ByVal idx As Long) As String
    If IsEmpty(mMetricNames) Then
        mMetricNames = Array("EC", "F_U", "F_Mn", "SO4", "Mg", "Ca", "TAN")
    End If
    If idx >= 1 And idx <= METRIC_COUNT Then
        MetricName = mMetricNames(idx - 1)
    End If
End Function

Public Function MetricNames() As Variant
    If IsEmpty(mMetricNames) Then
        mMetricNames = Array("EC", "F_U", "F_Mn", "SO4", "Mg", "Ca", "TAN")
    End If
    MetricNames = mMetricNames
End Function

' ==== Helper Functions ========================================================

' Create a fresh State with given volume and concentrations
Public Function MakeState(ByVal vol As Double, ByRef chem() As Double) As State
    Dim s As State
    Dim i As Long
    s.Vol = vol
    For i = 1 To METRIC_COUNT
        s.Chem(i) = chem(i)
    Next i
    MakeState = s
End Function

' Copy a State (since VBA passes UDTs by reference)
Public Function CopyState(ByRef s As State) As State
    Dim c As State
    Dim i As Long
    c.Vol = s.Vol
    c.HidVol = s.HidVol
    For i = 1 To METRIC_COUNT
        c.Chem(i) = s.Chem(i)
        c.Hidden(i) = s.Hidden(i)
    Next i
    CopyState = c
End Function
