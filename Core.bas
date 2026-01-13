Option Explicit
' Core: Type definitions (C sorts before D/M/S).
' Dependencies: None

Public Const METRIC_COUNT As Long = 7
Public Const NO_TRIGGER As Long = -1
Public Const EPS As Double = 0.000001

Public Type State
    Vol As Double
    Chem(1 To 7) As Double
    Hidden(1 To 7) As Double
    HidVol As Double
End Type

Public Type Config
    Mode As String
    Days As Long
    StartDate As Date
    Tau As Double
    Inflow As Double
    Outflow As Double
    RainVol As Double
    SurfaceFrac As Double
    InflowChem(1 To 7) As Double
    TriggerVol As Double
    TriggerChem(1 To 7) As Double
End Type

Public Type Result
    TriggerDay As Long
    TriggerDate As Date
    TriggerMetric As String
    Snaps() As State
    FinalState As State
End Type

Private mNames As Variant

Public Function MetricName(ByVal idx As Long) As String
    If IsEmpty(mNames) Then mNames = Array("EC", "F_U", "F_Mn", "SO4", "Mg", "Ca", "TAN")
    If idx >= 1 And idx <= METRIC_COUNT Then MetricName = mNames(idx - 1)
End Function

Public Function MetricNames() As Variant
    If IsEmpty(mNames) Then mNames = Array("EC", "F_U", "F_Mn", "SO4", "Mg", "Ca", "TAN")
    MetricNames = mNames
End Function

Public Function CopyState(ByRef s As State) As State
    Dim c As State, i As Long
    c.Vol = s.Vol: c.HidVol = s.HidVol
    For i = 1 To METRIC_COUNT: c.Chem(i) = s.Chem(i): c.Hidden(i) = s.Hidden(i): Next i
    CopyState = c
End Function
