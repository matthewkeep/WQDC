Option Explicit
' Core: Type definitions (C sorts before D/M/S).
' Dependencies: None

Public Const METRIC_COUNT As Long = 7
Public Const NO_TRIGGER As Long = -1
Public Const EPS As Double = 0.000001

' Chemistry metric indices (1-based to match array bounds)
Public Enum Metric
    mEC = 1
    mF_U = 2
    mF_Mn = 3
    mSO4 = 4
    mMg = 5
    mCa = 6
    mTAN = 7
End Enum

Public Type State
    Vol As Double
    Chem(1 To 7) As Double
    Hidden(1 To 7) As Double
End Type

Public Type Config
    Mode As String
    Site As String
    Days As Long
    StartDate As Date
    Tau As Double
    Inflow As Double
    Outflow As Double
    RainfallMode As String
    RainFactor As Double
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
    c.Vol = s.Vol
    For i = 1 To METRIC_COUNT: c.Chem(i) = s.Chem(i): c.Hidden(i) = s.Hidden(i): Next i
    CopyState = c
End Function

Public Function InitHiddenAtEquilibrium(ByRef s As State) As State
    ' Initializes hidden layer at equilibrium with visible layer
    ' Hidden mass = visible volume * visible concentration
    Dim init As State, i As Long
    init = CopyState(s)
    For i = 1 To METRIC_COUNT
        init.Hidden(i) = s.Vol * s.Chem(i)
    Next i
    InitHiddenAtEquilibrium = init
End Function

Public Function IsHiddenEmpty(ByRef s As State) As Boolean
    ' Returns True if hidden layer has no values (all indices near zero)
    Dim i As Long
    For i = 1 To METRIC_COUNT
        If s.Hidden(i) > EPS Then Exit Function
    Next i
    IsHiddenEmpty = True
End Function
