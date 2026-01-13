Attribute VB_Name = "WQOC"
Option Explicit
' WQOC: Simple entry point for Water Quality Optimisation Calculator.
' Purpose: One button, one function.
' Dependencies: Types, Data, Sim, History
'
' Usage:
'   WQOC.Run        - Run simulation with current inputs
'   WQOC.Rollback   - Undo last run
'   WQOC.TestCore   - Quick test without worksheet I/O

' ==== Main Entry Point ========================================================

' Run simulation with current inputs
' This is the ONE function operators call
Public Sub Run()
    Dim s As State
    Dim cfg As Config
    Dim r As Result
    Dim calcMode As XlCalculation

    On Error GoTo HandleError

    ' Capture app state
    calcMode = Application.Calculation
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual

    ' Load
    s = Data.LoadState()
    cfg = Data.LoadConfig()

    ' Simulate
    r = Sim.Run(s, cfg)

    ' Save
    Data.SaveResult r
    History.RecordRun cfg, r

    ' Restore
    Application.Calculation = calcMode
    Application.ScreenUpdating = True
    Application.EnableEvents = True

    ' Show result
    ShowResult r

    Exit Sub

HandleError:
    Application.Calculation = calcMode
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    MsgBox "Error: " & Err.Description, vbExclamation, "WQOC"
End Sub

' ==== Convenience Functions ===================================================

' Rollback most recent run
Public Sub Rollback()
    If History.RollbackLast() Then
        MsgBox "Last run rolled back.", vbInformation, "WQOC"
    Else
        MsgBox "No run to rollback.", vbExclamation, "WQOC"
    End If
End Sub

' Show run count for current site
Public Sub ShowRunCount()
    Dim count As Long
    count = History.CountRuns()
    MsgBox "Runs for this site: " & count, vbInformation, "WQOC"
End Sub

' Run and return trigger day (for testing/scripting)
Public Function GetTriggerDay() As Long
    Dim s As State
    Dim cfg As Config
    Dim r As Result

    s = Data.LoadState()
    cfg = Data.LoadConfig()
    r = Sim.Run(s, cfg)

    GetTriggerDay = r.TriggerDay
End Function

' ==== Result Display ==========================================================

Private Sub ShowResult(ByRef r As Result)
    Dim msg As String

    If r.TriggerDay = Types.NO_TRIGGER Then
        msg = "No trigger reached in " & UBound(r.Snaps) & " days." & vbNewLine & _
              "Final volume: " & Format$(r.FinalState.Vol, "0.0") & " ML"
    Else
        msg = "TRIGGER REACHED" & vbNewLine & vbNewLine & _
              "Metric: " & r.TriggerMetric & vbNewLine & _
              "Day: " & r.TriggerDay & vbNewLine & _
              "Date: " & Format$(r.TriggerDate, "dd-mmm-yyyy")
    End If

    MsgBox msg, vbInformation, "WQOC Result"
End Sub

' ==== Quick Tests =============================================================

' Test the simulation core (no worksheet I/O)
Public Sub TestCore()
    Dim s As State
    Dim cfg As Config
    Dim r As Result

    ' Setup test state: 100 ML reservoir
    s.Vol = 100
    s.Chem(1) = 200  ' EC = 200

    ' Setup test config
    cfg.Mode = "Simple"
    cfg.Days = 50
    cfg.Inflow = 2      ' 2 ML/day in
    cfg.Outflow = 1     ' 1 ML/day out (net +1/day)
    cfg.TriggerVol = 150

    ' Run simulation
    r = Sim.Run(s, cfg)

    ' Report
    If r.TriggerDay = Types.NO_TRIGGER Then
        Debug.Print "No trigger. Final vol: " & r.FinalState.Vol & " ML"
    Else
        Debug.Print "TRIGGER on day " & r.TriggerDay & ": " & r.TriggerMetric
        Debug.Print "  Date: " & r.TriggerDate
        Debug.Print "  Final vol: " & r.FinalState.Vol & " ML"
    End If
End Sub

' Test two-bucket mode
Public Sub TestTwoBucket()
    Dim s As State
    Dim cfg As Config
    Dim r As Result

    ' Setup: 100 ML visible, 50 ML hidden
    s.Vol = 100
    s.HidVol = 50
    s.Chem(1) = 200     ' Visible EC
    s.Hidden(1) = 5000  ' Hidden EC mass (will mix up slowly)

    ' Two-bucket config
    cfg.Mode = "TwoBucket"
    cfg.Days = 30
    cfg.Tau = 7         ' 7-day mixing time
    cfg.Inflow = 2
    cfg.Outflow = 1
    cfg.TriggerChem(1) = 300  ' EC trigger

    ' Run
    r = Sim.Run(s, cfg)

    ' Report
    Debug.Print "Two-bucket test:"
    Debug.Print "  Start EC: " & s.Chem(1)
    Debug.Print "  End EC: " & r.FinalState.Chem(1)
    If r.TriggerDay <> Types.NO_TRIGGER Then
        Debug.Print "  TRIGGER on day " & r.TriggerDay & ": " & r.TriggerMetric
    Else
        Debug.Print "  No trigger in " & cfg.Days & " days"
    End If
End Sub
