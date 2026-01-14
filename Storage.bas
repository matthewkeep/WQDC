Option Explicit
' Storage: Consolidated persistence for simulation runs.
' Dependencies: Core, SimLog, History, Data
'
' Consolidates the three-destination write pattern into a single call:
'   Storage.SaveRun record
' Instead of:
'   SimLog.WriteLog ...
'   History.RecordRun ...
'   Data.SaveResult ...

' ==== Run Record Type ===========================================================

Public Type RunRecord
    RunId As String
    Site As String
    RunType As String       ' "Standard" or "Enhanced"
    Config As Config
    Result As Result
End Type

' ==== Public API ================================================================

Public Function CreateRecord(ByVal runId As String, ByVal site As String, _
                            ByVal runType As String, ByRef cfg As Config, _
                            ByRef r As Result) As RunRecord
    ' Factory function to create a RunRecord
    Dim rec As RunRecord
    rec.RunId = runId
    rec.Site = site
    rec.RunType = runType
    rec.Config = cfg
    rec.Result = r
    CreateRecord = rec
End Function

Public Sub SaveRun(ByRef rec As RunRecord)
    ' Persists run to all three destinations:
    '   1. tblLive (SimLog) - date-centric predictions
    '   2. tblHistory (History) - audit trail with config snapshot
    '   3. Inputs sheet (Data) - trigger display and predicted row

    On Error GoTo Fail

    ' 1. Write to live log (tblLive)
    SimLog.WriteLog rec.Result, rec.Config, rec.RunId, rec.Site

    ' 2. Record in history (tblHistory)
    History.RecordRun rec.Config, rec.Result, rec.RunId, rec.Site

    ' 3. Update Inputs sheet display
    Data.SaveResult rec.Result, rec.RunType

    Exit Sub

Fail:
    Error.TraceErr "Storage.SaveRun"
End Sub

Public Sub SaveRunPair(ByRef recStd As RunRecord, ByRef recEnh As RunRecord)
    ' Saves both Standard and Enhanced runs
    ' Use when Enhanced mode is enabled
    SaveRun recStd
    SaveRun recEnh
End Sub
