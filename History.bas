Option Explicit
' History: Audit trail for simulation runs.
' Dependencies: Core, Schema, Helpers, SimLog, Setup (for EnsureSiteHistoryTable)
'
' All runs are stored per-site. Tables created on-demand: tblHistory_RP1, etc.
' No Site column in table - site is encoded in table name.

' ==== Public ==================================================================

Public Sub RecordRun(ByRef s As State, ByRef cfgStd As Config, ByRef rStd As Result, _
                     ByRef cfgEnh As Config, ByRef rEnh As Result, _
                     ByVal hasEnhanced As Boolean, ByVal telemCalEnabled As Boolean, _
                     ByVal runId As String, ByVal site As String)
    ' Records single history entry per run with bundled Std and Enh results
    ' Uses State/Config passed from caller - avoids re-reading Inputs sheet
    ' Only reads IR table and SignName (not in State/Config)
    Dim tbl As ListObject, row As ListRow, i As Long
    Dim idCol As Long, tsCol As Long, dateCol As Long
    Dim actionCol As Long, loadCol As Long

    On Error GoTo Fail

    Set tbl = GetHistoryTable(site)
    If tbl Is Nothing Then Error.Trace "History.RecordRun", "No history table": Exit Sub

    ' Get column indices
    idCol = Helpers.ColIdx(tbl, Schema.HISTORY_COL_RUNID)
    tsCol = Helpers.ColIdx(tbl, Schema.HISTORY_COL_TIMESTAMP)
    dateCol = Helpers.ColIdx(tbl, Schema.HISTORY_COL_RUNDATE)
    actionCol = Helpers.ColIdx(tbl, Schema.HISTORY_COL_ACTION)
    loadCol = Helpers.ColIdx(tbl, Schema.HISTORY_COL_LOAD)
    If idCol = 0 Or actionCol = 0 Then Error.Trace "History.RecordRun", "Missing columns": Exit Sub

    ' Update existing rows' action to "Rollback" and ensure Load is set
    If Not tbl.DataBodyRange Is Nothing Then
        For i = 1 To tbl.ListRows.Count
            tbl.DataBodyRange.Cells(i, actionCol).Value = Schema.ACTION_ROLLBACK
            If loadCol > 0 Then tbl.DataBodyRange.Cells(i, loadCol).Value = "Load"
        Next i
    End If

    Set row = tbl.ListRows.Add: If row Is Nothing Then Exit Sub

    ' Get Inputs sheet for run date, preset, IR, SignName
    Dim wsInput As Worksheet, irTbl As ListObject, runDate As Date
    Set wsInput = Helpers.GetSheet(Schema.SHEET_INPUT)
    runDate = Helpers.GetDateVal(wsInput, Schema.NAME_RUN_DATE)
    If runDate = 0 Then runDate = Date  ' Fallback to today if not set

    With row.Range
        .Cells(1, idCol).Value = runId
        If tsCol > 0 Then .Cells(1, tsCol).Value = Now
        If dateCol > 0 Then .Cells(1, dateCol).Value = runDate
        .Cells(1, actionCol).Value = Schema.ACTION_CURRENT
        If loadCol > 0 Then .Cells(1, loadCol).Value = "Load"
    End With

    ' State/Config data
    Helpers.WriteCell tbl, row, Schema.HISTORY_COL_SAMPLE_DATE, cfgStd.StartDate
    Helpers.WriteCell tbl, row, Schema.HISTORY_COL_OUTFLOW, cfgStd.Outflow
    Helpers.WriteCell tbl, row, Schema.HISTORY_COL_RES_CHEM, Helpers.SerializeStateArray(s.Chem)
    Helpers.WriteCell tbl, row, Schema.HISTORY_COL_TRIGGERS, _
        Helpers.SerializeTriggerConfig(cfgStd, CStr(Helpers.ReadFromRange(wsInput, Schema.NAME_TRIGGER_PRESET) & ""))
    Helpers.WriteCell tbl, row, Schema.HISTORY_COL_HIDDEN_MASS, Helpers.SerializeStateArray(s.Hidden)

    ' IR table and SignName (from sheet)
    If Not wsInput Is Nothing Then
        Set irTbl = Helpers.GetTable(Schema.SHEET_INPUT, Schema.TABLE_IR)
        Helpers.WriteCell tbl, row, Schema.HISTORY_COL_IR_SNAPSHOT, Helpers.SerializeIRTable(irTbl)
        Helpers.WriteCell tbl, row, Schema.HISTORY_COL_SIGN_NAME, Helpers.ReadFromRange(wsInput, Schema.NAME_SIGN_OFF_NAME)
    End If

    ' Results: Days|TriggerMetric (runDate already set from Inputs sheet)
    Helpers.WriteCell tbl, row, Schema.HISTORY_COL_STD_RESULT, _
        Helpers.SerializeResult(CalcTriggerDays(cfgStd.StartDate, rStd.TriggerDay, runDate), rStd.TriggerMetric)

    If hasEnhanced Then
        Helpers.WriteCell tbl, row, Schema.HISTORY_COL_ENH_RESULT, _
            Helpers.SerializeResult(CalcTriggerDays(cfgEnh.StartDate, rEnh.TriggerDay, runDate), rEnh.TriggerMetric)
        Helpers.WriteCell tbl, row, Schema.HISTORY_COL_ENH_SETTINGS, _
            Helpers.SerializeEnhSettingsHist("On", IIf(telemCalEnabled, "On", "Off"), _
                cfgEnh.RainfallMode, cfgEnh.RainFactor, cfgEnh.Mode, cfgEnh.Tau, cfgEnh.SurfaceFrac)
    Else
        Helpers.WriteCell tbl, row, Schema.HISTORY_COL_ENH_RESULT, ""
        Helpers.WriteCell tbl, row, Schema.HISTORY_COL_ENH_SETTINGS, Helpers.SerializeEnhSettingsHist("Off", "", "", 0, "", 0, 0)
    End If

    row.Range.WrapText = False
    SortHistoryTable tbl
    Exit Sub

Fail:
    Error.TraceErr "History.RecordRun"
End Sub

' ==== Private Helpers =========================================================

Private Function CalcTriggerDays(ByVal startDate As Date, ByVal triggerDay As Long, ByVal runDate As Date) As Long
    If triggerDay = Core.NO_TRIGGER Then
        CalcTriggerDays = Core.NO_TRIGGER
    Else
        CalcTriggerDays = CLng((startDate + triggerDay) - runDate)
    End If
End Function

Private Sub MarkLastRowCurrent(ByVal tbl As ListObject)
    Dim actionCol As Long
    If tbl.ListRows.Count = 0 Then Exit Sub
    actionCol = Helpers.ColIdx(tbl, Schema.HISTORY_COL_ACTION)
    If actionCol > 0 Then tbl.DataBodyRange.Cells(tbl.ListRows.Count, actionCol).Value = Schema.ACTION_CURRENT
End Sub

Public Function LoadSettings(ByVal runId As String, ByVal site As String) As Boolean
    ' Restores config from history row to Inputs sheet (no deletion, no run)
    ' Setting Sample Date triggers hidden mass load via Events.OnInputsChange
    Dim tbl As ListObject, ws As Worksheet, i As Long, rowIdx As Long
    Dim idCol As Long

    Set tbl = GetHistoryTable(site)
    If tbl Is Nothing Then Exit Function
    If tbl.DataBodyRange Is Nothing Then Exit Function

    idCol = Helpers.ColIdx(tbl, Schema.HISTORY_COL_RUNID)
    If idCol = 0 Then Exit Function

    ' Find row by RunId
    For i = 1 To tbl.ListRows.Count
        If tbl.DataBodyRange.Cells(i, idCol).Value = runId Then
            rowIdx = i
            Exit For
        End If
    Next i
    If rowIdx = 0 Then Exit Function

    Set ws = Helpers.GetSheet(Schema.SHEET_INPUT)
    If ws Is Nothing Then Exit Function

    Dim irTbl As ListObject, limitRng As Range, chemRng As Range, hiddenRng As Range
    Dim enabled As String, telemCal As String, rainfallMode As String, mixingModel As String
    Dim rainFactor As Double, tau As Double, surfaceFrac As Double, trigVol As Double, trigPreset As String
    Dim stdDays As Long, stdMetric As String, enhDays As Long, enhMetric As String
    Dim runDateVal As Variant, outflowVal As Variant

    On Error Resume Next
    Application.EnableEvents = False

    ws.Range(Schema.NAME_SITE).Value = site

    ' Dates
    runDateVal = Helpers.ReadCell(tbl, rowIdx, Schema.HISTORY_COL_RUNDATE)
    If IsDate(runDateVal) Then ws.Range(Schema.NAME_RUN_DATE).Value = runDateVal
    ws.Range(Schema.NAME_SAMPLE_DATE).Value = Helpers.ReadCell(tbl, rowIdx, Schema.HISTORY_COL_SAMPLE_DATE)

    ' Outflow
    outflowVal = Helpers.ReadCell(tbl, rowIdx, Schema.HISTORY_COL_OUTFLOW)
    If IsNumeric(outflowVal) And outflowVal > 0 Then ws.Range(Schema.NAME_OUTPUT).Value = outflowVal

    ' EnhSettings
    Helpers.DeserializeEnhSettingsHist Helpers.ReadCellStr(tbl, rowIdx, Schema.HISTORY_COL_ENH_SETTINGS), _
        enabled, telemCal, rainfallMode, rainFactor, mixingModel, tau, surfaceFrac
    ws.Range(Schema.NAME_ENHANCED_MODE).Value = IIf(Len(enabled) > 0, enabled, "Off")
    If UCase$(enabled) = "ON" Then
        ws.Range(Schema.NAME_TELEM_CAL).Value = telemCal
        ws.Range(Schema.NAME_RAINFALL_MODE).Value = rainfallMode
        ws.Range(Schema.NAME_RAIN_FACTOR).Value = rainFactor
        ws.Range(Schema.NAME_MIXING_MODEL).Value = mixingModel
        ws.Range(Schema.NAME_TAU).Value = tau
        ws.Range(Schema.NAME_SURFACE_FRACTION).Value = surfaceFrac
    End If

    ' Triggers
    Set limitRng = Helpers.GetRng(ws, Schema.NAME_LIMIT_ROW)
    Helpers.DeserializeVolChem Helpers.ReadCellStr(tbl, rowIdx, Schema.HISTORY_COL_TRIGGERS), trigVol, limitRng, trigPreset
    ws.Range(Schema.NAME_TRIGGER_VOL).Value = trigVol
    ws.Range(Schema.NAME_TRIGGER_PRESET).Value = trigPreset

    ' Chemistry
    Set chemRng = Helpers.GetRng(ws, Schema.NAME_RES_ROW)
    Helpers.DeserializeToRange Helpers.ReadCellStr(tbl, rowIdx, Schema.HISTORY_COL_RES_CHEM), chemRng, Core.METRIC_COUNT

    ' Hidden mass
    Set hiddenRng = Helpers.GetRng(ws, Schema.NAME_HIDDEN_MASS)
    Helpers.DeserializeToColumn Helpers.ReadCellStr(tbl, rowIdx, Schema.HISTORY_COL_HIDDEN_MASS), hiddenRng, Core.METRIC_COUNT

    ' IR table
    Set irTbl = Helpers.GetTable(Schema.SHEET_INPUT, Schema.TABLE_IR)
    Helpers.DeserializeIRTable Helpers.ReadCellStr(tbl, rowIdx, Schema.HISTORY_COL_IR_SNAPSHOT), irTbl

    ' Trigger display
    Helpers.DeserializeResult Helpers.ReadCellStr(tbl, rowIdx, Schema.HISTORY_COL_STD_RESULT), stdDays, stdMetric
    If stdDays <> Core.NO_TRIGGER Then ws.Range(Schema.NAME_STD_TRIGGER).Value = stdDays _
        Else ws.Range(Schema.NAME_STD_TRIGGER).ClearContents

    Helpers.DeserializeResult Helpers.ReadCellStr(tbl, rowIdx, Schema.HISTORY_COL_ENH_RESULT), enhDays, enhMetric
    If enhDays <> Core.NO_TRIGGER And UCase$(enabled) = "ON" Then ws.Range(Schema.NAME_ENH_TRIGGER).Value = enhDays _
        Else ws.Range(Schema.NAME_ENH_TRIGGER).ClearContents

    ' Pred_Mode
    ws.Range(Schema.NAME_PRED_MODE).Value = IIf(UCase$(enabled) = "ON", "Enhanced", "Standard")

    Application.EnableEvents = True
    On Error GoTo 0

    LoadSettings = True
End Function

Public Function CountRuns(ByVal site As String) As Long
    ' Returns count of runs for site
    Dim tbl As ListObject

    Set tbl = GetHistoryTable(site)
    If tbl Is Nothing Then Exit Function

    CountRuns = tbl.ListRows.Count
End Function

Public Function GetCurrentRunId(ByVal site As String) As String
    ' Returns runId of current (latest) history entry for site
    Dim tbl As ListObject, idCol As Long

    Set tbl = GetHistoryTable(site)
    If tbl Is Nothing Then Exit Function
    If tbl.DataBodyRange Is Nothing Then Exit Function
    If tbl.ListRows.Count = 0 Then Exit Function

    idCol = Helpers.ColIdx(tbl, Schema.HISTORY_COL_RUNID)
    If idCol = 0 Then Exit Function

    ' Last row is current (table sorted by date+timestamp)
    GetCurrentRunId = tbl.DataBodyRange.Cells(tbl.ListRows.Count, idCol).Value
End Function

Public Function RollbackLast(ByVal site As String) As Boolean
    ' Deletes last run from history AND log entries after that run's date
    Dim tbl As ListObject, dateCol As Long, runDate As Date

    Set tbl = GetHistoryTable(site)
    If tbl Is Nothing Or tbl.ListRows.Count = 0 Then Exit Function

    dateCol = Helpers.ColIdx(tbl, Schema.HISTORY_COL_RUNDATE)
    If dateCol = 0 Then Exit Function

    ' Delete log entries after the previous run's date
    If tbl.ListRows.Count > 1 Then
        SimLog.DeleteAfterDate tbl.ListRows(tbl.ListRows.Count - 1).Range.Cells(1, dateCol).Value, site
    Else
        runDate = tbl.ListRows(tbl.ListRows.Count).Range.Cells(1, dateCol).Value
        SimLog.DeleteAfterDate runDate - 1, site
    End If

    tbl.ListRows(tbl.ListRows.Count).Delete
    MarkLastRowCurrent tbl
    RollbackLast = True
End Function

Public Function RollbackTo(ByVal targetRunId As String, ByVal site As String) As Long
    ' Deletes all runs AFTER targetRunId for site (Jenga model), returns count removed
    Dim tbl As ListObject, i As Long, targetIdx As Long, removed As Long
    Dim idCol As Long, dateCol As Long, targetRunDate As Date

    Set tbl = GetHistoryTable(site)
    If tbl Is Nothing Or tbl.ListRows.Count = 0 Then Exit Function

    idCol = Helpers.ColIdx(tbl, Schema.HISTORY_COL_RUNID)
    dateCol = Helpers.ColIdx(tbl, Schema.HISTORY_COL_RUNDATE)
    If idCol = 0 Or dateCol = 0 Then Exit Function

    SortHistoryTable tbl

    ' Find target run
    For i = 1 To tbl.ListRows.Count
        If tbl.ListRows(i).Range.Cells(1, idCol).Value = targetRunId Then
            targetIdx = i
            targetRunDate = tbl.ListRows(i).Range.Cells(1, dateCol).Value
            Exit For
        End If
    Next i
    If targetIdx = 0 Then Exit Function

    SimLog.DeleteAfterDate targetRunDate, site

    ' Delete rows after target (backwards to avoid index issues)
    For i = tbl.ListRows.Count To targetIdx + 1 Step -1
        tbl.ListRows(i).Delete
        removed = removed + 1
    Next i

    MarkLastRowCurrent tbl
    RollbackTo = removed
End Function

' ==== Table Access ===========================================================

Private Function GetHistoryTable(ByVal site As String) As ListObject
    Set GetHistoryTable = Helpers.GetSiteTable(Schema.SHEET_RECORD, Schema.HISTORY_TABLE_PREFIX, site)
    If GetHistoryTable Is Nothing Then
        Setup.EnsureSiteHistoryTable site
        Set GetHistoryTable = Helpers.GetSiteTable(Schema.SHEET_RECORD, Schema.HISTORY_TABLE_PREFIX, site)
    End If
End Function

Private Sub SortHistoryTable(ByVal tbl As ListObject)
    If tbl Is Nothing Or tbl.ListRows.Count <= 1 Then Exit Sub
    tbl.Sort.SortFields.Clear
    tbl.Sort.SortFields.Add Key:=tbl.ListColumns(Schema.HISTORY_COL_RUNDATE).Range, Order:=xlAscending
    tbl.Sort.SortFields.Add Key:=tbl.ListColumns(Schema.HISTORY_COL_TIMESTAMP).Range, Order:=xlAscending
    tbl.Sort.Apply
End Sub
