Option Explicit
' History: Audit trail for simulation runs.
' Dependencies: Core, Schema, SimLog, Setup (for EnsureSiteHistoryTable)
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

    With row.Range
        .Cells(1, idCol).Value = runId
        If tsCol > 0 Then .Cells(1, tsCol).Value = Now
        If dateCol > 0 Then .Cells(1, dateCol).Value = Date  ' Actual run date (today)
        .Cells(1, actionCol).Value = Schema.ACTION_CURRENT
        If loadCol > 0 Then .Cells(1, loadCol).Value = "Load"
    End With

    ' Use State/Config data (already loaded by caller)
    WriteColIfExists tbl, row, Schema.HISTORY_COL_SAMPLE_DATE, cfgStd.StartDate
    WriteColIfExists tbl, row, Schema.HISTORY_COL_RES_CHEM, SerializeStateArray(s.Chem)
    WriteColIfExists tbl, row, Schema.HISTORY_COL_TRIGGERS, SerializeTriggerConfig(cfgStd)
    WriteColIfExists tbl, row, Schema.HISTORY_COL_HIDDEN_MASS, SerializeStateArray(s.Hidden)

    ' IR table and SignName must still be read (not in State/Config)
    Dim wsInput As Worksheet, irTbl As ListObject
    Set wsInput = Helpers.GetSheet(Schema.SHEET_INPUT)
    If Not wsInput Is Nothing Then
        Set irTbl = Helpers.GetTable(Schema.SHEET_INPUT, Schema.TABLE_IR)
        WriteColIfExists tbl, row, Schema.HISTORY_COL_IR_SNAPSHOT, Helpers.SerializeIRTable(irTbl)
        WriteColIfExists tbl, row, Schema.HISTORY_COL_SIGN_NAME, _
            Helpers.ReadFromRange(wsInput, Schema.NAME_SIGN_OFF_NAME)
    End If

    ' Standard results: Days|TriggerMetric (Days = days from run date)
    Dim runDate As Date, daysStd As Long, daysEnh As Long
    runDate = Date
    If rStd.TriggerDay = Core.NO_TRIGGER Then
        daysStd = Core.NO_TRIGGER
    Else
        daysStd = CLng((cfgStd.StartDate + rStd.TriggerDay) - runDate)
    End If
    WriteColIfExists tbl, row, Schema.HISTORY_COL_STD_RESULT, _
        Helpers.SerializeResult(daysStd, rStd.TriggerMetric)

    ' Enhanced results and settings
    If hasEnhanced Then
        If rEnh.TriggerDay = Core.NO_TRIGGER Then
            daysEnh = Core.NO_TRIGGER
        Else
            daysEnh = CLng((cfgEnh.StartDate + rEnh.TriggerDay) - runDate)
        End If
        WriteColIfExists tbl, row, Schema.HISTORY_COL_ENH_RESULT, _
            Helpers.SerializeResult(daysEnh, rEnh.TriggerMetric)
        WriteColIfExists tbl, row, Schema.HISTORY_COL_ENH_SETTINGS, _
            Helpers.SerializeEnhSettingsHist("On", _
                IIf(telemCalEnabled, "On", "Off"), _
                cfgEnh.RainfallMode, cfgEnh.RainFactor, cfgEnh.Mode, _
                cfgEnh.Tau, cfgEnh.SurfaceFrac)
    Else
        WriteColIfExists tbl, row, Schema.HISTORY_COL_ENH_RESULT, ""
        WriteColIfExists tbl, row, Schema.HISTORY_COL_ENH_SETTINGS, _
            Helpers.SerializeEnhSettingsHist("Off", "", "", 0, "", 0, 0)
    End If

    row.Range.WrapText = False
    SortHistoryTable tbl
    Exit Sub

Fail:
    Error.TraceErr "History.RecordRun"
End Sub

Private Function SerializeStateArray(ByRef arr() As Double) As String
    ' Serializes State.Chem or State.Hidden array (1-based) to pipe-delimited string
    ' Consistent with Helpers.SerializeRange pattern
    Dim i As Long, parts() As String
    ReDim parts(0 To Core.METRIC_COUNT - 1)
    For i = 1 To Core.METRIC_COUNT
        parts(i - 1) = CStr(arr(i))
    Next i
    SerializeStateArray = Join(parts, "|")
End Function

Private Function SerializeTriggerConfig(ByRef cfg As Config) As String
    ' Serializes triggers from Config: Vol|EC|F_U|F_Mn|SO4|Mg|Ca|TAN
    ' Consistent with Helpers.SerializeTriggers pattern
    Dim i As Long, parts(0 To 7) As String
    parts(0) = CStr(cfg.TriggerVol)
    For i = 1 To Core.METRIC_COUNT
        parts(i) = CStr(cfg.TriggerChem(i))
    Next i
    SerializeTriggerConfig = Join(parts, "|")
End Function

' ==== Private Helpers =========================================================

Private Sub WriteColIfExists(ByVal tbl As ListObject, ByVal row As ListRow, ByVal colName As String, ByVal v As Variant)
    ' Writes value to column if it exists (for backward compatibility with old tables)
    Dim col As Long
    col = Helpers.ColIdx(tbl, colName)
    If col > 0 Then row.Range.Cells(1, col).Value = v
End Sub

Private Function ReadColIfExists(ByVal tbl As ListObject, ByVal rowIdx As Long, ByVal colName As String) As Variant
    ' Reads value from column if it exists (for backward compatibility with old tables)
    Dim col As Long
    col = Helpers.ColIdx(tbl, colName)
    If col > 0 Then ReadColIfExists = tbl.DataBodyRange.Cells(rowIdx, col).Value
End Function

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

    ' Restore settings to Inputs sheet
    Dim irTbl As ListObject
    Dim chemRng As Range, limitRng As Range, hiddenRng As Range
    Dim enhSettingsStr As String, enabled As String, telemCal As String, rainfallMode As String
    Dim rainFactor As Double, mixingModel As String, tau As Double, surfaceFrac As Double
    Dim trigVol As Double

    On Error Resume Next
    Application.EnableEvents = False

    ' Site (ensures Replay runs against correct log)
    ws.Range(Schema.NAME_SITE).Value = site

    ' Run date
    ws.Range(Schema.NAME_RUN_DATE).Value = ReadColIfExists(tbl, rowIdx, Schema.HISTORY_COL_RUNDATE)

    ' Sample date
    ws.Range(Schema.NAME_SAMPLE_DATE).Value = ReadColIfExists(tbl, rowIdx, Schema.HISTORY_COL_SAMPLE_DATE)

    ' EnhSettings: Enabled|TelemCal|RainfallMode|RainFactor|MixingModel|Tau|SurfaceFrac
    enhSettingsStr = CStr(ReadColIfExists(tbl, rowIdx, Schema.HISTORY_COL_ENH_SETTINGS) & "")
    Helpers.DeserializeEnhSettingsHist enhSettingsStr, enabled, telemCal, rainfallMode, rainFactor, mixingModel, tau, surfaceFrac
    ws.Range(Schema.NAME_ENHANCED_MODE).Value = IIf(Len(enabled) > 0, enabled, "Off")
    If UCase$(enabled) = "ON" Then
        ws.Range(Schema.NAME_TELEM_CAL).Value = telemCal
        ws.Range(Schema.NAME_RAINFALL_MODE).Value = rainfallMode
        ws.Range(Schema.NAME_RAIN_FACTOR).Value = rainFactor
        ws.Range(Schema.NAME_MIXING_MODEL).Value = mixingModel
        ws.Range(Schema.NAME_TAU).Value = tau
        ws.Range(Schema.NAME_SURFACE_FRACTION).Value = surfaceFrac
    End If

    ' Triggers: Vol|EC|F_U|F_Mn|SO4|Mg|Ca|TAN
    Set limitRng = Helpers.GetRng(ws, Schema.NAME_LIMIT_ROW)
    Helpers.DeserializeTriggers CStr(ReadColIfExists(tbl, rowIdx, Schema.HISTORY_COL_TRIGGERS) & ""), trigVol, limitRng
    ws.Range(Schema.NAME_TRIGGER_VOL).Value = trigVol

    ' Reservoir chemistry
    Set chemRng = Helpers.GetRng(ws, Schema.NAME_RES_ROW)
    Helpers.DeserializeToRange CStr(ReadColIfExists(tbl, rowIdx, Schema.HISTORY_COL_RES_CHEM) & ""), chemRng, Core.METRIC_COUNT

    ' Hidden mass
    Set hiddenRng = Helpers.GetRng(ws, Schema.NAME_HIDDEN_MASS)
    Helpers.DeserializeToColumn CStr(ReadColIfExists(tbl, rowIdx, Schema.HISTORY_COL_HIDDEN_MASS) & ""), hiddenRng, Core.METRIC_COUNT

    ' IR table
    Set irTbl = Helpers.GetTable(Schema.SHEET_INPUT, Schema.TABLE_IR)
    Helpers.DeserializeIRTable CStr(ReadColIfExists(tbl, rowIdx, Schema.HISTORY_COL_IR_SNAPSHOT) & ""), irTbl

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
    Dim tbl As ListObject, runDate As Date
    Dim dateCol As Long

    Set tbl = GetHistoryTable(site)
    If tbl Is Nothing Then Exit Function
    If tbl.ListRows.Count = 0 Then Exit Function

    dateCol = Helpers.ColIdx(tbl, Schema.HISTORY_COL_RUNDATE)
    If dateCol = 0 Then Exit Function

    ' Delete log entries after the previous run's date
    If tbl.ListRows.Count > 1 Then
        ' Roll back to previous run's date
        Dim prevRunDate As Date
        prevRunDate = tbl.ListRows(tbl.ListRows.Count - 1).Range.Cells(1, dateCol).Value
        SimLog.DeleteAfterDate prevRunDate, site
    Else
        ' Only run - get its run date and delete everything after the day before
        runDate = tbl.ListRows(tbl.ListRows.Count).Range.Cells(1, dateCol).Value
        SimLog.DeleteAfterDate runDate - 1, site
    End If

    tbl.ListRows(tbl.ListRows.Count).Delete

    ' Update new last row to Current
    If tbl.ListRows.Count > 0 Then
        Dim actionCol As Long
        actionCol = Helpers.ColIdx(tbl, Schema.HISTORY_COL_ACTION)
        If actionCol > 0 Then
            tbl.DataBodyRange.Cells(tbl.ListRows.Count, actionCol).Value = Schema.ACTION_CURRENT
        End If
    End If

    RollbackLast = True
End Function

Public Function RollbackTo(ByVal targetRunId As String, ByVal site As String) As Long
    ' Deletes all runs AFTER targetRunId for site (Jenga model)
    ' Returns count of runs removed
    Dim tbl As ListObject
    Dim i As Long, targetIdx As Long, removed As Long
    Dim targetRunDate As Date
    Dim idCol As Long, dateCol As Long

    Set tbl = GetHistoryTable(site)
    If tbl Is Nothing Then Exit Function
    If tbl.ListRows.Count = 0 Then Exit Function

    idCol = Helpers.ColIdx(tbl, Schema.HISTORY_COL_RUNID)
    dateCol = Helpers.ColIdx(tbl, Schema.HISTORY_COL_RUNDATE)
    If idCol = 0 Or dateCol = 0 Then Exit Function

    ' Ensure table is sorted by date then timestamp (handles multiple runs per day)
    SortHistoryTable tbl

    ' Find the target run to get its run date
    targetIdx = 0
    For i = 1 To tbl.ListRows.Count
        If tbl.ListRows(i).Range.Cells(1, idCol).Value = targetRunId Then
            targetIdx = i
            targetRunDate = tbl.ListRows(i).Range.Cells(1, dateCol).Value
            Exit For
        End If
    Next i
    If targetIdx = 0 Then Exit Function

    ' Delete log entries after target run date (preserves data up to when sim was run)
    SimLog.DeleteAfterDate targetRunDate, site

    ' Delete all history rows that come AFTER target
    ' Work backwards from end to avoid index issues
    For i = tbl.ListRows.Count To targetIdx + 1 Step -1
        tbl.ListRows(i).Delete
        removed = removed + 1
    Next i

    ' Update target row to Current
    If tbl.ListRows.Count > 0 Then
        Dim actionCol As Long
        actionCol = Helpers.ColIdx(tbl, Schema.HISTORY_COL_ACTION)
        If actionCol > 0 Then
            tbl.DataBodyRange.Cells(tbl.ListRows.Count, actionCol).Value = Schema.ACTION_CURRENT
        End If
    End If

    RollbackTo = removed
End Function

' ==== Table Access ===========================================================

Private Function GetHistoryTable(ByVal site As String) As ListObject
    ' Returns site's history table, creating it if necessary
    Dim ws As Worksheet, tblName As String

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(Schema.SHEET_RECORD)
    On Error GoTo 0
    If ws Is Nothing Then Exit Function

    tblName = Helpers.HistoryTableName(site)

    ' Try to get existing table
    On Error Resume Next
    Set GetHistoryTable = ws.ListObjects(tblName)
    On Error GoTo 0

    ' Create if doesn't exist
    If GetHistoryTable Is Nothing Then
        Setup.EnsureSiteHistoryTable site
        On Error Resume Next
        Set GetHistoryTable = ws.ListObjects(tblName)
        On Error GoTo 0
    End If
End Function

Private Sub SortHistoryTable(ByVal tbl As ListObject)
    ' Sorts history table by RunDate then Timestamp (oldest first, newest last = Current)
    If tbl Is Nothing Then Exit Sub
    If tbl.ListRows.Count <= 1 Then Exit Sub

    tbl.Sort.SortFields.Clear
    tbl.Sort.SortFields.Add Key:=tbl.ListColumns(Schema.HISTORY_COL_RUNDATE).Range, Order:=xlAscending
    tbl.Sort.SortFields.Add Key:=tbl.ListColumns(Schema.HISTORY_COL_TIMESTAMP).Range, Order:=xlAscending
    tbl.Sort.Apply
End Sub
