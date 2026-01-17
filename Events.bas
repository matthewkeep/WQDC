Option Explicit
' Events: Worksheet event handlers.
' Dependencies: Loader, Schema, WQOC, History, Data, RRState, Helpers, Setup

' Module-level state for site change tracking
Private mPrevSite As String

' NOTE: To enable events, add this code to each sheet module
' (right-click sheet tab > View Code):
'
' === Inputs sheet ===
'   Private Sub Worksheet_SelectionChange(ByVal Target As Range)
'       Events.OnInputsSelectionChange Target
'   End Sub
'   Private Sub Worksheet_Change(ByVal Target As Range)
'       Events.OnInputsChange Target
'   End Sub
'   Private Sub Worksheet_BeforeDoubleClick(ByVal Target As Range, Cancel As Boolean)
'       Events.OnInputsDoubleClick Target, Cancel
'   End Sub
'
' === Record sheet ===
'   Private Sub Worksheet_BeforeDoubleClick(ByVal Target As Range, Cancel As Boolean)
'       Events.OnHistoryDoubleClick Target, Cancel
'   End Sub

' ==== Selection Events ========================================================

Public Sub OnInputsSelectionChange(ByVal Target As Range)
    ' Captures site value BEFORE user changes it
    If Target.Cells.Count = 1 And MatchesRange(Target, Schema.NAME_SITE) Then
        mPrevSite = Trim$(CStr(Target.Value))
    End If
End Sub

' ==== Change Events ===========================================================

Public Sub OnInputsChange(ByVal Target As Range)
    Dim v As String

    If Target.Cells.Count > 1 Then Exit Sub
    On Error Resume Next
    v = Trim$(CStr(Target.Value))
    On Error GoTo 0

    ' Site change
    If MatchesRange(Target, Schema.NAME_SITE) And Len(v) > 0 Then
        HandleSiteChange mPrevSite, v
        mPrevSite = v
        Exit Sub
    End If

    ' Date fields - validate, then handle sample date specifics
    If MatchesRange(Target, Schema.NAME_RUN_DATE) Or MatchesRange(Target, Schema.NAME_SAMPLE_DATE) Then
        If Not ValidateDateEntry(Target) Then Exit Sub
        If MatchesRange(Target, Schema.NAME_SAMPLE_DATE) Then
            Dim site As String, sampleDate As Date
            site = Data.GetSite()
            sampleDate = Helpers.GetDateVal(Target.Worksheet, Schema.NAME_SAMPLE_DATE)
            If Len(site) > 0 And sampleDate > 0 Then Data.LoadHiddenForDate site, sampleDate
        End If
        Exit Sub
    End If

    ' Trigger preset change
    If MatchesRange(Target, Schema.NAME_TRIGGER_PRESET) And Len(v) > 0 Then
        Loader.LoadTriggerPreset v
    End If
End Sub

Private Sub HandleSiteChange(ByVal oldSite As String, ByVal newSite As String)
    If Len(Trim$(oldSite)) > 0 Then RRState.Save oldSite
    If Not RRState.Load(newSite) Then Loader.LoadSiteData newSite
End Sub

' ==== Double-Click Events =====================================================

Public Sub OnInputsDoubleClick(ByVal Target As Range, ByRef Cancel As Boolean)
    Dim ws As Worksheet
    Set ws = Target.Worksheet

    ' Action cells (Run, Load Latest)
    If DispatchAction(Target, ws, Schema.NAME_RUN_CELL, "WQOC.Run") Then Cancel = True: Exit Sub
    If DispatchAction(Target, ws, Schema.NAME_LOAD_CELL, "Loader.LoadLatest") Then Cancel = True: Exit Sub

    ' Toggle cells
    If Toggle(Target, ws, Schema.NAME_ENHANCED_MODE, "On", "Off") Then Cancel = True: Exit Sub
    If Toggle(Target, ws, Schema.NAME_TELEM_CAL, "On", "Off") Then Cancel = True: Exit Sub
    If TogglePredMode(Target, ws) Then Cancel = True: Exit Sub

    ' IR table interactions
    If HandleIRClick(Target) Then Cancel = True
End Sub

Public Sub OnHistoryDoubleClick(ByVal Target As Range, ByRef Cancel As Boolean)
    Dim ws As Worksheet, tbl As ListObject, lo As ListObject
    Dim idCol As Long, actionCol As Long, loadCol As Long, rowIdx As Long
    Dim runId As String, site As String

    Set ws = Target.Worksheet

    ' Find clicked history table
    For Each lo In ws.ListObjects
        If Left$(lo.Name, Len(Schema.HISTORY_TABLE_PREFIX)) = Schema.HISTORY_TABLE_PREFIX Then
            If Helpers.HasData(lo) And Not Intersect(Target, lo.Range) Is Nothing Then
                Set tbl = lo
                Exit For
            End If
        End If
    Next lo
    If tbl Is Nothing Then Exit Sub

    ' Get row context
    site = Mid$(tbl.Name, Len(Schema.HISTORY_TABLE_PREFIX) + 1)
    idCol = Helpers.ColIdx(tbl, Schema.HISTORY_COL_RUNID)
    actionCol = Helpers.ColIdx(tbl, Schema.HISTORY_COL_ACTION)
    loadCol = Helpers.ColIdx(tbl, Schema.HISTORY_COL_LOAD)
    If idCol = 0 Or actionCol = 0 Then Exit Sub

    rowIdx = Target.Row - tbl.DataBodyRange.Row + 1
    If rowIdx < 1 Or rowIdx > tbl.ListRows.Count Then Exit Sub
    runId = tbl.DataBodyRange.Cells(rowIdx, idCol).Value

    ' Load column - restore settings only
    If loadCol > 0 And Not Intersect(Target, tbl.DataBodyRange.Columns(loadCol)) Is Nothing Then
        Cancel = True
        If History.LoadSettings(runId, site) Then MsgBox "Settings loaded from " & runId, vbInformation, "WQOC"
        Exit Sub
    End If

    ' Action column - rollback
    If Not Intersect(Target, tbl.DataBodyRange.Columns(actionCol)) Is Nothing Then
        Cancel = True
        If rowIdx = tbl.ListRows.Count Then
            MsgBox "This is the current run.", vbInformation, "WQOC"
        ElseIf MsgBox("Rollback to run " & runId & "?" & vbNewLine & _
                      "This will remove all runs after this one and re-run.", vbYesNo + vbQuestion, "WQOC") = vbYes Then
            History.RollbackTo runId, site
            RefreshHistoryActions tbl
            History.LoadSettings runId, site
            WQOC.Replay
        End If
    End If
End Sub

' ==== IR Table Actions ========================================================

Private Function HandleIRClick(ByVal Target As Range) As Boolean
    Dim tbl As ListObject, actionCol As Long, activeCol As Long, rowIdx As Long

    Set tbl = Helpers.GetTable(Schema.SHEET_INPUT, Schema.TABLE_IR)
    If tbl Is Nothing Then Exit Function

    actionCol = Helpers.ColIdx(tbl, Schema.IR_COL_ACTION)
    activeCol = Helpers.ColIdx(tbl, Schema.IR_COL_ACTIVE)

    ' Header click - add row
    If actionCol > 0 And Not Intersect(Target, tbl.HeaderRowRange.Cells(1, actionCol)) Is Nothing Then
        AddIRRow tbl
        HandleIRClick = True
        Exit Function
    End If

    ' Data row click
    If Not Helpers.HasData(tbl) Then Exit Function
    rowIdx = Target.Row - tbl.DataBodyRange.Row + 1
    If rowIdx < 1 Or rowIdx > tbl.ListRows.Count Then Exit Function

    ' Active column - toggle
    If activeCol > 0 And Not Intersect(Target, tbl.DataBodyRange.Columns(activeCol)) Is Nothing Then
        ToggleCell tbl.DataBodyRange.Cells(rowIdx, activeCol), "Yes", "No"
        HandleIRClick = True
        Exit Function
    End If

    ' Action column - remove
    If actionCol > 0 And Not Intersect(Target, tbl.DataBodyRange.Columns(actionCol)) Is Nothing Then
        tbl.ListRows(rowIdx).Delete
        HandleIRClick = True
    End If
End Function

Private Sub AddIRRow(ByVal tbl As ListObject)
    Dim newRow As ListRow, activeCol As Long, isFirst As Boolean
    isFirst = Not Helpers.HasData(tbl)
    Set newRow = tbl.ListRows.Add
    activeCol = Helpers.ColIdx(tbl, Schema.IR_COL_ACTIVE)
    If activeCol > 0 Then newRow.Range.Cells(1, activeCol).Value = "Yes"
    Helpers.InitIRRowAction newRow.Range, tbl
    If isFirst Then Setup.ApplyIRActiveConditionalFormat tbl
End Sub

Private Sub RefreshHistoryActions(ByVal tbl As ListObject)
    Dim i As Long, actionCol As Long
    If Not Helpers.HasData(tbl) Then Exit Sub
    actionCol = Helpers.ColIdx(tbl, Schema.HISTORY_COL_ACTION)
    If actionCol = 0 Then Exit Sub
    For i = 1 To tbl.ListRows.Count
        tbl.DataBodyRange.Cells(i, actionCol).Value = IIf(i = tbl.ListRows.Count, Schema.ACTION_CURRENT, Schema.ACTION_ROLLBACK)
    Next i
End Sub

' ==== Helpers =================================================================

Private Function MatchesRange(ByVal Target As Range, ByVal nm As String) As Boolean
    Dim rng As Range
    On Error Resume Next
    Set rng = Target.Worksheet.Range(nm)
    On Error GoTo 0
    If Not rng Is Nothing Then MatchesRange = Not Intersect(Target, rng) Is Nothing
End Function

Private Function DispatchAction(ByVal Target As Range, ByVal ws As Worksheet, ByVal nm As String, ByVal action As String) As Boolean
    ' Check if target matches named range; if so, run action and return True
    Dim rng As Range
    On Error Resume Next
    Set rng = ws.Range(nm)
    On Error GoTo 0
    If rng Is Nothing Or Intersect(Target, rng) Is Nothing Then Exit Function

    Select Case action
        Case "WQOC.Run": WQOC.Run
        Case "Loader.LoadLatest": Loader.LoadLatest
    End Select
    DispatchAction = True
End Function

Private Function Toggle(ByVal Target As Range, ByVal ws As Worksheet, ByVal nm As String, ByVal valA As String, ByVal valB As String) As Boolean
    ' Generic toggle between two values; returns True if handled
    Dim rng As Range
    On Error Resume Next
    Set rng = ws.Range(nm)
    On Error GoTo 0
    If rng Is Nothing Or Intersect(Target, rng) Is Nothing Then Exit Function

    rng.Value = IIf(UCase$(Trim$(rng.Value)) = UCase$(valA), valB, valA)
    Toggle = True
End Function

Private Function TogglePredMode(ByVal Target As Range, ByVal ws As Worksheet) As Boolean
    ' Special toggle for Pred Mode that also refreshes display
    Dim rng As Range, newMode As String
    On Error Resume Next
    Set rng = ws.Range(Schema.NAME_PRED_MODE)
    On Error GoTo 0
    If rng Is Nothing Or Intersect(Target, rng) Is Nothing Then Exit Function

    newMode = IIf(UCase$(Trim$(rng.Value)) = "STANDARD", "Enhanced", "Standard")
    rng.Value = newMode
    Data.RefreshPredictedRow newMode
    TogglePredMode = True
End Function

Private Sub ToggleCell(ByVal cell As Range, ByVal valA As String, ByVal valB As String)
    cell.Value = IIf(UCase$(Trim$(cell.Value)) = UCase$(valA), valB, valA)
End Sub

Private Function ValidateDateEntry(ByVal Target As Range) As Boolean
    Dim v As Variant
    v = Target.Value
    If IsEmpty(v) Or Len(Trim$(CStr(v))) = 0 Then ValidateDateEntry = True: Exit Function
    If IsDate(v) Then ValidateDateEntry = True: Exit Function

    Application.EnableEvents = False
    Target.ClearContents
    Application.EnableEvents = True
    MsgBox "Please enter a valid date.", vbExclamation, "WQOC"
End Function
