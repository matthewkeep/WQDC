Option Explicit
' RRState: Per-site settings persistence.
' Dependencies: Helpers, Schema, Core

' ==== Public Entry Points ===================================================

Public Sub Save(ByVal site As String)
    ' Saves current Inputs sheet state to tblRRState for site (9 bundled columns)
    Dim ws As Worksheet, tbl As ListObject, row As ListRow
    Dim n As Long, trigVol As Double

    On Error GoTo Fail
    If Len(Trim$(site)) = 0 Then Exit Sub

    Set ws = Helpers.GetSheet(Schema.SHEET_INPUT)
    Set tbl = Helpers.GetTable(Schema.SHEET_CONFIG, Schema.TABLE_RRSTATE)
    If ws Is Nothing Or tbl Is Nothing Then Exit Sub

    n = Core.METRIC_COUNT
    Set row = FindOrCreateRow(tbl, site)
    If row Is Nothing Then Exit Sub

    ' Read current state from Inputs and write to row
    With row.Range
        .Cells(1, Helpers.ColIdx(tbl, Schema.RRSTATE_COL_SAMPLE_DATE)).Value = _
            Helpers.GetDateVal(ws, Schema.NAME_SAMPLE_DATE)

        ' ResChemistry
        .Cells(1, Helpers.ColIdx(tbl, Schema.RRSTATE_COL_RES_CHEM)).Value = _
            Helpers.SerializeRange(Helpers.GetRng(ws, Schema.NAME_RES_ROW), n)

        ' IR table snapshot
        Dim tblIR As ListObject
        Set tblIR = Helpers.GetTable(Schema.SHEET_INPUT, Schema.TABLE_IR)
        If Not tblIR Is Nothing Then
            .Cells(1, Helpers.ColIdx(tbl, Schema.RRSTATE_COL_IR_SNAPSHOT)).Value = _
                Helpers.SerializeIRTable(tblIR)
        End If

        ' Triggers: Vol|EC|F_U|F_Mn|SO4|Mg|Ca|TAN
        trigVol = Val(Helpers.ReadFromRange(ws, Schema.NAME_TRIGGER_VOL))
        .Cells(1, Helpers.ColIdx(tbl, Schema.RRSTATE_COL_TRIGGERS)).Value = _
            Helpers.SerializeTriggers(trigVol, Helpers.GetRng(ws, Schema.NAME_LIMIT_ROW))

        ' EnhSettings: Enabled|TelemCal|RainfallMode|RainFactor|MixingModel|Tau|SurfaceFrac
        .Cells(1, Helpers.ColIdx(tbl, Schema.RRSTATE_COL_ENH_SETTINGS)).Value = _
            Helpers.SerializeEnhSettingsState(ws)

        ' HiddenMass
        .Cells(1, Helpers.ColIdx(tbl, Schema.RRSTATE_COL_HIDDEN_MASS)).Value = _
            Helpers.SerializeColumn(Helpers.GetRng(ws, Schema.NAME_HIDDEN_MASS), n)

        ' SignName
        .Cells(1, Helpers.ColIdx(tbl, Schema.RRSTATE_COL_SIGN_NAME)).Value = _
            Helpers.ReadFromRange(ws, Schema.NAME_SIGN_OFF_NAME)

        .Cells(1, Helpers.ColIdx(tbl, Schema.RRSTATE_COL_LAST_MODIFIED)).Value = Now
    End With

    ' Prevent wrap text (IRSnapshot contains newlines)
    row.Range.WrapText = False
    Exit Sub
Fail:
    Error.TraceErr "RRState.Save"
End Sub

Public Function Load(ByVal site As String) As Boolean
    ' Loads saved state from tblRRState to Inputs sheet (if exists)
    ' Returns True if state was loaded (9 bundled columns)
    Dim ws As Worksheet, tbl As ListObject, row As ListRow
    Dim n As Long, rowIdx As Long
    Dim sampleDate As Variant, trigVol As Double
    Dim irSnapshot As String, tblIR As ListObject

    If Len(Trim$(site)) = 0 Then Exit Function

    Set ws = Helpers.GetSheet(Schema.SHEET_INPUT)
    Set tbl = Helpers.GetTable(Schema.SHEET_CONFIG, Schema.TABLE_RRSTATE)
    If ws Is Nothing Or tbl Is Nothing Then Exit Function
    If tbl.DataBodyRange Is Nothing Then Exit Function

    rowIdx = FindRowIndex(tbl, site)
    If rowIdx = 0 Then Exit Function  ' No saved state for this site

    Set row = tbl.ListRows(rowIdx)
    n = Core.METRIC_COUNT

    On Error Resume Next
    Application.EnableEvents = False

    ' Restore settings to Inputs sheet
    With row.Range
        ' SampleDate
        sampleDate = .Cells(1, Helpers.ColIdx(tbl, Schema.RRSTATE_COL_SAMPLE_DATE)).Value
        If IsDate(sampleDate) And sampleDate > 0 Then
            Helpers.WriteToRange ws, Schema.NAME_SAMPLE_DATE, sampleDate
        End If

        ' ResChemistry
        Helpers.DeserializeToRange _
            CStr(.Cells(1, Helpers.ColIdx(tbl, Schema.RRSTATE_COL_RES_CHEM)).Value), _
            Helpers.GetRng(ws, Schema.NAME_RES_ROW), n

        ' IR table
        irSnapshot = CStr(.Cells(1, Helpers.ColIdx(tbl, Schema.RRSTATE_COL_IR_SNAPSHOT)).Value)
        If Len(irSnapshot) > 0 Then
            Set tblIR = Helpers.GetTable(Schema.SHEET_INPUT, Schema.TABLE_IR)
            If Not tblIR Is Nothing Then
                Helpers.DeserializeIRTable irSnapshot, tblIR
                Setup.ApplyIRActiveConditionalFormat tblIR
            End If
        End If

        ' Triggers: Vol|EC|F_U|F_Mn|SO4|Mg|Ca|TAN
        Helpers.DeserializeTriggers _
            CStr(.Cells(1, Helpers.ColIdx(tbl, Schema.RRSTATE_COL_TRIGGERS)).Value), _
            trigVol, Helpers.GetRng(ws, Schema.NAME_LIMIT_ROW)
        Helpers.WriteToRange ws, Schema.NAME_TRIGGER_VOL, trigVol

        ' EnhSettings: Enabled|TelemCal|RainfallMode|RainFactor|MixingModel|Tau|SurfaceFrac
        Helpers.DeserializeEnhSettingsState _
            CStr(.Cells(1, Helpers.ColIdx(tbl, Schema.RRSTATE_COL_ENH_SETTINGS)).Value), ws

        ' HiddenMass
        Helpers.DeserializeToColumn _
            CStr(.Cells(1, Helpers.ColIdx(tbl, Schema.RRSTATE_COL_HIDDEN_MASS)).Value), _
            Helpers.GetRng(ws, Schema.NAME_HIDDEN_MASS), n

        ' SignName
        Helpers.WriteToRange ws, Schema.NAME_SIGN_OFF_NAME, _
            .Cells(1, Helpers.ColIdx(tbl, Schema.RRSTATE_COL_SIGN_NAME)).Value
    End With

    Application.EnableEvents = True
    On Error GoTo 0

    Load = True
End Function

' ==== Private Helpers =======================================================

Private Function FindRowIndex(ByVal tbl As ListObject, ByVal site As String) As Long
    ' Returns row index (1-based) for site in tblRRState, or 0 if not found
    Dim i As Long, siteCol As Long
    If tbl Is Nothing Then Exit Function
    If tbl.DataBodyRange Is Nothing Then Exit Function

    siteCol = Helpers.ColIdx(tbl, Schema.RRSTATE_COL_SITE)
    If siteCol = 0 Then Exit Function

    For i = 1 To tbl.ListRows.Count
        If StrComp(Trim$(tbl.DataBodyRange.Cells(i, siteCol).Value), site, vbTextCompare) = 0 Then
            FindRowIndex = i
            Exit Function
        End If
    Next i
End Function

Private Function FindOrCreateRow(ByVal tbl As ListObject, ByVal site As String) As ListRow
    ' Finds existing row for site or creates a new one
    Dim rowIdx As Long, siteCol As Long

    rowIdx = FindRowIndex(tbl, site)
    If rowIdx > 0 Then
        Set FindOrCreateRow = tbl.ListRows(rowIdx)
        Exit Function
    End If

    ' Create new row
    Set FindOrCreateRow = tbl.ListRows.Add
    siteCol = Helpers.ColIdx(tbl, Schema.RRSTATE_COL_SITE)
    If siteCol > 0 Then
        FindOrCreateRow.Range.Cells(1, siteCol).Value = site
    End If
End Function
