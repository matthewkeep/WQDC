Option Explicit
' Loader: Site selection and data population.
' Dependencies: Schema, Helpers, Setup, DictionaryShim

' ==== Public Entry Points ===================================================

Public Sub LoadSiteData(ByVal site As String)
    ' Main orchestrator: clears IR, loads from catalog, loads RR latest
    If Len(Trim$(site)) = 0 Then Exit Sub
    Application.ScreenUpdating = False
    Application.EnableEvents = False

    ClearIRTable
    PopulateIRFromIndex site
    PopulateRRLatest site

    Application.EnableEvents = True
    Application.ScreenUpdating = True
End Sub

Public Sub LoadTriggerPreset(ByVal preset As String)
    ' Loads trigger values from tblTriggers into Limit row
    Dim tbl As ListObject, ws As Worksheet
    Dim row As ListRow, chemNames As Variant
    Dim i As Long
    Dim presetCol As Long, volCol As Long

    If Len(Trim$(preset)) = 0 Then Exit Sub

    Set tbl = Helpers.GetTable(Schema.SHEET_CONFIG, Schema.TABLE_TRIGGERS)
    Set ws = Helpers.GetSheet(Schema.SHEET_INPUT)
    If tbl Is Nothing Or ws Is Nothing Then Exit Sub
    If Not Helpers.HasData(tbl) Then Exit Sub

    presetCol = Helpers.ColIdx(tbl, Schema.TRIGGERS_COL_PRESET)
    volCol = Helpers.ColIdx(tbl, Schema.TRIGGERS_COL_VOL)
    chemNames = Schema.ChemistryNames()

    Application.EnableEvents = False

    ' Find preset row
    For Each row In tbl.ListRows
        If StrComp(Trim$(row.Range.Cells(1, presetCol).Value), preset, vbTextCompare) = 0 Then
            ws.Range(Schema.NAME_TRIGGER_VOL).Value = row.Range.Cells(1, volCol).Value
            For i = 0 To UBound(chemNames)
                ws.Range(Schema.NAME_LIMIT_ROW).Cells(1, i + 1).Value = _
                    row.Range.Cells(1, Helpers.ColIdx(tbl, chemNames(i))).Value
            Next i
            Application.EnableEvents = True
            Exit Sub
        End If
    Next row

    Application.EnableEvents = True
End Sub

' ==== IR Source Dropdown Support =============================================

Public Function GetResultsSources() As String
    ' Returns comma-separated unique Site values from tblResults for dropdown
    Dim tbl As ListObject, row As ListRow, dict As Object, site As String
    Dim siteCol As Long

    Set tbl = Helpers.GetTable(Schema.SHEET_RESULTS, Schema.TABLE_RESULTS)
    If Not Helpers.HasData(tbl) Then Exit Function

    siteCol = Helpers.ColIdx(tbl, Schema.RESULTS_COL_SITE)
    Set dict = New DictionaryShim
    For Each row In tbl.ListRows
        site = Trim$(CStr(row.Range.Cells(1, siteCol).Value))
        If Len(site) > 0 And Not dict.Exists(site) Then dict.Add site, 1
    Next row

    GetResultsSources = Join(dict.Keys, ",")
End Function

Public Function GetIndexFlow(ByVal irSite As String) As Double
    ' Returns Flow value from tblIndex for given IR, or 0 if not found
    Dim tbl As ListObject, row As ListRow, ir As String
    Dim irCol As Long, flowCol As Long

    Set tbl = Helpers.GetTable(Schema.SHEET_CONFIG, Schema.TABLE_INDEX)
    If Not Helpers.HasData(tbl) Then Exit Function

    irCol = Helpers.ColIdx(tbl, Schema.INDEX_COL_IR)
    flowCol = Helpers.ColIdx(tbl, Schema.INDEX_COL_FLOW)

    For Each row In tbl.ListRows
        ir = Trim$(CStr(row.Range.Cells(1, irCol).Value))
        If StrComp(ir, irSite, vbTextCompare) = 0 Then
            GetIndexFlow = Val(row.Range.Cells(1, flowCol).Value)
            Exit Function
        End If
    Next row
End Function

Public Sub RefreshIRRow(ByVal sourceCell As Range)
    ' Loads chemistry and flow for selected IR source
    ' Called when user selects source from dropdown
    Dim tbl As ListObject, ws As Worksheet, rowRng As Range
    Dim source As String, runDate As Date, flow As Double

    Set tbl = Helpers.GetTable(Schema.SHEET_INPUT, Schema.TABLE_IR)
    Set ws = Helpers.GetSheet(Schema.SHEET_INPUT)
    If tbl Is Nothing Or ws Is Nothing Then Exit Sub

    source = Trim$(CStr(sourceCell.Value))
    If Len(source) = 0 Then Exit Sub

    Set rowRng = Intersect(sourceCell.EntireRow, tbl.DataBodyRange)
    If rowRng Is Nothing Then Exit Sub

    runDate = Helpers.GetDateVal(ws, Schema.NAME_RUN_DATE)
    If runDate = 0 Then runDate = Date

    Application.EnableEvents = False

    ' Load flow from tblIndex (if available)
    flow = GetIndexFlow(source)
    If flow > 0 Then
        rowRng.Cells(1, Helpers.ColIdx(tbl, Schema.IR_COL_FLOW)).Value = flow
    End If

    ' Load chemistry from tblResults
    WriteIRRowChemistry rowRng, tbl, source, runDate, True

    Application.EnableEvents = True
End Sub

' ==== IR Table Population ===================================================

Private Sub ClearIRTable()
    Dim tbl As ListObject
    Set tbl = Helpers.GetTable(Schema.SHEET_INPUT, Schema.TABLE_IR)
    If tbl Is Nothing Then Exit Sub
    If Not tbl.DataBodyRange Is Nothing Then
        tbl.DataBodyRange.Delete
    End If
End Sub

Private Sub PopulateIRFromIndex(ByVal site As String)
    ' Reads tblIndex, adds matching IR sites to tblIR, loads chemistry from tblResults
    Dim tblIdx As ListObject, tblIR As ListObject
    Dim idxRow As ListRow, irSite As String, flow As Double
    Dim siteCol As Long, irCol As Long, flowCol As Long

    Set tblIdx = Helpers.GetTable(Schema.SHEET_CONFIG, Schema.TABLE_INDEX)
    Set tblIR = Helpers.GetTable(Schema.SHEET_INPUT, Schema.TABLE_IR)
    If tblIdx Is Nothing Or tblIR Is Nothing Then Exit Sub
    If Not Helpers.HasData(tblIdx) Then Exit Sub

    siteCol = Helpers.ColIdx(tblIdx, Schema.INDEX_COL_SITE)
    irCol = Helpers.ColIdx(tblIdx, Schema.INDEX_COL_IR)
    flowCol = Helpers.ColIdx(tblIdx, Schema.INDEX_COL_FLOW)

    For Each idxRow In tblIdx.ListRows
        If Helpers.MatchesSite(idxRow.Range.Cells(1, siteCol).Value, site) Then
            irSite = Trim$(idxRow.Range.Cells(1, irCol).Value)
            flow = Val(idxRow.Range.Cells(1, flowCol).Value)
            AddIRRow tblIR, irSite, flow, 0
        End If
    Next idxRow

    Setup.ApplyIRActiveConditionalFormat tblIR
    Setup.ApplyIRSourceDropdown tblIR
End Sub

' ==== Results Table Query ===================================================

Private Function GetLatestLabData(ByVal site As String, Optional ByVal cutoffDate As Date = 0) As Variant
    ' Returns array: (SampleDate, Chem1..Chem7) or Empty if not found
    ' Finds most recent sample for given site in tblResults
    ' If cutoffDate provided, only considers samples where date <= cutoffDate
    Dim tbl As ListObject, row As ListRow
    Dim latestDate As Date, latestRow As ListRow
    Dim chemNames As Variant, result() As Variant
    Dim sampleDate As Date, i As Long
    Dim maxDate As Date
    Dim siteCol As Long, dateCol As Long, chemCols() As Long

    Set tbl = Helpers.GetTable(Schema.SHEET_RESULTS, Schema.TABLE_RESULTS)
    If Not Helpers.HasData(tbl) Then Exit Function

    ' Pre-fetch column indices
    chemNames = Schema.ChemistryNames()
    siteCol = Helpers.ColIdx(tbl, Schema.RESULTS_COL_SITE)
    dateCol = Helpers.ColIdx(tbl, Schema.RESULTS_COL_DATE)
    ReDim chemCols(0 To UBound(chemNames))
    For i = 0 To UBound(chemNames)
        chemCols(i) = Helpers.ColIdx(tbl, chemNames(i))
    Next i

    maxDate = IIf(cutoffDate = 0, DateSerial(9999, 12, 31), cutoffDate)

    ' Find most recent sample for this site
    latestDate = 0
    For Each row In tbl.ListRows
        If Helpers.MatchesSite(row.Range.Cells(1, siteCol).Value, site) Then
            On Error Resume Next
            sampleDate = CDate(row.Range.Cells(1, dateCol).Value)
            On Error GoTo 0
            If sampleDate <= maxDate And sampleDate > latestDate Then
                latestDate = sampleDate
                Set latestRow = row
            End If
        End If
    Next row

    If latestRow Is Nothing Then Exit Function

    ' Build result array: (Date, Chem1..Chem7)
    ReDim result(0 To UBound(chemNames) + 1)
    result(0) = latestDate
    For i = 0 To UBound(chemNames)
        result(i + 1) = Val(latestRow.Range.Cells(1, chemCols(i)).Value)
    Next i

    GetLatestLabData = result
End Function

' ==== Load Latest (Date-Filtered) ==============================================

Public Sub LoadLatest()
    ' Loads latest chemistry for RR and IRs where sample date <= run date
    Dim ws As Worksheet, site As String, runDate As Date

    Set ws = Helpers.GetSheet(Schema.SHEET_INPUT)
    If ws Is Nothing Then Exit Sub

    On Error Resume Next
    site = CStr(ws.Range(Schema.NAME_SITE).Value)
    runDate = CDate(ws.Range(Schema.NAME_RUN_DATE).Value)
    On Error GoTo 0

    If Len(Trim$(site)) = 0 Or runDate = 0 Then
        MsgBox "Please set Site and Run Date first.", vbExclamation, "WQOC"
        Exit Sub
    End If

    Application.ScreenUpdating = False
    Application.EnableEvents = False

    PopulateRRLatest site, runDate
    PopulateIRLatestFiltered site, runDate

    Application.EnableEvents = True
    Application.ScreenUpdating = True
End Sub

Public Sub LoadLatestInternal(ByVal site As String, ByVal cutoffDate As Date)
    ' Internal version for Backtest: loads RR and catalog IRs filtered by date
    ' Does not show UI messages, does not toggle screen updating
    ' Used to simulate historical data availability during season replay
    PopulateRRLatest site, cutoffDate
    PopulateIRLatestFiltered site, cutoffDate
End Sub

Private Sub PopulateRRLatest(ByVal site As String, Optional ByVal cutoffDate As Date = 0)
    ' Loads latest RR chemistry from tblResults (optionally filtered by cutoffDate)
    Dim ws As Worksheet, labData As Variant
    Dim chemNames As Variant, rng As Range, i As Long

    Set ws = Helpers.GetSheet(Schema.SHEET_INPUT)
    If ws Is Nothing Then Exit Sub

    labData = GetLatestLabData(site, cutoffDate)
    If IsEmpty(labData) Then Exit Sub

    chemNames = Schema.ChemistryNames()

    ' Write sample date
    On Error Resume Next
    ws.Range(Schema.NAME_SAMPLE_DATE).Value = labData(0)
    On Error GoTo 0

    ' Write chemistry to RES_ROW
    Set rng = Nothing
    On Error Resume Next
    Set rng = ws.Range(Schema.NAME_RES_ROW)
    On Error GoTo 0

    If Not rng Is Nothing Then
        For i = 0 To UBound(chemNames)
            If i < rng.Columns.Count Then
                rng.Cells(1, i + 1).Value = labData(i + 1)
            End If
        Next i
    End If
End Sub

Private Sub PopulateIRLatestFiltered(ByVal site As String, ByVal cutoffDate As Date)
    ' Updates IR table chemistry from tblResults where sample date <= cutoffDate
    ' Preserves existing IRs (including manual entries not in catalog)
    ' Adds missing IRs from catalog
    Dim tblIdx As ListObject, tblIR As ListObject
    Dim idxRow As ListRow, irSite As String, flow As Double
    Dim wasEmpty As Boolean
    Dim siteCol As Long, irCol As Long, flowCol As Long

    Set tblIdx = Helpers.GetTable(Schema.SHEET_CONFIG, Schema.TABLE_INDEX)
    Set tblIR = Helpers.GetTable(Schema.SHEET_INPUT, Schema.TABLE_IR)
    If tblIdx Is Nothing Or tblIR Is Nothing Then Exit Sub

    wasEmpty = Not Helpers.HasData(tblIR)

    ' Update existing IR rows with filtered chemistry
    If Not wasEmpty Then RefreshAllIRRows tblIR, cutoffDate

    ' Add missing IRs from catalog
    If Helpers.HasData(tblIdx) Then
        siteCol = Helpers.ColIdx(tblIdx, Schema.INDEX_COL_SITE)
        irCol = Helpers.ColIdx(tblIdx, Schema.INDEX_COL_IR)
        flowCol = Helpers.ColIdx(tblIdx, Schema.INDEX_COL_FLOW)

        For Each idxRow In tblIdx.ListRows
            If Helpers.MatchesSite(idxRow.Range.Cells(1, siteCol).Value, site) Then
                irSite = Trim$(idxRow.Range.Cells(1, irCol).Value)
                If Not IRExistsInTable(tblIR, irSite) Then
                    flow = Val(idxRow.Range.Cells(1, flowCol).Value)
                    AddIRRow tblIR, irSite, flow, cutoffDate
                End If
            End If
        Next idxRow
    End If

    If wasEmpty Then
        Setup.ApplyIRActiveConditionalFormat tblIR
        Setup.ApplyIRSourceDropdown tblIR
    End If
End Sub

Private Sub AddIRRow(ByVal tbl As ListObject, ByVal irSite As String, _
                                   ByVal flow As Double, Optional ByVal cutoffDate As Date = 0)
    ' Add IR row and load chemistry (optionally filtered by date)
    tbl.ListRows.Add
    With tbl.ListRows(tbl.ListRows.Count).Range
        .Cells(1, Helpers.ColIdx(tbl, Schema.IR_COL_SOURCE)) = irSite
        .Cells(1, Helpers.ColIdx(tbl, Schema.IR_COL_FLOW)) = flow
        .Cells(1, Helpers.ColIdx(tbl, Schema.IR_COL_ACTIVE)) = "Yes"
        WriteIRRowChemistry tbl.ListRows(tbl.ListRows.Count).Range, tbl, irSite, cutoffDate, False
    End With
    Helpers.InitIRRowAction tbl.ListRows(tbl.ListRows.Count).Range, tbl
End Sub

Private Sub WriteIRRowChemistry(ByVal rowRng As Range, ByVal tbl As ListObject, _
                                 ByVal irSite As String, ByVal cutoffDate As Date, _
                                 ByVal clearIfEmpty As Boolean)
    ' Writes chemistry from Results, optionally clears if no data found
    ' Flow is handled separately by caller (single responsibility)
    Dim labData As Variant, chemNames As Variant, i As Long
    Dim dateCol As Long, chemCols() As Long

    labData = GetLatestLabData(irSite, cutoffDate)
    chemNames = Schema.ChemistryNames()

    ' Pre-fetch column indices
    dateCol = Helpers.ColIdx(tbl, Schema.IR_COL_SAMPLE_DATE)
    ReDim chemCols(0 To UBound(chemNames))
    For i = 0 To UBound(chemNames)
        chemCols(i) = Helpers.ColIdx(tbl, chemNames(i))
    Next i

    If Not IsEmpty(labData) Then
        rowRng.Cells(1, dateCol) = labData(0)
        For i = 0 To UBound(chemNames)
            rowRng.Cells(1, chemCols(i)) = labData(i + 1)
        Next i
    ElseIf clearIfEmpty Then
        rowRng.Cells(1, dateCol).ClearContents
        For i = 0 To UBound(chemNames)
            rowRng.Cells(1, chemCols(i)).ClearContents
        Next i
    End If
End Sub

Private Sub RefreshAllIRRows(ByVal tbl As ListObject, ByVal cutoffDate As Date)
    ' Update chemistry for all existing IR rows
    Dim irSite As String, rowIdx As Long
    For rowIdx = 1 To tbl.ListRows.Count
        irSite = Trim$(tbl.DataBodyRange.Cells(rowIdx, Helpers.ColIdx(tbl, Schema.IR_COL_SOURCE)).Value)
        If Len(irSite) > 0 Then
            WriteIRRowChemistry tbl.ListRows(rowIdx).Range, tbl, irSite, cutoffDate, False
        End If
    Next rowIdx
End Sub

Private Function IRExistsInTable(ByVal tbl As ListObject, ByVal irSite As String) As Boolean
    Dim rowIdx As Long, srcCol As Long
    If tbl.DataBodyRange Is Nothing Then Exit Function
    srcCol = Helpers.ColIdx(tbl, Schema.IR_COL_SOURCE)
    For rowIdx = 1 To tbl.ListRows.Count
        If StrComp(Trim$(tbl.DataBodyRange.Cells(rowIdx, srcCol).Value), irSite, vbTextCompare) = 0 Then
            IRExistsInTable = True
            Exit Function
        End If
    Next rowIdx
End Function

