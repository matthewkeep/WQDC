Option Explicit
' Loader: Site selection and data population.
' Dependencies: Schema

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
    ' Loads trigger values from tblTriggers into B4:I4
    Dim tbl As ListObject, ws As Worksheet
    Dim row As ListRow, chemNames As Variant
    Dim i As Long

    If Len(Trim$(preset)) = 0 Then Exit Sub

    Set tbl = Helpers.GetTable(Schema.SHEET_CONFIG, Schema.TABLE_TRIGGERS)
    Set ws = Helpers.GetSheet(Schema.SHEET_INPUT)
    If tbl Is Nothing Or ws Is Nothing Then Exit Sub
    If tbl.DataBodyRange Is Nothing Then Exit Sub

    Application.EnableEvents = False

    ' Find preset row
    For Each row In tbl.ListRows
        If StrComp(Trim$(row.Range.Cells(1, 1).Value), preset, vbTextCompare) = 0 Then
            ws.Range(Schema.NAME_TRIGGER_VOL).Value = row.Range.Cells(1, 2).Value
            chemNames = Schema.ChemistryNames()
            For i = 0 To UBound(chemNames)
                ws.Range(Schema.NAME_LIMIT_ROW).Cells(1, i + 1).Value = row.Range.Cells(1, 3 + i).Value
            Next i
            Application.EnableEvents = True
            Exit Sub
        End If
    Next row

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

    Set tblIdx = Helpers.GetTable(Schema.SHEET_CONFIG, Schema.TABLE_INDEX)
    Set tblIR = Helpers.GetTable(Schema.SHEET_INPUT, Schema.TABLE_IR)
    If tblIdx Is Nothing Or tblIR Is Nothing Then Exit Sub

    For Each idxRow In tblIdx.ListRows
        If Helpers.MatchesSite(idxRow.Range.Cells(1, 1).Value, site) Then
            irSite = Trim$(idxRow.Range.Cells(1, 2).Value)
            flow = Val(idxRow.Range.Cells(1, 3).Value)
            AddIRRowWithChemistry tblIR, irSite, flow, 0
        End If
    Next idxRow

    Setup.ApplyIRActiveConditionalFormat tblIR
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

    Set tbl = Helpers.GetTable(Schema.SHEET_RESULTS, Schema.TABLE_RESULTS)
    If tbl Is Nothing Then Exit Function
    If tbl.ListRows.Count = 0 Then Exit Function

    chemNames = Schema.ChemistryNames()
    maxDate = IIf(cutoffDate = 0, DateSerial(9999, 12, 31), cutoffDate)

    ' Find most recent sample for this site
    latestDate = 0
    For Each row In tbl.ListRows
        If Helpers.MatchesSite(row.Range.Cells(1, 1).Value, site) Then
            On Error Resume Next
            sampleDate = CDate(row.Range.Cells(1, 2).Value)
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
        result(i + 1) = Val(latestRow.Range.Cells(1, Helpers.ColIdx(tbl, chemNames(i))).Value)
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

    Set tblIdx = Helpers.GetTable(Schema.SHEET_CONFIG, Schema.TABLE_INDEX)
    Set tblIR = Helpers.GetTable(Schema.SHEET_INPUT, Schema.TABLE_IR)
    If tblIdx Is Nothing Or tblIR Is Nothing Then Exit Sub

    wasEmpty = (tblIR.DataBodyRange Is Nothing)

    ' Update existing IR rows with filtered chemistry
    If Not wasEmpty Then UpdateIRChemistry tblIR, cutoffDate

    ' Add missing IRs from catalog
    For Each idxRow In tblIdx.ListRows
        If Helpers.MatchesSite(idxRow.Range.Cells(1, 1).Value, site) Then
            irSite = Trim$(idxRow.Range.Cells(1, 2).Value)
            If Not IRExistsInTable(tblIR, irSite) Then
                flow = Val(idxRow.Range.Cells(1, 3).Value)
                AddIRRowWithChemistry tblIR, irSite, flow, cutoffDate
            End If
        End If
    Next idxRow

    If wasEmpty Then Setup.ApplyIRActiveConditionalFormat tblIR
End Sub

Private Sub AddIRRowWithChemistry(ByVal tbl As ListObject, ByVal irSite As String, _
                                   ByVal flow As Double, Optional ByVal cutoffDate As Date = 0)
    ' Add IR row and load chemistry (optionally filtered by date)
    Dim labData As Variant, chemNames As Variant, i As Long

    tbl.ListRows.Add
    With tbl.ListRows(tbl.ListRows.Count).Range
        .Cells(1, Helpers.ColIdx(tbl, Schema.IR_COL_SOURCE)) = irSite
        .Cells(1, Helpers.ColIdx(tbl, Schema.IR_COL_FLOW)) = flow
        .Cells(1, Helpers.ColIdx(tbl, Schema.IR_COL_ACTIVE)) = "Yes"

        labData = GetLatestLabData(irSite, cutoffDate)
        If Not IsEmpty(labData) Then
            chemNames = Schema.ChemistryNames()
            .Cells(1, Helpers.ColIdx(tbl, Schema.IR_COL_SAMPLE_DATE)) = labData(0)
            For i = 0 To UBound(chemNames)
                .Cells(1, Helpers.ColIdx(tbl, chemNames(i))) = labData(i + 1)
            Next i
        End If
    End With
    Helpers.InitIRRowAction tbl.ListRows(tbl.ListRows.Count).Range, tbl
End Sub

Private Sub UpdateIRChemistry(ByVal tbl As ListObject, ByVal cutoffDate As Date)
    ' Update chemistry for all existing IR rows
    Dim labData As Variant, chemNames As Variant
    Dim irSite As String, i As Long, rowIdx As Long

    chemNames = Schema.ChemistryNames()
    For rowIdx = 1 To tbl.ListRows.Count
        irSite = Trim$(tbl.DataBodyRange.Cells(rowIdx, Helpers.ColIdx(tbl, Schema.IR_COL_SOURCE)).Value)
        If Len(irSite) > 0 Then
            labData = GetLatestLabData(irSite, cutoffDate)
            If Not IsEmpty(labData) Then
                tbl.DataBodyRange.Cells(rowIdx, Helpers.ColIdx(tbl, Schema.IR_COL_SAMPLE_DATE)) = labData(0)
                For i = 0 To UBound(chemNames)
                    tbl.DataBodyRange.Cells(rowIdx, Helpers.ColIdx(tbl, chemNames(i))) = labData(i + 1)
                Next i
            End If
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

