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
    PopulateIRFromCatalog site
    PopulateRRLatest site

    Application.EnableEvents = True
    Application.ScreenUpdating = True
End Sub

' ==== IR Table Population ===================================================

Private Sub ClearIRTable()
    Dim tbl As ListObject
    Set tbl = Schema.GetTable(Schema.SHEET_INPUT, Schema.TABLE_IR)
    If tbl Is Nothing Then Exit Sub
    If Not tbl.DataBodyRange Is Nothing Then
        tbl.DataBodyRange.Delete
    End If
End Sub

Private Sub PopulateIRFromCatalog(ByVal site As String)
    ' Reads tblCatalog, adds matching IR sites to tblIR, loads chemistry from tblResults
    Dim tblCat As ListObject, tblIR As ListObject
    Dim catRow As ListRow
    Dim irSite As String, flow As Double
    Dim chemNames As Variant, labData As Variant
    Dim i As Long

    Set tblCat = Schema.GetTable(Schema.SHEET_CONFIG, Schema.TABLE_CATALOG)
    Set tblIR = Schema.GetTable(Schema.SHEET_INPUT, Schema.TABLE_IR)
    If tblCat Is Nothing Or tblIR Is Nothing Then Exit Sub

    chemNames = Schema.ChemistryNames()

    For Each catRow In tblCat.ListRows
        If Schema.MatchesSite(catRow.Range.Cells(1, 1).Value, site) Then
            irSite = Trim$(catRow.Range.Cells(1, 2).Value)
            flow = Val(catRow.Range.Cells(1, 3).Value)

            ' Add row to IR table
            tblIR.ListRows.Add
            With tblIR.ListRows(tblIR.ListRows.Count).Range
                .Cells(1, Schema.ColIdx(tblIR, Schema.IR_COL_SOURCE)) = irSite
                .Cells(1, Schema.ColIdx(tblIR, Schema.IR_COL_FLOW)) = flow
                .Cells(1, Schema.ColIdx(tblIR, Schema.IR_COL_ACTIVE)) = "Yes"

                ' Load latest chemistry from Results (if exists)
                labData = GetLatestLabData(irSite)
                If Not IsEmpty(labData) Then
                    .Cells(1, Schema.ColIdx(tblIR, Schema.IR_COL_SAMPLE_DATE)) = labData(0)
                    For i = 0 To UBound(chemNames)
                        .Cells(1, Schema.ColIdx(tblIR, chemNames(i))) = labData(i + 1)
                    Next i
                End If
            End With
            Schema.InitIRRowAction tblIR.ListRows(tblIR.ListRows.Count).Range, tblIR
        End If
    Next catRow
End Sub

' ==== RR Latest Population ==================================================

Private Sub PopulateRRLatest(ByVal site As String)
    ' Loads latest RR chemistry from tblResults into Latest row (Row 3)
    Dim ws As Worksheet, labData As Variant
    Dim chemNames As Variant, rng As Range
    Dim i As Long

    Set ws = Schema.GetSheet(Schema.SHEET_INPUT)
    If ws Is Nothing Then Exit Sub

    labData = GetLatestLabData(site)
    If IsEmpty(labData) Then Exit Sub

    chemNames = Schema.ChemistryNames()

    ' Write sample date
    On Error Resume Next
    ws.Range(Schema.NAME_SAMPLE_DATE).Value = labData(0)
    On Error GoTo 0

    ' Write chemistry to RES_ROW (C3:I3)
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

' ==== Results Table Query ===================================================

Private Function GetLatestLabData(ByVal site As String) As Variant
    ' Returns array: (SampleDate, Chem1..Chem7) or Empty if not found
    ' Finds most recent sample for given site in tblResults
    Dim tbl As ListObject, row As ListRow
    Dim latestDate As Date, latestRow As ListRow
    Dim chemNames As Variant, result() As Variant
    Dim sampleDate As Date, i As Long

    Set tbl = Schema.GetTable(Schema.SHEET_RESULTS, Schema.TABLE_RESULTS)
    If tbl Is Nothing Then Exit Function
    If tbl.ListRows.Count = 0 Then Exit Function

    chemNames = Schema.ChemistryNames()

    ' Find most recent sample for this site
    latestDate = 0
    For Each row In tbl.ListRows
        If Schema.MatchesSite(row.Range.Cells(1, 1).Value, site) Then
            On Error Resume Next
            sampleDate = CDate(row.Range.Cells(1, 2).Value)
            On Error GoTo 0
            If sampleDate > latestDate Then
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
        result(i + 1) = Val(latestRow.Range.Cells(1, Schema.ColIdx(tbl, chemNames(i))).Value)
    Next i

    GetLatestLabData = result
End Function

