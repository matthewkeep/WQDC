Option Explicit
' Validate: Pre-flight workbook structure checks.
' Dependencies: Schema

Private mIssues As Collection

Public Function Check() As Boolean
    Set mIssues = New Collection
    ChkSheets
    ChkRanges
    ChkTables
    Check = (mIssues.Count = 0)
    If mIssues.Count > 0 Then
        Debug.Print "FAIL: " & mIssues.Count & " issue(s)"
    Else
        Debug.Print "PASS: Structure valid"
    End If
End Function

Public Sub Report()
    Dim i As Long
    If Not Check() Then
        Debug.Print ""
        For i = 1 To mIssues.Count
            Debug.Print "  " & i & ". " & mIssues(i)
        Next i
        Debug.Print ""
    End If
End Sub

Private Sub ChkSheets()
    ChkSheet Schema.SHEET_INPUT
    ChkSheet Schema.SHEET_CONFIG
    ChkSheet Schema.SHEET_RESULTS
    ChkSheet Schema.SHEET_TELEMETRY
    ChkSheet Schema.SHEET_HISTORY
    ChkSheet Schema.SHEET_LOG
    ChkSheet Schema.SHEET_CHART
End Sub

Private Sub ChkSheet(ByVal nm As String)
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(nm)
    On Error GoTo 0
    If ws Is Nothing Then mIssues.Add "Missing sheet: " & nm
End Sub

Private Sub ChkRanges()
    ChkRange Schema.NAME_SITE
    ChkRange Schema.NAME_INIT_VOL
    ChkRange Schema.NAME_TRIGGER_VOL
    ChkRange Schema.NAME_SAMPLE_DATE
    ChkRange Schema.NAME_RUN_DATE
    ChkRange Schema.NAME_OUTPUT
    ChkRange Schema.NAME_RES_ROW
    ChkRange Schema.NAME_LIMIT_ROW
    ChkRange Schema.NAME_PRED_ROW
    ChkRange Schema.NAME_HIDDEN_MASS
    ChkRange Schema.NAME_TAU
    ChkRange Schema.NAME_SURFACE_FRACTION
    ChkRange Schema.NAME_NET_OUT
    ChkRange Schema.NAME_ENHANCED_MODE
    ChkRange Schema.NAME_STD_TRIGGER
    ChkRange Schema.NAME_MIXING_MODEL
    ChkRange Schema.NAME_RAINFALL_MODE
    ChkRange Schema.NAME_TELEM_CAL
End Sub

Private Sub ChkRange(ByVal nm As String)
    Dim rng As Range
    On Error Resume Next
    Set rng = ThisWorkbook.Names(nm).RefersToRange
    On Error GoTo 0
    If rng Is Nothing Then mIssues.Add "Missing range: " & nm
End Sub

Private Sub ChkTables()
    ChkTable Schema.SHEET_INPUT, Schema.TABLE_IR
    ChkTable Schema.SHEET_CONFIG, Schema.TABLE_CATALOG
    ChkTable Schema.SHEET_CONFIG, Schema.TABLE_TRIGGER
    ChkTable Schema.SHEET_RESULTS, Schema.TABLE_RESULTS
    ChkTable Schema.SHEET_TELEMETRY, Schema.TABLE_TELEMETRY
    ' Note: Log and History tables are per-site, created on-demand
End Sub

Private Sub ChkTable(ByVal sht As String, ByVal tbl As String)
    Dim ws As Worksheet
    Dim lo As ListObject
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(sht)
    If Not ws Is Nothing Then Set lo = ws.ListObjects(tbl)
    On Error GoTo 0
    If lo Is Nothing Then mIssues.Add "Missing table: " & tbl
End Sub
