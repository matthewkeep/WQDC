Option Explicit
' Validate: Pre-flight workbook structure checks.
' Dependencies: Schema, Helpers

Private mIssues As Collection

' ==== Public ==================================================================

Public Function Check() As Boolean
    Set mIssues = New Collection
    ChkSheets
    ChkRanges
    ChkTables
    ChkDates
    Check = (mIssues.Count = 0)
    Debug.Print IIf(Check, "PASS: Structure valid", "FAIL: " & mIssues.Count & " issue(s)")
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

' ==== Private Implementation ==================================================

Private Sub ChkSheets()
    Dim v As Variant
    For Each v In Array(Schema.SHEET_INPUT, Schema.SHEET_CONFIG, Schema.SHEET_RESULTS, _
                        Schema.SHEET_RECORD, Schema.SHEET_LOG, Schema.SHEET_CHART)
        If Helpers.GetSheet(CStr(v)) Is Nothing Then mIssues.Add "Missing sheet: " & v
    Next v
End Sub

Private Sub ChkRanges()
    Dim v As Variant
    For Each v In Array(Schema.NAME_SITE, Schema.NAME_INIT_VOL, Schema.NAME_TRIGGER_VOL, _
                        Schema.NAME_SAMPLE_DATE, Schema.NAME_RUN_DATE, Schema.NAME_OUTPUT, _
                        Schema.NAME_RES_ROW, Schema.NAME_LIMIT_ROW, Schema.NAME_PRED_ROW, _
                        Schema.NAME_HIDDEN_MASS, Schema.NAME_TAU, Schema.NAME_SURFACE_FRACTION, _
                        Schema.NAME_ENHANCED_MODE, Schema.NAME_STD_TRIGGER, Schema.NAME_MIXING_MODEL, _
                        Schema.NAME_RAINFALL_MODE, Schema.NAME_TELEM_CAL, Schema.NAME_SIGN_OFF_NAME)
        If Not RangeExists(CStr(v)) Then mIssues.Add "Missing range: " & v
    Next v
End Sub

Private Sub ChkTables()
    ' Format: Array(sheet, table, sheet, table, ...)
    Dim items As Variant, i As Long
    items = Array(Schema.SHEET_INPUT, Schema.TABLE_IR, _
                  Schema.SHEET_CONFIG, Schema.TABLE_INDEX, _
                  Schema.SHEET_CONFIG, Schema.TABLE_TRIGGERS, _
                  Schema.SHEET_CONFIG, Schema.TABLE_USERS, _
                  Schema.SHEET_RESULTS, Schema.TABLE_RESULTS, _
                  Schema.SHEET_RESULTS, Schema.TABLE_TELEMETRY)
    For i = LBound(items) To UBound(items) Step 2
        If Helpers.GetTable(CStr(items(i)), CStr(items(i + 1))) Is Nothing Then
            mIssues.Add "Missing table: " & items(i + 1)
        End If
    Next i
    ' Note: Log and History tables are per-site, created on-demand
End Sub

Private Sub ChkDates()
    Dim rng As Range, v As Variant
    Dim items As Variant, i As Long
    ' Format: Array(rangeName, label, rangeName, label, ...)
    items = Array(Schema.NAME_RUN_DATE, "Run Date", Schema.NAME_SAMPLE_DATE, "Sample Date")
    For i = LBound(items) To UBound(items) Step 2
        Set rng = Nothing
        On Error Resume Next
        Set rng = ThisWorkbook.Names(items(i)).RefersToRange
        On Error GoTo 0
        If rng Is Nothing Then GoTo NextDate  ' Missing range caught by ChkRanges
        v = rng.Value
        If IsEmpty(v) Or Len(Trim$(CStr(v))) = 0 Then GoTo NextDate
        If Not IsDate(v) Then mIssues.Add "Invalid date in " & items(i + 1) & ": " & CStr(v)
NextDate:
    Next i
End Sub

Private Function RangeExists(ByVal nm As String) As Boolean
    Dim rng As Range
    On Error Resume Next
    Set rng = ThisWorkbook.Names(nm).RefersToRange
    On Error GoTo 0
    RangeExists = Not rng Is Nothing
End Function
