Option Explicit
' Setup: Workbook scaffolding and test data.
' Dependencies: Schema

Public Sub Build()
    Dim cm As XlCalculation
    On Error GoTo Fail
    cm = Application.Calculation
    Application.ScreenUpdating = False: Application.EnableEvents = False
    Application.Calculation = xlCalculationManual

    MakeSheets
    SetupInput
    SetupConfig
    SetupResults
    SetupTelemetry
    SetupChart
    SetupControls
    ' Note: Log and History tables are created on-demand per site

    Application.Calculation = cm: Application.ScreenUpdating = True: Application.EnableEvents = True
    MsgBox "Structure created.", vbInformation, "Setup"
    Exit Sub
Fail:
    Application.Calculation = cm: Application.ScreenUpdating = True: Application.EnableEvents = True
    MsgBox "Error: " & Err.Description, vbExclamation, "Setup"
End Sub

Public Sub Seed()
    Dim cm As XlCalculation
    On Error GoTo Fail
    cm = Application.Calculation
    Application.ScreenUpdating = False: Application.EnableEvents = False
    Application.Calculation = xlCalculationManual

    SeedInput
    SeedConfig
    SeedResults
    SeedTelemetry

    Application.Calculation = cm: Application.ScreenUpdating = True: Application.EnableEvents = True
    MsgBox "Test data seeded.", vbInformation, "Setup"
    Exit Sub
Fail:
    Application.Calculation = cm: Application.ScreenUpdating = True: Application.EnableEvents = True
    MsgBox "Error: " & Err.Description, vbExclamation, "Setup"
End Sub

Public Sub BuildAll(): Build: Seed: Initialize: End Sub

Public Sub Initialize()
    ' Reads all RR sites from tblCatalog and creates per-site infrastructure:
    ' - Telemetry columns (EC, Vol for each site)
    ' - Log tables (tblLog_{site})
    ' - History tables (tblHistory_{site})
    ' Safe to run multiple times - only creates what doesn't exist
    Dim cm As XlCalculation, sites As Variant, site As Variant
    Dim created As Long
    On Error GoTo Fail
    cm = Application.Calculation
    Application.ScreenUpdating = False: Application.EnableEvents = False
    Application.Calculation = xlCalculationManual

    sites = GetAllSites()
    If Not IsArray(sites) Then
        MsgBox "No sites found in Catalog. Add sites to tblCatalog first.", vbExclamation, "Initialize"
        GoTo Done
    End If

    created = 0
    For Each site In sites
        If Len(site) > 0 Then
            If EnsureSiteTelemColumns(CStr(site)) Then
                created = created + 1
                SeedSiteTelemetry CStr(site)  ' Seed sample data for new columns
            End If
            EnsureSiteLogTable CStr(site)
            EnsureSiteHistoryTable CStr(site)
        End If
    Next site

    Application.Calculation = cm: Application.ScreenUpdating = True: Application.EnableEvents = True
    MsgBox "Initialized " & UBound(sites) - LBound(sites) + 1 & " site(s).", vbInformation, "Initialize"
    Exit Sub
Done:
    Application.Calculation = cm: Application.ScreenUpdating = True: Application.EnableEvents = True
    Exit Sub
Fail:
    Application.Calculation = cm: Application.ScreenUpdating = True: Application.EnableEvents = True
    MsgBox "Error: " & Err.Description, vbExclamation, "Initialize"
End Sub

Private Function GetAllSites() As Variant
    ' Returns array of unique RR site names from first column of tblCatalog
    Dim ws As Worksheet, tbl As ListObject, row As ListRow
    Dim dict As Object, site As String

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(Schema.SHEET_CONFIG)
    If ws Is Nothing Then Exit Function
    Set tbl = ws.ListObjects(Schema.TABLE_CATALOG)
    On Error GoTo 0
    If tbl Is Nothing Then Exit Function
    If tbl.ListRows.Count = 0 Then Exit Function

    Set dict = New DictionaryShim
    For Each row In tbl.ListRows
        site = Trim$(CStr(row.Range.Cells(1, 1).Value))
        If Len(site) > 0 And Not dict.Exists(site) Then
            dict.Add site, True
        End If
    Next row

    If dict.Count = 0 Then Exit Function
    GetAllSites = dict.Keys
End Function

Private Function EnsureSiteTelemColumns(ByVal site As String) As Boolean
    ' Adds EC and Vol columns for site to telemetry table if they don't exist
    ' Returns True if columns were added
    Dim ws As Worksheet, tbl As ListObject
    Dim ecCol As String, volCol As String
    Dim addedAny As Boolean

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(Schema.SHEET_TELEMETRY)
    Set tbl = ws.ListObjects(Schema.TABLE_TELEMETRY)
    On Error GoTo 0
    If tbl Is Nothing Then Exit Function

    ecCol = Schema.TelemECColName(site)
    volCol = Schema.TelemVolColName(site)

    ' Add EC column if missing
    On Error Resume Next
    If tbl.ListColumns(ecCol) Is Nothing Then
        tbl.ListColumns.Add
        tbl.ListColumns(tbl.ListColumns.Count).Name = ecCol
        addedAny = True
    End If
    On Error GoTo 0

    ' Add Vol column if missing
    On Error Resume Next
    If tbl.ListColumns(volCol) Is Nothing Then
        tbl.ListColumns.Add
        tbl.ListColumns(tbl.ListColumns.Count).Name = volCol
        addedAny = True
    End If
    On Error GoTo 0

    EnsureSiteTelemColumns = addedAny
End Function

Public Sub Clean()
    Dim ws As Worksheet, nm As Name, sheets As Variant, i As Long
    sheets = Array(Schema.SHEET_INPUT, Schema.SHEET_CONFIG, Schema.SHEET_RESULTS, _
                   Schema.SHEET_TELEMETRY, Schema.SHEET_HISTORY, Schema.SHEET_CHART, Schema.SHEET_LOG)
    Application.DisplayAlerts = False
    For i = LBound(sheets) To UBound(sheets)
        On Error Resume Next
        Set ws = ThisWorkbook.Worksheets(sheets(i))
        If Not ws Is Nothing Then ws.Delete
        Set ws = Nothing
        On Error GoTo 0
    Next i
    Application.DisplayAlerts = True
    On Error Resume Next
    For Each nm In ThisWorkbook.Names: nm.Delete: Next nm
    On Error GoTo 0
    MsgBox "Cleaned.", vbInformation, "Setup"
End Sub

' ==== Sheet Creation =========================================================

Private Sub MakeSheets()
    MakeSheet Schema.SHEET_INPUT
    MakeSheet Schema.SHEET_CONFIG
    MakeSheet Schema.SHEET_RESULTS
    MakeSheet Schema.SHEET_TELEMETRY
    MakeSheet Schema.SHEET_HISTORY
    MakeSheet Schema.SHEET_CHART
    MakeSheet Schema.SHEET_LOG
End Sub

Private Sub MakeSheet(ByVal nm As String)
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(nm)
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        ws.Name = nm
    End If
    ws.Cells.Clear
End Sub

' ==== Input Sheet ============================================================

Private Sub SetupInput()
    Dim ws As Worksheet, chem As Variant, n As Long, i As Long
    Set ws = ThisWorkbook.Worksheets(Schema.SHEET_INPUT)
    chem = Schema.ChemistryNames(): n = Schema.ChemistryCount()

    ' Reservoir block
    ws.Range("A1") = "Reservoir": ws.Range("A1:L1").Font.Bold = True
    ws.Range("B2") = Schema.VOLUME_METRIC_NAME
    For i = 0 To n - 1: ws.Cells(2, 3 + i) = chem(i): Next i
    ws.Range("A3") = "Latest": ws.Range("A4") = "Trigger": ws.Range("A5") = "Predicted"
    AddNm Schema.NAME_INIT_VOL, ws.Range("B3")
    AddNm Schema.NAME_TRIGGER_VOL, ws.Range("B4")
    AddNm Schema.NAME_RESULT_VOL, ws.Range("B5")
    AddNm Schema.NAME_RES_ROW, ws.Range("C3").Resize(1, n)
    AddNm Schema.NAME_LIMIT_ROW, ws.Range("C4").Resize(1, n)

    ' Run info
    ws.Range("J2") = "Run Date": AddNm Schema.NAME_RUN_DATE, ws.Range("K2")
    ws.Range("J3") = "Site": AddNm Schema.NAME_SITE, ws.Range("K3")
    ws.Range("J4") = "Output": AddNm Schema.NAME_OUTPUT, ws.Range("K4")
    ws.Range("J5") = "Sample Date": AddNm Schema.NAME_SAMPLE_DATE, ws.Range("K5")

    ' Results
    ws.Range("N1") = "Results": ws.Range("N1:P1").Font.Bold = True
    ws.Range("N2") = "Std Trigger": AddNm Schema.NAME_STD_TRIGGER, ws.Range("O2")
    ws.Range("N3") = "Enh Trigger": AddNm Schema.NAME_ENH_TRIGGER, ws.Range("O3")
    ws.Range("N4") = "Mode": AddNm Schema.NAME_ENHANCED_MODE, ws.Range("O4")

    ' Calibration
    ws.Range("N6") = "Calibration": ws.Range("N6:O6").Font.Bold = True
    ws.Range("N7") = "Tau": AddNm Schema.NAME_TAU, ws.Range("O7")
    ws.Range("N8") = "Rain Factor": AddNm Schema.NAME_RAIN_FACTOR, ws.Range("O8")
    ws.Range("N9") = "Rain Mode": AddNm Schema.NAME_RAIN_MODE, ws.Range("O9")
    ws.Range("N10") = "Surface Frac": AddNm Schema.NAME_SURFACE_FRACTION, ws.Range("O10")
    ws.Range("N11") = "Net Outflow": AddNm Schema.NAME_NET_OUT, ws.Range("O11")

    ' Enhanced Config
    ws.Range("N13") = "Enhanced Config": ws.Range("N13:O13").Font.Bold = True
    ws.Range("N14") = "Enhanced Mode": AddNm Schema.NAME_ENHANCED_MODE, ws.Range("O14")
    ws.Range("N15") = "Mixing Model": AddNm Schema.NAME_MIXING_MODEL, ws.Range("O15")
    ws.Range("N16") = "Rainfall": AddNm Schema.NAME_RAINFALL_MODE, ws.Range("O16")
    ws.Range("N17") = "Telemetry Cal": AddNm Schema.NAME_TELEM_CAL, ws.Range("O17")

    ' Hidden mass
    ws.Range("Q6") = "Hidden Mass": ws.Range("Q6:R6").Font.Bold = True
    For i = 0 To n - 1: ws.Cells(7 + i, 17) = chem(i): Next i
    AddNm Schema.NAME_HIDDEN_MASS, ws.Range("R7").Resize(n, 1)

    ' IR Table
    ws.Range("A8") = "Inflow Sources": ws.Range("A8:L8").Font.Bold = True
    MakeIRTable ws, chem, n
End Sub

Private Sub MakeIRTable(ByVal ws As Worksheet, ByVal chem As Variant, ByVal n As Long)
    Dim h() As String, i As Long, tbl As ListObject
    ReDim h(1 To n + 5)
    h(1) = Schema.IR_COL_SOURCE: h(2) = Schema.IR_COL_FLOW
    For i = 0 To n - 1: h(3 + i) = chem(i): Next i
    h(n + 3) = Schema.IR_COL_SAMPLE_DATE: h(n + 4) = Schema.IR_COL_ACTIVE
    h(n + 5) = Schema.IR_COL_ACTION
    MakeTbl ws, ws.Range("A9"), Schema.TABLE_IR, h

    ' Set action column header text (no blue styling - table style has blue background)
    Set tbl = ws.ListObjects(Schema.TABLE_IR)
    If Not tbl Is Nothing Then
        tbl.ListColumns(Schema.IR_COL_ACTION).Name = Schema.ACTION_ADD
    End If
End Sub

' ==== Config Sheet ===========================================================

Private Sub SetupConfig()
    Dim ws As Worksheet, chem As Variant, n As Long, h() As String, i As Long
    Set ws = ThisWorkbook.Worksheets(Schema.SHEET_CONFIG)
    chem = Schema.ChemistryNames(): n = Schema.ChemistryCount()

    ws.Range("A1") = "Catalog": ws.Range("A1").Font.Bold = True
    MakeTbl ws, ws.Range("A2"), Schema.TABLE_CATALOG, Array("RR", "IR", "Flow")

    ws.Range("E1") = "Triggers": ws.Range("E1").Font.Bold = True
    ReDim h(1 To n + 2): h(1) = "Preset": h(2) = Schema.VOLUME_METRIC_NAME
    For i = 1 To n: h(2 + i) = chem(i - 1): Next i
    MakeTbl ws, ws.Range("E2"), Schema.TABLE_TRIGGER, h
End Sub

' ==== Results Sheet ==========================================================

Private Sub SetupResults()
    Dim ws As Worksheet, chem As Variant, n As Long, h() As String, i As Long
    Set ws = ThisWorkbook.Worksheets(Schema.SHEET_RESULTS)
    chem = Schema.ChemistryNames(): n = Schema.ChemistryCount()

    ws.Range("A1") = "Lab Results": ws.Range("A1").Font.Bold = True
    ReDim h(1 To n + 3): h(1) = "Site": h(2) = "Sample Date": h(3) = "Sample ID"
    For i = 1 To n: h(3 + i) = chem(i - 1): Next i
    MakeTbl ws, ws.Range("A2"), Schema.TABLE_RESULTS, h
End Sub

' ==== Telemetry Sheet ========================================================

Private Sub SetupTelemetry()
    ' Creates base telemetry table with Date and Rain columns only
    ' Per-site EC/Vol columns are added by Initialize
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(Schema.SHEET_TELEMETRY)
    ws.Range("A1") = "Telemetry Data": ws.Range("A1").Font.Bold = True
    ws.Range("A2") = "Daily observations - leave cells blank if data unavailable"
    ws.Range("A3") = "Run 'Initialize' after setting up Catalog to add site columns"
    MakeTbl ws, ws.Range("A5"), Schema.TABLE_TELEMETRY, _
        Array(Schema.TELEM_COL_DATE, Schema.TELEM_COL_RAIN)
End Sub

' ==== Test Data ==============================================================

Private Sub SeedInput()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(Schema.SHEET_INPUT)
    ws.Range(Schema.NAME_RUN_DATE) = Date
    ws.Range(Schema.NAME_SITE) = "RP1"
    ws.Range(Schema.NAME_OUTPUT) = 2
    ws.Range(Schema.NAME_SAMPLE_DATE) = Date - 5
    ws.Range(Schema.NAME_INIT_VOL) = 150
    ws.Range(Schema.NAME_TRIGGER_VOL) = 200
    ws.Range(Schema.NAME_RES_ROW) = Array(280, 45, 2.5, 15, 6, 10, 0.1)
    ws.Range(Schema.NAME_LIMIT_ROW) = Array(450, 100, 4, 22, 9, 15, 0.2)
    ws.Range(Schema.NAME_TAU) = 7
    ws.Range(Schema.NAME_RAIN_FACTOR) = 3.2
    ws.Range(Schema.NAME_RAIN_MODE) = "Typical"
    ws.Range(Schema.NAME_SURFACE_FRACTION) = 0.8
    ws.Range(Schema.NAME_NET_OUT) = 1

    ' Enhanced config defaults
    ws.Range(Schema.NAME_ENHANCED_MODE) = "On"
    ws.Range(Schema.NAME_MIXING_MODEL) = Schema.MIXING_TWOBUCKET
    ws.Range(Schema.NAME_RAINFALL_MODE) = Schema.RAINFALL_HINDCAST
    ws.Range(Schema.NAME_TELEM_CAL) = Schema.TELEM_CAL_OFF

    Dim tbl As ListObject
    Set tbl = ws.ListObjects(Schema.TABLE_IR)
    If Not tbl Is Nothing Then
        EnsureRows tbl, 2
        tbl.DataBodyRange.Rows(1) = Array("CB1", 1.5, 350, 60, 2.8, 16, 7, 11, 0.12, Date - 3, "Yes", Schema.ACTION_REMOVE)
        tbl.DataBodyRange.Rows(2) = Array("CB2", 0.8, 280, 40, 2.2, 14, 5.5, 9, 0.08, Date - 4, "Yes", Schema.ACTION_REMOVE)
        StyleActionColumn tbl, Schema.IR_COL_ACTION
    End If
End Sub

Private Sub SeedConfig()
    Dim ws As Worksheet, tbl As ListObject
    Set ws = ThisWorkbook.Worksheets(Schema.SHEET_CONFIG)

    Set tbl = ws.ListObjects(Schema.TABLE_CATALOG)
    If Not tbl Is Nothing Then
        EnsureRows tbl, 2
        tbl.DataBodyRange.Rows(1) = Array("RP1", "CB1", 1.5)
        tbl.DataBodyRange.Rows(2) = Array("RP1", "CB2", 0.8)
    End If

    Set tbl = ws.ListObjects(Schema.TABLE_TRIGGER)
    If Not tbl Is Nothing Then
        EnsureRows tbl, 2
        tbl.DataBodyRange.Rows(1) = Array("L1", 210, 350, 50, 3.5, 18, 7.5, 12, 0.15)
        tbl.DataBodyRange.Rows(2) = Array("L2", 200, 450, 100, 4.2, 22, 9, 15, 0.18)
    End If
End Sub

Private Sub SeedResults()
    Dim ws As Worksheet, tbl As ListObject, d As Date
    Set ws = ThisWorkbook.Worksheets(Schema.SHEET_RESULTS)
    Set tbl = ws.ListObjects(Schema.TABLE_RESULTS)
    d = Date - 10
    If Not tbl Is Nothing Then
        EnsureRows tbl, 3
        tbl.DataBodyRange.Rows(1) = Array("RP1", d, "RP1-001", 280, 45, 2.5, 15, 6, 10, 0.1)
        tbl.DataBodyRange.Rows(2) = Array("CB1", d + 1, "CB1-001", 350, 60, 2.8, 16, 7, 11, 0.12)
        tbl.DataBodyRange.Rows(3) = Array("CB2", d + 2, "CB2-001", 280, 40, 2.2, 14, 5.5, 9, 0.08)
    End If
End Sub

Private Sub SeedTelemetry()
    ' Seeds 14 days of sample telemetry data (rain only - EC/Vol are per-site)
    ' Run Initialize after SeedConfig to add site-specific columns
    Dim ws As Worksheet, tbl As ListObject, d As Date, i As Long
    Dim rain As Variant
    Set ws = ThisWorkbook.Worksheets(Schema.SHEET_TELEMETRY)
    Set tbl = ws.ListObjects(Schema.TABLE_TELEMETRY)
    d = Date - 14
    ' Rain data (mm) - all days have data
    rain = Array(0, 2.5, 0, 8.3, 1.2, 0, 0, 5.6, 3.1, 0, 12.4, 4.2, 0.8, 0)

    If Not tbl Is Nothing Then
        EnsureRows tbl, 14
        For i = 0 To 13
            tbl.DataBodyRange.Cells(i + 1, 1) = d + i
            tbl.DataBodyRange.Cells(i + 1, 2) = rain(i)
        Next i
    End If
End Sub

Private Sub SeedSiteTelemetry(ByVal site As String)
    ' Seeds sample EC/Vol data for a specific site
    Dim ws As Worksheet, tbl As ListObject, i As Long
    Dim ecCol As Long, volCol As Long
    Dim ec As Variant, vol As Variant

    Set ws = ThisWorkbook.Worksheets(Schema.SHEET_TELEMETRY)
    Set tbl = ws.ListObjects(Schema.TABLE_TELEMETRY)
    If tbl Is Nothing Then Exit Sub
    If tbl.DataBodyRange Is Nothing Then Exit Sub

    ' Get column indices for this site
    On Error Resume Next
    ecCol = tbl.ListColumns(Schema.TelemECColName(site)).Index
    volCol = tbl.ListColumns(Schema.TelemVolColName(site)).Index
    On Error GoTo 0
    If ecCol = 0 Or volCol = 0 Then Exit Sub

    ' Sample data - EC with some gaps, Volume sparse
    ec = Array(280, 285, 290, Empty, Empty, 310, 305, 300, Empty, 295, 290, 285, 280, 275)
    vol = Array(Empty, Empty, Empty, Empty, Empty, 155, Empty, Empty, Empty, Empty, 160, Empty, Empty, 158)

    For i = 0 To 13
        If i < tbl.ListRows.Count Then
            If Not IsEmpty(ec(i)) Then tbl.DataBodyRange.Cells(i + 1, ecCol) = ec(i)
            If Not IsEmpty(vol(i)) Then tbl.DataBodyRange.Cells(i + 1, volCol) = vol(i)
        End If
    Next i
End Sub

' ==== Chart Sheet ============================================================

Private Sub SetupChart()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(Schema.SHEET_CHART)
    ws.Range("A1") = "Simulation Charts": ws.Range("A1").Font.Bold = True
    ws.Range("A2") = "Run WQOC.Run to generate charts"
End Sub

' ==== Per-Site Table Creation (called on-demand) =============================

Public Sub EnsureSiteLogTable(ByVal site As String)
    ' Creates log table for site if it doesn't exist
    Dim ws As Worksheet, tbl As ListObject, tblName As String
    Dim chem As Variant, n As Long, h() As String, i As Long
    Dim startCol As Long

    Set ws = ThisWorkbook.Worksheets(Schema.SHEET_LOG)
    tblName = Schema.LogTableName(site)

    ' Check if table already exists
    On Error Resume Next
    Set tbl = ws.ListObjects(tblName)
    On Error GoTo 0
    If Not tbl Is Nothing Then Exit Sub

    ' Find position for new table (after existing tables)
    startCol = FindNextTableColumn(ws)

    ' Build header
    chem = Schema.ChemistryNames(): n = Schema.ChemistryCount()
    ReDim h(1 To n + 4)
    h(1) = "RunId": h(2) = "Date": h(3) = "Day": h(4) = Schema.VOLUME_METRIC_NAME
    For i = 1 To n: h(4 + i) = chem(i - 1): Next i

    ' Add site label above table
    ws.Cells(1, startCol).Value = site & " Log"
    ws.Cells(1, startCol).Font.Bold = True

    ' Create table
    MakeTbl ws, ws.Cells(3, startCol), tblName, h
End Sub

Public Sub EnsureSiteHistoryTable(ByVal site As String)
    ' Creates history table for site if it doesn't exist
    Dim ws As Worksheet, tbl As ListObject, tblName As String
    Dim startCol As Long

    Set ws = ThisWorkbook.Worksheets(Schema.SHEET_HISTORY)
    tblName = Schema.HistoryTableName(site)

    ' Check if table already exists
    On Error Resume Next
    Set tbl = ws.ListObjects(tblName)
    On Error GoTo 0
    If Not tbl Is Nothing Then Exit Sub

    ' Find position for new table (after existing tables)
    startCol = FindNextTableColumn(ws)

    ' Add site label above table
    ws.Cells(1, startCol).Value = site & " History"
    ws.Cells(1, startCol).Font.Bold = True

    ' Create table (no Site column - site is in table name)
    MakeTbl ws, ws.Cells(3, startCol), tblName, _
        Array("RunId", "Timestamp", "RunDate", "Days", "Mode", "TriggerDay", "TriggerMetric", Schema.HISTORY_COL_ACTION)

    ' Style action column header
    Set tbl = ws.ListObjects(tblName)
    If Not tbl Is Nothing Then
        StyleActionHeader tbl, Schema.HISTORY_COL_ACTION, ""
    End If
End Sub

Public Sub EnsureSiteTables(ByVal site As String)
    ' Ensures both log and history tables exist for site
    EnsureSiteLogTable site
    EnsureSiteHistoryTable site
End Sub

Private Function FindNextTableColumn(ByVal ws As Worksheet) As Long
    ' Returns the column where the next table should start (after existing tables + gap)
    Dim tbl As ListObject, maxCol As Long
    maxCol = 0
    For Each tbl In ws.ListObjects
        If tbl.Range.Column + tbl.Range.Columns.Count > maxCol Then
            maxCol = tbl.Range.Column + tbl.Range.Columns.Count
        End If
    Next tbl
    If maxCol = 0 Then
        FindNextTableColumn = 1
    Else
        FindNextTableColumn = maxCol + Schema.TABLE_GAP_COLS
    End If
End Function

' ==== Controls (Run Cell, Dropdowns) =========================================

Private Sub SetupControls()
    Dim ws As Worksheet, runCell As Range
    Set ws = ThisWorkbook.Worksheets(Schema.SHEET_INPUT)

    ' Remove old buttons if they exist
    On Error Resume Next
    ws.Shapes("btnRun").Delete
    ws.Shapes("btnRollback").Delete
    ws.Shapes("btnRefresh").Delete
    On Error GoTo 0

    ' Create Run Simulation cell (L1) - double-click to run
    Set runCell = ws.Range("L1")
    runCell.Value = "Run Simulation"
    With runCell
        .Font.Bold = True
        .Font.Color = Schema.COLOR_FONT_WHITE
        .Interior.Color = Schema.COLOR_BUTTON_ON
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .Borders.LineStyle = xlContinuous
        .Borders.Weight = xlThin
    End With
    AddNm Schema.NAME_RUN_CELL, runCell

    ' Rain mode dropdown validation
    With ws.Range(Schema.NAME_RAIN_MODE).Validation
        .Delete
        .Add Type:=xlValidateList, Formula1:=Schema.RAIN_MODE_LIST
    End With

    ' Enhanced mode dropdown validation
    With ws.Range(Schema.NAME_ENHANCED_MODE).Validation
        .Delete
        .Add Type:=xlValidateList, Formula1:="On,Off"
    End With

    ' Site dropdown validation (from tblCatalog RR column)
    With ws.Range(Schema.NAME_SITE).Validation
        .Delete
        .Add Type:=xlValidateList, Formula1:="=INDIRECT(""" & Schema.TABLE_CATALOG & "[RR]"")"
    End With

    ' Mixing model dropdown validation
    With ws.Range(Schema.NAME_MIXING_MODEL).Validation
        .Delete
        .Add Type:=xlValidateList, Formula1:=Schema.MIXING_MODEL_LIST
    End With

    ' Rainfall mode dropdown validation
    With ws.Range(Schema.NAME_RAINFALL_MODE).Validation
        .Delete
        .Add Type:=xlValidateList, Formula1:=Schema.RAINFALL_MODE_LIST
    End With

    ' Telemetry calibration dropdown validation
    With ws.Range(Schema.NAME_TELEM_CAL).Validation
        .Delete
        .Add Type:=xlValidateList, Formula1:=Schema.TELEM_CAL_LIST
    End With
End Sub

' ==== Helpers ================================================================

Private Sub AddNm(ByVal nm As String, ByVal rng As Range)
    On Error Resume Next: ThisWorkbook.Names(nm).Delete: On Error GoTo 0
    ThisWorkbook.Names.Add nm, "=" & rng.Address(True, True, xlA1, True)
End Sub

Private Sub MakeTbl(ByVal ws As Worksheet, ByVal start As Range, ByVal nm As String, ByVal h As Variant)
    Dim n As Long, tbl As ListObject
    On Error Resume Next: ws.ListObjects(nm).Delete: On Error GoTo 0
    If IsArray(h) Then n = UBound(h) - LBound(h) + 1 Else n = 1
    start.Resize(1, n) = h: start.Resize(1, n).Font.Bold = True
    Set tbl = ws.ListObjects.Add(xlSrcRange, start.Resize(2, n), , xlYes)
    tbl.Name = nm
    On Error Resume Next: tbl.TableStyle = Schema.TABLE_STYLE_DEFAULT: On Error GoTo 0
End Sub

Private Sub EnsureRows(ByVal tbl As ListObject, ByVal n As Long)
    Do While tbl.ListRows.Count < n: tbl.ListRows.Add: Loop
    If Not tbl.DataBodyRange Is Nothing Then tbl.DataBodyRange.ClearContents
End Sub

Private Sub StyleActionHeader(ByVal tbl As ListObject, ByVal colName As String, ByVal txt As String)
    ' Style the header cell of an action column (blue, underlined, hyperlink-like)
    Dim col As ListColumn, hdrCell As Range
    On Error Resume Next
    Set col = tbl.ListColumns(colName)
    On Error GoTo 0
    If col Is Nothing Then Exit Sub

    Set hdrCell = tbl.HeaderRowRange.Cells(1, col.Index)
    If Len(txt) > 0 Then hdrCell.Value = txt
    With hdrCell
        .Font.Color = Schema.COLOR_ACTION_FONT
        .Font.Underline = xlUnderlineStyleSingle
        .Font.Bold = False
    End With
End Sub

Private Sub StyleActionColumn(ByVal tbl As ListObject, ByVal colName As String)
    ' Style data cells in action column (blue, underlined)
    Dim col As ListColumn, dataRng As Range
    On Error Resume Next
    Set col = tbl.ListColumns(colName)
    On Error GoTo 0
    If col Is Nothing Then Exit Sub
    If tbl.DataBodyRange Is Nothing Then Exit Sub

    Set dataRng = tbl.DataBodyRange.Columns(col.Index)
    With dataRng
        .Font.Color = Schema.COLOR_ACTION_FONT
        .Font.Underline = xlUnderlineStyleSingle
    End With
End Sub
