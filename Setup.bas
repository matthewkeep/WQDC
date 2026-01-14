Option Explicit
' Setup: Workbook scaffolding and test data.
' Dependencies: Schema

Public Sub RepairEvents()
    ' Resets Application state - call if events stop working after an error
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
End Sub

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

Public Sub BuildAll(): RepairEvents: Build: Seed: Initialize: End Sub

Public Sub Initialize()
    ' Reads all RR sites from tblCatalog and creates per-site infrastructure:
    ' - Telemetry columns (EC, Vol for each site)
    ' - Log tables (tblLog_{site})
    ' - History tables (tblHistory_{site})
    ' Safe to run multiple times - only creates what doesn't exist
    Dim cm As XlCalculation, sites As Variant, site As Variant
    Dim created As Long
    RepairEvents  ' Ensure clean state before starting
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
    AddNm Schema.NAME_PRED_ROW, ws.Range("C5").Resize(1, n)

    ' Run info
    ws.Range("J2") = "Run Date": AddNm Schema.NAME_RUN_DATE, ws.Range("K2")
    ws.Range("J3") = "Site": AddNm Schema.NAME_SITE, ws.Range("K3")
    ws.Range("J4") = "Output": AddNm Schema.NAME_OUTPUT, ws.Range("K4")
    ws.Range("J5") = "Sample Date": AddNm Schema.NAME_SAMPLE_DATE, ws.Range("K5")

    ' Results
    ws.Range("N1") = "Results": ws.Range("N1:O1").Font.Bold = True
    ws.Range("N2") = "Standard": AddNm Schema.NAME_STD_TRIGGER, ws.Range("O2")
    ws.Range("N3") = "Enhanced": AddNm Schema.NAME_ENH_TRIGGER, ws.Range("O3")

    ' Model Parameters (apply to both Standard and Enhanced)
    ws.Range("N5") = "Parameters": ws.Range("N5:O5").Font.Bold = True
    ws.Range("N6") = "Net Outflow (ML/d)": AddNm Schema.NAME_NET_OUT, ws.Range("O6")
    ws.Range("N7") = "Tau (days)": AddNm Schema.NAME_TAU, ws.Range("O7")
    ws.Range("N8") = "Surface Fraction": AddNm Schema.NAME_SURFACE_FRACTION, ws.Range("O8")

    ' Enhanced Mode (toggle + sub-options)
    ws.Range("N10") = "Enhanced Mode": ws.Range("N10:O10").Font.Bold = True
    ws.Range("N11") = "Enabled": AddNm Schema.NAME_ENHANCED_MODE, ws.Range("O11")
    ws.Range("N12") = "Mixing Model": AddNm Schema.NAME_MIXING_MODEL, ws.Range("O12")
    ws.Range("N13") = "Rainfall": AddNm Schema.NAME_RAINFALL_MODE, ws.Range("O13")
    ws.Range("N14") = "Telemetry Cal": AddNm Schema.NAME_TELEM_CAL, ws.Range("O14")

    ' Apply conditional formatting to grey out Enhanced options when disabled
    ApplyEnhancedConditionalFormat ws.Range("N12:O14"), ws.Range("O11")

    ' Hidden mass (internal state for TwoBucket)
    ws.Range("Q1") = "Hidden Mass": ws.Range("Q1:R1").Font.Bold = True
    For i = 0 To n - 1: ws.Cells(2 + i, 17) = chem(i): Next i
    AddNm Schema.NAME_HIDDEN_MASS, ws.Range("R2").Resize(n, 1)

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

    ' Style action column header (keep column name as "Action" for event lookup)
    Set tbl = ws.ListObjects(Schema.TABLE_IR)
    If Not tbl Is Nothing Then
        StyleActionHeader tbl, Schema.IR_COL_ACTION, ""
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
    ws.Range(Schema.NAME_SAMPLE_DATE) = Date - 7  ' Most recent weekly sample
    ws.Range(Schema.NAME_INIT_VOL) = 165          ' Current reservoir volume (ML)
    ws.Range(Schema.NAME_TRIGGER_VOL) = 200       ' Volume trigger (ML)

    ' Current RR chemistry (from latest sample)
    ws.Range(Schema.NAME_RES_ROW) = Array(365, 58, 3.2, 19, 7.2, 11.8, 0.13)

    ' Release triggers (regulatory limits)
    ws.Range(Schema.NAME_LIMIT_ROW) = Array(450, 100, 5, 25, 10, 18, 0.25)

    ' Model parameters
    ws.Range(Schema.NAME_NET_OUT) = 1.2           ' Net outflow (ML/d)
    ws.Range(Schema.NAME_TAU) = 7                 ' Mixing time constant (days)
    ws.Range(Schema.NAME_SURFACE_FRACTION) = 0.8  ' TwoBucket surface fraction

    ' Enhanced config defaults
    ws.Range(Schema.NAME_ENHANCED_MODE) = "On"
    ws.Range(Schema.NAME_MIXING_MODEL) = Schema.MIXING_TWOBUCKET
    ws.Range(Schema.NAME_RAINFALL_MODE) = Schema.RAINFALL_HINDCAST
    ws.Range(Schema.NAME_TELEM_CAL) = Schema.TELEM_CAL_ON

    ' IR table will be populated by Loader when site is selected
    Dim tbl As ListObject
    Set tbl = ws.ListObjects(Schema.TABLE_IR)
    If Not tbl Is Nothing Then
        If Not tbl.DataBodyRange Is Nothing Then tbl.DataBodyRange.Delete
    End If
End Sub

Private Sub SeedConfig()
    Dim ws As Worksheet, tbl As ListObject
    Set ws = ThisWorkbook.Worksheets(Schema.SHEET_CONFIG)

    ' Catalog: RR sites with their IR inflow sources
    Set tbl = ws.ListObjects(Schema.TABLE_CATALOG)
    If Not tbl Is Nothing Then
        If Not tbl.DataBodyRange Is Nothing Then tbl.DataBodyRange.Delete
        EnsureRows tbl, 4
        ' RP1: Main reservoir with 2 catchment basins
        tbl.DataBodyRange.Rows(1) = Array("RP1", "CB1", 2.5)    ' High flow, cleaner
        tbl.DataBodyRange.Rows(2) = Array("RP1", "CB2", 1.8)    ' Medium flow
        ' RP2: Secondary reservoir
        tbl.DataBodyRange.Rows(3) = Array("RP2", "CB3", 1.5)    ' Clean catchment
        tbl.DataBodyRange.Rows(4) = Array("RP2", "CB4", 1.2)    ' Clean catchment
    End If

    ' Trigger levels (regulatory release limits)
    Set tbl = ws.ListObjects(Schema.TABLE_TRIGGER)
    If Not tbl Is Nothing Then
        If Not tbl.DataBodyRange Is Nothing Then tbl.DataBodyRange.Delete
        EnsureRows tbl, 2
        '                    Name, Vol,  EC,   F_U,  F_Mn, SO4,  Mg,   Ca,   TAN
        tbl.DataBodyRange.Rows(1) = Array("Alert", 180, 400, 80, 4, 20, 8, 15, 0.2)
        tbl.DataBodyRange.Rows(2) = Array("Limit", 200, 450, 100, 5, 25, 10, 18, 0.25)
    End If
End Sub

Private Sub SeedResults()
    ' Seeds minimal results for quick testing
    ' Use SeedFullSeason for comprehensive backtest data
    Dim ws As Worksheet, tbl As ListObject, d As Date
    Set ws = ThisWorkbook.Worksheets(Schema.SHEET_RESULTS)
    Set tbl = ws.ListObjects(Schema.TABLE_RESULTS)
    If tbl Is Nothing Then Exit Sub
    If Not tbl.DataBodyRange Is Nothing Then tbl.DataBodyRange.Delete

    d = Date - 14  ' Two weeks of samples

    EnsureRows tbl, 8
    '                         Site, Date, SampleID, EC, F_U, F_Mn, SO4, Mg, Ca, TAN
    ' RP1 samples (2 weeks)
    tbl.DataBodyRange.Rows(1) = Array("RP1", d, "RP1-001", 340, 52, 2.9, 17, 6.5, 10.5, 0.11)
    tbl.DataBodyRange.Rows(2) = Array("RP1", d + 7, "RP1-002", 365, 58, 3.2, 19, 7.2, 11.8, 0.13)

    ' RP2 samples (2 weeks)
    tbl.DataBodyRange.Rows(3) = Array("RP2", d + 1, "RP2-001", 330, 48, 2.7, 16, 6.2, 10, 0.1)
    tbl.DataBodyRange.Rows(4) = Array("RP2", d + 8, "RP2-002", 315, 45, 2.5, 15, 5.8, 9.5, 0.09)

    ' IR source samples
    tbl.DataBodyRange.Rows(5) = Array("CB1", d + 2, "CB1-001", 195, 28, 1.6, 9, 4, 6.5, 0.05)
    tbl.DataBodyRange.Rows(6) = Array("CB2", d + 3, "CB2-001", 235, 38, 1.9, 11, 4.8, 8, 0.07)
    tbl.DataBodyRange.Rows(7) = Array("CB3", d + 4, "CB3-001", 170, 22, 1.3, 7.5, 3.2, 5.8, 0.04)
    tbl.DataBodyRange.Rows(8) = Array("CB4", d + 5, "CB4-001", 200, 30, 1.5, 9.5, 4, 7, 0.055)
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

' ==== Full Season Test Data ==================================================

Public Sub SeedFullSeason()
    ' Creates comprehensive test data for a full wet season (90 days)
    ' Includes: 2 RR sites, 4 IR sources each, daily telemetry, weekly samples
    ' Run BuildAll first, then SeedFullSeason, then Initialize
    Dim cm As XlCalculation
    On Error GoTo Fail
    cm = Application.Calculation
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual

    SeedFullCatalog
    SeedFullResults
    SeedFullTelemetry

    Application.Calculation = cm
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    MsgBox "Full season data seeded (90 days, 2 sites)." & vbNewLine & _
           "Run 'Initialize' to create site columns.", vbInformation, "Setup"
    Exit Sub
Fail:
    Application.Calculation = cm
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    MsgBox "Error: " & Err.Description, vbExclamation, "Setup"
End Sub

Private Sub SeedFullCatalog()
    ' Seeds catalog with 2 RR sites and their IR sources
    Dim ws As Worksheet, tbl As ListObject
    Set ws = ThisWorkbook.Worksheets(Schema.SHEET_CONFIG)
    Set tbl = ws.ListObjects(Schema.TABLE_CATALOG)
    If tbl Is Nothing Then Exit Sub

    ' Clear existing and add comprehensive catalog
    If Not tbl.DataBodyRange Is Nothing Then tbl.DataBodyRange.Delete
    EnsureRows tbl, 8

    ' RP1: Main reservoir - 4 inflow sources
    tbl.DataBodyRange.Rows(1) = Array("RP1", "CB1", 2.5)    ' High flow, cleaner
    tbl.DataBodyRange.Rows(2) = Array("RP1", "CB2", 1.8)    ' Medium flow
    tbl.DataBodyRange.Rows(3) = Array("RP1", "MW1", 0.6)    ' Mine water - higher EC
    tbl.DataBodyRange.Rows(4) = Array("RP1", "MW2", 0.4)    ' Mine water - highest EC

    ' RP2: Secondary reservoir - 4 inflow sources
    tbl.DataBodyRange.Rows(5) = Array("RP2", "CB3", 1.5)    ' Clean catchment
    tbl.DataBodyRange.Rows(6) = Array("RP2", "CB4", 1.2)    ' Clean catchment
    tbl.DataBodyRange.Rows(7) = Array("RP2", "MW3", 0.8)    ' Mine affected
    tbl.DataBodyRange.Rows(8) = Array("RP2", "PW1", 0.3)    ' Process water
End Sub

Private Sub SeedFullResults()
    ' Seeds ~90 days of weekly samples for RR sites and IR sources
    ' Realistic mining wastewater chemistry progression
    Dim ws As Worksheet, tbl As ListObject
    Dim baseDate As Date, d As Date, i As Long, row As Long
    Dim ec As Double, fU As Double, fMn As Double, so4 As Double
    Dim mg As Double, ca As Double, tan As Double

    Set ws = ThisWorkbook.Worksheets(Schema.SHEET_RESULTS)
    Set tbl = ws.ListObjects(Schema.TABLE_RESULTS)
    If tbl Is Nothing Then Exit Sub
    If Not tbl.DataBodyRange Is Nothing Then tbl.DataBodyRange.Delete

    baseDate = Date - 90  ' Start 90 days ago
    row = 0

    ' --- RP1 weekly samples (13 weeks) ---
    ' EC trending up slightly over wet season due to mine water mixing
    For i = 0 To 12
        d = baseDate + (i * 7)
        ec = 280 + (i * 8) + Rnd() * 20 - 10       ' 280-400 trend up
        fU = 42 + (i * 1.5) + Rnd() * 5 - 2.5     ' 42-65 trend up
        fMn = 2.3 + (i * 0.1) + Rnd() * 0.3       ' 2.3-3.8 slight increase
        so4 = 14 + (i * 0.5) + Rnd() * 2          ' 14-22 gradual rise
        mg = 5.5 + (i * 0.15) + Rnd() * 0.5       ' 5.5-7.8
        ca = 9 + (i * 0.2) + Rnd() * 1            ' 9-12.5
        tan = 0.08 + (i * 0.005) + Rnd() * 0.02   ' 0.08-0.15
        row = row + 1
        EnsureRows tbl, row
        tbl.DataBodyRange.Rows(row) = Array("RP1", d, "RP1-" & Format$(i + 1, "000"), _
            Round(ec, 0), Round(fU, 1), Round(fMn, 2), Round(so4, 1), _
            Round(mg, 1), Round(ca, 1), Round(tan, 3))
    Next i

    ' --- RP2 weekly samples (13 weeks) ---
    ' Different profile: starts higher, stabilizes (better mixing)
    For i = 0 To 12
        d = baseDate + (i * 7) + 1  ' Offset by 1 day
        ec = 350 - (i * 5) + Rnd() * 15 - 7.5     ' 350-290 trending down
        fU = 55 - (i * 1) + Rnd() * 4 - 2         ' 55-40 improving
        fMn = 3.0 - (i * 0.05) + Rnd() * 0.2      ' 3.0-2.4
        so4 = 18 - (i * 0.3) + Rnd() * 2          ' 18-14
        mg = 7 - (i * 0.1) + Rnd() * 0.4          ' 7-5.5
        ca = 11 - (i * 0.15) + Rnd() * 0.8        ' 11-9
        tan = 0.12 - (i * 0.003) + Rnd() * 0.015  ' 0.12-0.08
        row = row + 1
        EnsureRows tbl, row
        tbl.DataBodyRange.Rows(row) = Array("RP2", d, "RP2-" & Format$(i + 1, "000"), _
            Round(ec, 0), Round(fU, 1), Round(fMn, 2), Round(so4, 1), _
            Round(mg, 1), Round(ca, 1), Round(tan, 3))
    Next i

    ' --- IR source samples (fortnightly) ---
    ' CB1: Clean catchment - low EC
    For i = 0 To 5
        d = baseDate + (i * 14) + 2
        row = row + 1: EnsureRows tbl, row
        tbl.DataBodyRange.Rows(row) = Array("CB1", d, "CB1-" & Format$(i + 1, "000"), _
            180 + Rnd() * 30, 25 + Rnd() * 8, 1.5 + Rnd() * 0.4, 8 + Rnd() * 3, _
            3.5 + Rnd() * 1, 6 + Rnd() * 2, 0.04 + Rnd() * 0.02)
    Next i

    ' CB2: Moderate catchment
    For i = 0 To 5
        d = baseDate + (i * 14) + 3
        row = row + 1: EnsureRows tbl, row
        tbl.DataBodyRange.Rows(row) = Array("CB2", d, "CB2-" & Format$(i + 1, "000"), _
            220 + Rnd() * 40, 35 + Rnd() * 10, 1.8 + Rnd() * 0.5, 10 + Rnd() * 4, _
            4.5 + Rnd() * 1.2, 7.5 + Rnd() * 2.5, 0.06 + Rnd() * 0.025)
    Next i

    ' MW1: Mine water - elevated EC
    For i = 0 To 5
        d = baseDate + (i * 14) + 4
        row = row + 1: EnsureRows tbl, row
        tbl.DataBodyRange.Rows(row) = Array("MW1", d, "MW1-" & Format$(i + 1, "000"), _
            850 + Rnd() * 150, 120 + Rnd() * 30, 5.5 + Rnd() * 1.5, 45 + Rnd() * 15, _
            15 + Rnd() * 4, 22 + Rnd() * 6, 0.25 + Rnd() * 0.08)
    Next i

    ' MW2: Mine water - highest EC
    For i = 0 To 5
        d = baseDate + (i * 14) + 5
        row = row + 1: EnsureRows tbl, row
        tbl.DataBodyRange.Rows(row) = Array("MW2", d, "MW2-" & Format$(i + 1, "000"), _
            1200 + Rnd() * 200, 180 + Rnd() * 40, 7.5 + Rnd() * 2, 65 + Rnd() * 20, _
            20 + Rnd() * 5, 30 + Rnd() * 8, 0.35 + Rnd() * 0.1)
    Next i

    ' CB3, CB4: Clean catchments for RP2
    For i = 0 To 5
        d = baseDate + (i * 14) + 6
        row = row + 1: EnsureRows tbl, row
        tbl.DataBodyRange.Rows(row) = Array("CB3", d, "CB3-" & Format$(i + 1, "000"), _
            160 + Rnd() * 25, 20 + Rnd() * 6, 1.2 + Rnd() * 0.3, 7 + Rnd() * 2, _
            3 + Rnd() * 0.8, 5.5 + Rnd() * 1.5, 0.03 + Rnd() * 0.015)
    Next i

    For i = 0 To 5
        d = baseDate + (i * 14) + 7
        row = row + 1: EnsureRows tbl, row
        tbl.DataBodyRange.Rows(row) = Array("CB4", d, "CB4-" & Format$(i + 1, "000"), _
            190 + Rnd() * 30, 28 + Rnd() * 7, 1.4 + Rnd() * 0.35, 9 + Rnd() * 2.5, _
            3.8 + Rnd() * 0.9, 6.5 + Rnd() * 1.8, 0.05 + Rnd() * 0.02)
    Next i

    ' MW3: Mine affected for RP2
    For i = 0 To 5
        d = baseDate + (i * 14) + 8
        row = row + 1: EnsureRows tbl, row
        tbl.DataBodyRange.Rows(row) = Array("MW3", d, "MW3-" & Format$(i + 1, "000"), _
            750 + Rnd() * 120, 95 + Rnd() * 25, 4.5 + Rnd() * 1.2, 38 + Rnd() * 12, _
            12 + Rnd() * 3.5, 18 + Rnd() * 5, 0.2 + Rnd() * 0.06)
    Next i

    ' PW1: Process water - variable
    For i = 0 To 5
        d = baseDate + (i * 14) + 9
        row = row + 1: EnsureRows tbl, row
        tbl.DataBodyRange.Rows(row) = Array("PW1", d, "PW1-" & Format$(i + 1, "000"), _
            600 + Rnd() * 300, 80 + Rnd() * 40, 3.5 + Rnd() * 2, 30 + Rnd() * 20, _
            10 + Rnd() * 5, 15 + Rnd() * 8, 0.15 + Rnd() * 0.1)
    Next i
End Sub

Private Sub SeedFullTelemetry()
    ' Seeds 90 days of daily telemetry with realistic wet season patterns
    Dim ws As Worksheet, tbl As ListObject
    Dim baseDate As Date, d As Date, i As Long
    Dim rain As Double

    Set ws = ThisWorkbook.Worksheets(Schema.SHEET_TELEMETRY)
    Set tbl = ws.ListObjects(Schema.TABLE_TELEMETRY)
    If tbl Is Nothing Then Exit Sub
    If Not tbl.DataBodyRange Is Nothing Then tbl.DataBodyRange.Delete

    baseDate = Date - 90

    ' Create 90 days of telemetry
    EnsureRows tbl, 90

    For i = 0 To 89
        d = baseDate + i

        ' Wet season rainfall pattern:
        ' - Most days are dry (0mm)
        ' - Some light rain (1-5mm)
        ' - Occasional moderate rain (5-20mm)
        ' - Rare heavy storms (20-80mm)
        Select Case Rnd()
            Case Is < 0.55: rain = 0                        ' 55% dry
            Case Is < 0.75: rain = Round(1 + Rnd() * 4, 1)  ' 20% light (1-5mm)
            Case Is < 0.9: rain = Round(5 + Rnd() * 15, 1)  ' 15% moderate (5-20mm)
            Case Is < 0.97: rain = Round(20 + Rnd() * 30, 1) ' 7% heavy (20-50mm)
            Case Else: rain = Round(50 + Rnd() * 30, 1)     ' 3% storm (50-80mm)
        End Select

        tbl.DataBodyRange.Cells(i + 1, 1) = d
        tbl.DataBodyRange.Cells(i + 1, 2) = rain
    Next i
End Sub

Public Sub SeedSiteTelemFull(ByVal site As String)
    ' Seeds 90 days of EC/Vol telemetry for a specific site
    ' Call after Initialize creates the site columns
    Dim ws As Worksheet, tbl As ListObject, i As Long
    Dim ecCol As Long, volCol As Long
    Dim baseEC As Double, baseVol As Double
    Dim ec As Double, vol As Double, rain As Double

    Set ws = ThisWorkbook.Worksheets(Schema.SHEET_TELEMETRY)
    Set tbl = ws.ListObjects(Schema.TABLE_TELEMETRY)
    If tbl Is Nothing Then Exit Sub
    If tbl.DataBodyRange Is Nothing Then Exit Sub

    On Error Resume Next
    ecCol = tbl.ListColumns(Schema.TelemECColName(site)).Index
    volCol = tbl.ListColumns(Schema.TelemVolColName(site)).Index
    On Error GoTo 0
    If ecCol = 0 Or volCol = 0 Then Exit Sub

    ' Set baseline based on site
    Select Case UCase$(site)
        Case "RP1": baseEC = 280: baseVol = 150
        Case "RP2": baseEC = 350: baseVol = 120
        Case Else: baseEC = 300: baseVol = 130
    End Select

    ec = baseEC
    vol = baseVol

    For i = 1 To tbl.ListRows.Count
        rain = tbl.DataBodyRange.Cells(i, 2).Value

        ' Volume responds to rain (simplified water balance)
        vol = vol + rain * 0.5 - 1.5  ' Catchment factor, outflow
        If vol < 50 Then vol = 50
        If vol > 250 Then vol = 250

        ' EC diluted by rain, concentrated by evap
        If rain > 5 Then
            ec = ec * 0.97 - rain * 0.3  ' Dilution from significant rain
        Else
            ec = ec * 1.002  ' Slight concentration
        End If
        If ec < 150 Then ec = 150
        If ec > 600 Then ec = 600

        ' Add measurement noise and some gaps
        If Rnd() > 0.15 Then  ' 85% data availability
            tbl.DataBodyRange.Cells(i, ecCol) = Round(ec + Rnd() * 10 - 5, 0)
        End If
        If Rnd() > 0.7 Then  ' 30% volume readings (less frequent)
            tbl.DataBodyRange.Cells(i, volCol) = Round(vol + Rnd() * 5 - 2.5, 1)
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

    ' Format Date column
    Set tbl = ws.ListObjects(tblName)
    If Not tbl Is Nothing Then
        tbl.ListColumns(2).DataBodyRange.NumberFormat = "dd-mmm-yyyy"
    End If
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

    ' Style action column header and format date columns
    Set tbl = ws.ListObjects(tblName)
    If Not tbl Is Nothing Then
        StyleActionHeader tbl, Schema.HISTORY_COL_ACTION, ""
        tbl.ListColumns(2).DataBodyRange.NumberFormat = "dd-mmm-yyyy hh:mm"  ' Timestamp
        tbl.ListColumns(3).DataBodyRange.NumberFormat = "dd-mmm-yyyy"         ' RunDate
    End If
End Sub

Public Sub EnsureSiteTables(ByVal site As String)
    ' Ensures both log and history tables exist for site
    EnsureSiteLogTable site
    EnsureSiteHistoryTable site
End Sub

Public Sub EnsureSeasonLogTable(ByVal site As String)
    ' Creates season backtest log table for site if it doesn't exist
    ' Columns support A/B comparison between Standard and Enhanced modes
    Dim ws As Worksheet, tbl As ListObject, tblName As String
    Dim startCol As Long

    Set ws = ThisWorkbook.Worksheets(Schema.SHEET_LOG)
    tblName = Schema.SeasonLogTableName(site)

    ' Check if table already exists
    On Error Resume Next
    Set tbl = ws.ListObjects(tblName)
    On Error GoTo 0
    If Not tbl Is Nothing Then Exit Sub

    ' Find position for new table (after existing tables)
    startCol = FindNextTableColumn(ws)

    ' Add site label above table
    ws.Cells(1, startCol).Value = site & " Season Backtest"
    ws.Cells(1, startCol).Font.Bold = True

    ' Create table with A/B comparison columns
    MakeTbl ws, ws.Cells(3, startCol), tblName, _
        Array("RunDate", "SampleDate", "ActualEC", "ActualVol", _
              "StdPredEC", "StdErrEC", "StdPredVol", "StdErrVol", _
              "EnhPredEC", "EnhErrEC", "EnhPredVol", "EnhErrVol")

    ' Format date columns
    Set tbl = ws.ListObjects(tblName)
    If Not tbl Is Nothing Then
        tbl.ListColumns(1).DataBodyRange.NumberFormat = "dd-mmm-yyyy"  ' RunDate
        tbl.ListColumns(2).DataBodyRange.NumberFormat = "dd-mmm-yyyy"  ' SampleDate
    End If
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

Private Sub ApplyEnhancedConditionalFormat(ByVal targetRange As Range, ByVal toggleCell As Range)
    ' Greys out target range when Enhanced Mode is Off
    Dim fc As FormatCondition
    targetRange.FormatConditions.Delete
    Set fc = targetRange.FormatConditions.Add(Type:=xlExpression, _
        Formula1:="=" & toggleCell.Address(True, True) & "<>""On""")
    With fc
        .Font.Color = RGB(180, 180, 180)  ' Grey text
        .Interior.Color = RGB(242, 242, 242)  ' Light grey background
    End With
End Sub

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
