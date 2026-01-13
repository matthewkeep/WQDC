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
    SetupRain
    SetupHistory
    SetupChart
    SetupControls

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
    SeedRain

    Application.Calculation = cm: Application.ScreenUpdating = True: Application.EnableEvents = True
    MsgBox "Test data seeded.", vbInformation, "Setup"
    Exit Sub
Fail:
    Application.Calculation = cm: Application.ScreenUpdating = True: Application.EnableEvents = True
    MsgBox "Error: " & Err.Description, vbExclamation, "Setup"
End Sub

Public Sub BuildAll(): Build: Seed: End Sub

Public Sub Clean()
    Dim ws As Worksheet, nm As Name, sheets As Variant, i As Long
    sheets = Array(Schema.SHEET_INPUT, Schema.SHEET_CONFIG, Schema.SHEET_RESULTS, _
                   Schema.SHEET_RAIN, Schema.SHEET_HISTORY, Schema.SHEET_CHART, Schema.SHEET_LOG)
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
    MakeSheet Schema.SHEET_RAIN
    MakeSheet Schema.SHEET_HISTORY
    MakeSheet Schema.SHEET_CHART
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
    AddNm Schema.NAME_TRIGGER_RESULT_VOL, ws.Range("B5")
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

    ' Hidden mass
    ws.Range("Q6") = "Hidden Mass": ws.Range("Q6:R6").Font.Bold = True
    For i = 0 To n - 1: ws.Cells(7 + i, 17) = chem(i): Next i
    AddNm Schema.NAME_HIDDEN_MASS, ws.Range("R7").Resize(n, 1)

    ' IR Table
    ws.Range("A8") = "Inflow Sources": ws.Range("A8:L8").Font.Bold = True
    MakeIRTable ws, chem, n
End Sub

Private Sub MakeIRTable(ByVal ws As Worksheet, ByVal chem As Variant, ByVal n As Long)
    Dim h() As String, i As Long
    ReDim h(1 To n + 4)
    h(1) = Schema.IR_COL_SOURCE: h(2) = Schema.IR_COL_FLOW
    For i = 0 To n - 1: h(3 + i) = chem(i): Next i
    h(n + 3) = Schema.IR_COL_SAMPLE_DATE: h(n + 4) = Schema.IR_COL_ACTIVE
    MakeTbl ws, ws.Range("A9"), Schema.TABLE_IR, h
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

' ==== Rain Sheet =============================================================

Private Sub SetupRain()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(Schema.SHEET_RAIN)
    ws.Range("A1") = "Rainfall": ws.Range("A1").Font.Bold = True
    MakeTbl ws, ws.Range("A2"), Schema.TABLE_RAIN, Array("Date", "Rain (mm)")
End Sub

' ==== History Sheet ==========================================================

Private Sub SetupHistory()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(Schema.SHEET_HISTORY)
    ws.Range("A1") = "History": ws.Range("A1").Font.Bold = True
    MakeTbl ws, ws.Range("A2"), Schema.TABLE_HISTORY, _
        Array("RunId", "Timestamp", "RunDate", "Site", "SampleDate", "Mode", "TriggerDay", "TriggerMetric", "Status")
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
    ws.Range(Schema.NAME_ENHANCED_MODE) = "On"

    Dim tbl As ListObject
    Set tbl = ws.ListObjects(Schema.TABLE_IR)
    If Not tbl Is Nothing Then
        EnsureRows tbl, 2
        tbl.DataBodyRange.Rows(1) = Array("CB1", 1.5, 350, 60, 2.8, 16, 7, 11, 0.12, Date - 3, "Yes")
        tbl.DataBodyRange.Rows(2) = Array("CB2", 0.8, 280, 40, 2.2, 14, 5.5, 9, 0.08, Date - 4, "Yes")
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

Private Sub SeedRain()
    Dim ws As Worksheet, tbl As ListObject, d As Date, i As Long
    Dim rain As Variant
    Set ws = ThisWorkbook.Worksheets(Schema.SHEET_RAIN)
    Set tbl = ws.ListObjects(Schema.TABLE_RAIN)
    d = Date - 14
    rain = Array(0, 2.5, 0, 8.3, 1.2, 0, 0, 5.6, 3.1, 0, 12.4, 4.2, 0.8, 0)
    If Not tbl Is Nothing Then
        EnsureRows tbl, 14
        For i = 0 To 13
            tbl.DataBodyRange.Rows(i + 1) = Array(d + i, rain(i))
        Next i
    End If
End Sub

' ==== Chart Sheet ============================================================

Private Sub SetupChart()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(Schema.SHEET_CHART)
    ws.Range("A1") = "Simulation Charts": ws.Range("A1").Font.Bold = True
    ws.Range("A2") = "Run WQOC.Run to generate charts"
End Sub

' ==== Controls (Buttons, Dropdowns) =========================================

Private Sub SetupControls()
    Dim ws As Worksheet, btn As Shape
    Set ws = ThisWorkbook.Worksheets(Schema.SHEET_INPUT)

    ' Remove existing buttons
    On Error Resume Next
    ws.Shapes("btnRun").Delete
    ws.Shapes("btnRollback").Delete
    On Error GoTo 0

    ' Run button
    Set btn = ws.Shapes.AddShape(msoShapeRoundedRectangle, 400, 5, 80, 28)
    With btn
        .Name = "btnRun"
        .TextFrame2.TextRange.Text = "Run"
        .TextFrame2.TextRange.Font.Size = 11
        .TextFrame2.TextRange.Font.Bold = msoTrue
        .TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignCenter
        .Fill.ForeColor.RGB = RGB(112, 173, 71)
        .Line.Visible = msoFalse
        .OnAction = "WQOC.Run"
    End With

    ' Rollback button
    Set btn = ws.Shapes.AddShape(msoShapeRoundedRectangle, 490, 5, 80, 28)
    With btn
        .Name = "btnRollback"
        .TextFrame2.TextRange.Text = "Rollback"
        .TextFrame2.TextRange.Font.Size = 11
        .TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignCenter
        .Fill.ForeColor.RGB = RGB(191, 191, 191)
        .Line.Visible = msoFalse
        .OnAction = "WQOC.Rollback"
    End With

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
