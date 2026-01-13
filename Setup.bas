Attribute VB_Name = "Setup"
Option Explicit
' Setup: Standalone workbook scaffolding and test data seeding.
' Purpose: Create sheets, tables, named ranges, and seed test data for testing.
' Dependencies: Schema (constants only)
'
' Usage:
'   Setup.Build      - Create full workbook structure
'   Setup.Seed       - Populate with test data
'   Setup.BuildAll   - Build + Seed in one call
'   Setup.Clean      - Remove all WQOC sheets (reset)
'
' This module is standalone and can be removed after testing.

' ==== Public Entry Points ====================================================

' Build complete workbook structure
Public Sub Build()
    Dim calcMode As XlCalculation

    On Error GoTo HandleError
    calcMode = Application.Calculation
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual

    CreateSheets
    SetupInputSheet
    SetupConfigSheet
    SetupResultsSheet
    SetupRainSheet
    SetupHistorySheet

    Application.Calculation = calcMode
    Application.ScreenUpdating = True
    Application.EnableEvents = True

    MsgBox "Workbook structure created.", vbInformation, "Setup"
    Exit Sub

HandleError:
    Application.Calculation = calcMode
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    MsgBox "Build error: " & Err.Description, vbExclamation, "Setup"
End Sub

' Seed test data (assumes structure exists)
Public Sub Seed()
    Dim calcMode As XlCalculation

    On Error GoTo HandleError
    calcMode = Application.Calculation
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual

    SeedInputData
    SeedConfigData
    SeedResultsData
    SeedRainData

    Application.Calculation = calcMode
    Application.ScreenUpdating = True
    Application.EnableEvents = True

    MsgBox "Test data seeded.", vbInformation, "Setup"
    Exit Sub

HandleError:
    Application.Calculation = calcMode
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    MsgBox "Seed error: " & Err.Description, vbExclamation, "Setup"
End Sub

' Build + Seed in one call
Public Sub BuildAll()
    Build
    Seed
End Sub

' Remove all WQOC sheets (clean slate)
Public Sub Clean()
    Dim ws As Worksheet
    Dim sheetNames As Variant
    Dim i As Long

    sheetNames = Array(Schema.SHEET_INPUT, Schema.SHEET_CONFIG, Schema.SHEET_RESULTS, _
                       Schema.SHEET_RAIN, Schema.SHEET_HISTORY, Schema.SHEET_CHART, Schema.SHEET_LOG)

    Application.DisplayAlerts = False
    For i = LBound(sheetNames) To UBound(sheetNames)
        On Error Resume Next
        Set ws = ThisWorkbook.Worksheets(sheetNames(i))
        If Not ws Is Nothing Then ws.Delete
        Set ws = Nothing
        On Error GoTo 0
    Next i
    Application.DisplayAlerts = True

    ' Remove named ranges
    On Error Resume Next
    Dim nm As Name
    For Each nm In ThisWorkbook.Names
        nm.Delete
    Next nm
    On Error GoTo 0

    MsgBox "All WQOC sheets removed.", vbInformation, "Setup"
End Sub

' ==== Sheet Creation =========================================================

Private Sub CreateSheets()
    EnsureSheet Schema.SHEET_INPUT
    EnsureSheet Schema.SHEET_CONFIG
    EnsureSheet Schema.SHEET_RESULTS
    EnsureSheet Schema.SHEET_RAIN
    EnsureSheet Schema.SHEET_HISTORY
End Sub

Private Sub EnsureSheet(ByVal sheetName As String)
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        ws.Name = sheetName
    End If
    ws.Cells.Clear
End Sub

' ==== Input Sheet Setup ======================================================

Private Sub SetupInputSheet()
    Dim ws As Worksheet
    Dim chemNames As Variant
    Dim chemCount As Long
    Dim i As Long

    Set ws = ThisWorkbook.Worksheets(Schema.SHEET_INPUT)
    chemNames = Schema.ChemistryNames()
    chemCount = Schema.ChemistryCount()

    ' === Reservoir Summary Block (A1:I5) ===
    ws.Range("A1").Value = "Reservoir Summary"
    ws.Range("A1:L1").Font.Bold = True
    ws.Range("A1:L1").Interior.Color = Schema.COLOR_INPUT_RES_FILL

    ' Headers row
    ws.Range("B2").Value = Schema.VOLUME_METRIC_NAME
    For i = 0 To chemCount - 1
        ws.Cells(2, 3 + i).Value = chemNames(i)
    Next i
    ws.Range("A2:I2").Font.Bold = True

    ' Row labels
    ws.Range("A3").Value = "Latest WQ"
    ws.Range("A4").Value = "Trigger"
    ws.Range("A5").Value = "Predicted"
    ws.Range("A3:A5").Font.Bold = True

    ' Row colors
    ws.Range("A3:I3").Interior.Color = Schema.COLOR_LATEST_FILL
    ws.Range("A4:I4").Interior.Color = Schema.COLOR_TRIGGER_FILL
    ws.Range("A5:I5").Interior.Color = Schema.COLOR_PREDICTED_FILL

    ' Named ranges for reservoir data
    AddName Schema.NAME_INIT_VOL, ws.Range("B3")
    AddName Schema.NAME_TRIGGER_VOL, ws.Range("B4")
    AddName Schema.NAME_TRIGGER_RESULT_VOL, ws.Range("B5")
    AddName Schema.NAME_RES_ROW, ws.Range("C3").Resize(1, chemCount)
    AddName Schema.NAME_LIMIT_ROW, ws.Range("C4").Resize(1, chemCount)

    ' === Run Info Block (J2:L5) ===
    ws.Range("J2").Value = "Run Date": ws.Range("J2").Font.Bold = True
    AddName Schema.NAME_RUN_DATE, ws.Range("K2")
    ws.Range("J3").Value = "Site": ws.Range("J3").Font.Bold = True
    AddName Schema.NAME_SITE, ws.Range("K3")
    ws.Range("J4").Value = "Output (ML/d)": ws.Range("J4").Font.Bold = True
    AddName Schema.NAME_OUTPUT, ws.Range("K4")
    ws.Range("J5").Value = "Sample Date": ws.Range("J5").Font.Bold = True
    AddName Schema.NAME_SAMPLE_DATE, ws.Range("K5")

    ' === Results Block (N1:P4) ===
    ws.Range("N1").Value = "Last Run Results"
    ws.Range("N1:P1").Font.Bold = True
    ws.Range("N1:P1").Interior.Color = Schema.COLOR_INPUT_RUN_FILL
    ws.Range("N2").Value = "Std Trigger": ws.Range("N2").Font.Bold = True
    AddName Schema.NAME_STD_TRIGGER, ws.Range("O2")
    ws.Range("N3").Value = "Enh Trigger": ws.Range("N3").Font.Bold = True
    AddName Schema.NAME_ENH_TRIGGER, ws.Range("O3")
    ws.Range("N4").Value = "Mode": ws.Range("N4").Font.Bold = True
    AddName Schema.NAME_ENHANCED_MODE, ws.Range("O4")

    ' === Calibration Block (N6:O10) ===
    ws.Range("N6").Value = "Calibration"
    ws.Range("N6:O6").Font.Bold = True
    ws.Range("N6:O6").Interior.Color = Schema.COLOR_INPUT_CALIB_FILL
    ws.Range("N7").Value = "Tau (days)": ws.Range("N7").Font.Bold = True
    AddName Schema.NAME_TAU, ws.Range("O7")
    ws.Range("N8").Value = "Rain Factor": ws.Range("N8").Font.Bold = True
    AddName Schema.NAME_RAIN_FACTOR, ws.Range("O8")
    ws.Range("N9").Value = "Rain Mode": ws.Range("N9").Font.Bold = True
    AddName Schema.NAME_RAIN_MODE, ws.Range("O9")
    ws.Range("N10").Value = "Surface Frac": ws.Range("N10").Font.Bold = True
    AddName Schema.NAME_SURFACE_FRACTION, ws.Range("O10")
    ws.Range("N11").Value = "Net Outflow": ws.Range("N11").Font.Bold = True
    AddName Schema.NAME_NET_OUT, ws.Range("O11")

    ' === Hidden Mass Block (Q6:R13) ===
    ws.Range("Q6").Value = "Hidden Mass"
    ws.Range("Q6:R6").Font.Bold = True
    ws.Range("Q6:R6").Interior.Color = Schema.COLOR_INPUT_HIDDEN_FILL
    For i = 0 To chemCount - 1
        ws.Cells(7 + i, 17).Value = chemNames(i)
        ws.Cells(7 + i, 17).Font.Bold = True
    Next i
    AddName Schema.NAME_HIDDEN_MASS, ws.Range("R7").Resize(chemCount, 1)

    ' === IR Table (A8) ===
    ws.Range("A8").Value = "Inflow Sources"
    ws.Range("A8:L8").Font.Bold = True
    ws.Range("A8:L8").Interior.Color = Schema.COLOR_INPUT_RES_FILL
    CreateIRTable ws, chemNames, chemCount
End Sub

Private Sub CreateIRTable(ByVal ws As Worksheet, ByVal chemNames As Variant, ByVal chemCount As Long)
    Dim headers() As String
    Dim colCount As Long
    Dim tbl As ListObject
    Dim i As Long

    ' Headers: Source, Flow, [Chemistry x7], Sample Date, Active
    colCount = 3 + chemCount + 2
    ReDim headers(1 To colCount)
    headers(1) = Schema.IR_COL_SOURCE
    headers(2) = Schema.IR_COL_FLOW
    For i = 0 To chemCount - 1
        headers(3 + i) = chemNames(i)
    Next i
    headers(chemCount + 3) = Schema.IR_COL_SAMPLE_DATE
    headers(chemCount + 4) = Schema.IR_COL_ACTIVE

    Set tbl = CreateTable(ws, ws.Range("A9"), Schema.TABLE_IR, headers)
    If Not tbl Is Nothing Then
        tbl.ListColumns(chemCount + 3).Range.NumberFormat = "dd/mm/yyyy"
    End If
End Sub

' ==== Config Sheet Setup =====================================================

Private Sub SetupConfigSheet()
    Dim ws As Worksheet
    Dim chemNames As Variant
    Dim chemCount As Long
    Dim triggerHeaders() As String
    Dim catalogHeaders As Variant
    Dim i As Long

    Set ws = ThisWorkbook.Worksheets(Schema.SHEET_CONFIG)
    chemNames = Schema.ChemistryNames()
    chemCount = Schema.ChemistryCount()

    ' Catalog table (site mappings)
    ws.Range("A1").Value = "Site Catalog"
    ws.Range("A1").Font.Bold = True
    catalogHeaders = Array("RR", "IR", "Flow (ML/d)")
    CreateTable ws, ws.Range("A2"), Schema.TABLE_CATALOG, catalogHeaders

    ' Trigger presets table
    ws.Range("E1").Value = "Trigger Presets"
    ws.Range("E1").Font.Bold = True
    ReDim triggerHeaders(1 To chemCount + 2)
    triggerHeaders(1) = "Preset"
    triggerHeaders(2) = Schema.VOLUME_METRIC_NAME
    For i = 1 To chemCount
        triggerHeaders(2 + i) = chemNames(i - 1)
    Next i
    CreateTable ws, ws.Range("E2"), Schema.TABLE_TRIGGER, triggerHeaders
End Sub

' ==== Results Sheet Setup ====================================================

Private Sub SetupResultsSheet()
    Dim ws As Worksheet
    Dim chemNames As Variant
    Dim chemCount As Long
    Dim headers() As String
    Dim i As Long

    Set ws = ThisWorkbook.Worksheets(Schema.SHEET_RESULTS)
    chemNames = Schema.ChemistryNames()
    chemCount = Schema.ChemistryCount()

    ws.Range("A1").Value = "Lab Results"
    ws.Range("A1").Font.Bold = True

    ReDim headers(1 To chemCount + 3)
    headers(1) = "Site"
    headers(2) = "Sample Date"
    headers(3) = "Sample ID"
    For i = 1 To chemCount
        headers(3 + i) = chemNames(i - 1)
    Next i

    Dim tbl As ListObject
    Set tbl = CreateTable(ws, ws.Range("A2"), Schema.TABLE_RESULTS, headers)
    If Not tbl Is Nothing Then
        tbl.ListColumns(2).Range.NumberFormat = "dd/mm/yyyy"
    End If
End Sub

' ==== Rain Sheet Setup =======================================================

Private Sub SetupRainSheet()
    Dim ws As Worksheet
    Dim headers As Variant

    Set ws = ThisWorkbook.Worksheets(Schema.SHEET_RAIN)

    ws.Range("A1").Value = "Rainfall Data"
    ws.Range("A1").Font.Bold = True

    headers = Array("Date", "Rain (mm)")
    Dim tbl As ListObject
    Set tbl = CreateTable(ws, ws.Range("A2"), Schema.TABLE_RAIN, headers)
    If Not tbl Is Nothing Then
        tbl.ListColumns(1).Range.NumberFormat = "dd/mm/yyyy"
    End If
End Sub

' ==== History Sheet Setup ====================================================

Private Sub SetupHistorySheet()
    Dim ws As Worksheet
    Dim headers As Variant

    Set ws = ThisWorkbook.Worksheets(Schema.SHEET_HISTORY)

    ws.Range("A1").Value = "Run History"
    ws.Range("A1").Font.Bold = True

    headers = Array("RunId", "Timestamp", "RunDate", "Site", "SampleDate", "Mode", "TriggerDay", "TriggerMetric", "Status")
    Dim tbl As ListObject
    Set tbl = CreateTable(ws, ws.Range("A2"), Schema.TABLE_HISTORY, headers)
    If Not tbl Is Nothing Then
        tbl.ListColumns(2).Range.NumberFormat = "dd/mm/yyyy hh:mm"
        tbl.ListColumns(3).Range.NumberFormat = "dd/mm/yyyy"
        tbl.ListColumns(5).Range.NumberFormat = "dd/mm/yyyy"
    End If
End Sub

' ==== Test Data Seeding ======================================================

Private Sub SeedInputData()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(Schema.SHEET_INPUT)

    ' Run info
    ws.Range(Schema.NAME_RUN_DATE).Value = Date
    ws.Range(Schema.NAME_SITE).Value = "RP1"
    ws.Range(Schema.NAME_OUTPUT).Value = 2
    ws.Range(Schema.NAME_SAMPLE_DATE).Value = Date - 5

    ' Reservoir state
    ws.Range(Schema.NAME_INIT_VOL).Value = 150
    ws.Range(Schema.NAME_TRIGGER_VOL).Value = 200

    ' Latest chemistry (7 values)
    ws.Range(Schema.NAME_RES_ROW).Value = Array(280, 45, 2.5, 15, 6, 10, 0.1)

    ' Trigger limits (7 values)
    ws.Range(Schema.NAME_LIMIT_ROW).Value = Array(450, 100, 4, 22, 9, 15, 0.2)

    ' Calibration
    ws.Range(Schema.NAME_TAU).Value = 7
    ws.Range(Schema.NAME_RAIN_FACTOR).Value = 3.2
    ws.Range(Schema.NAME_RAIN_MODE).Value = "Typical"
    ws.Range(Schema.NAME_SURFACE_FRACTION).Value = 0.8
    ws.Range(Schema.NAME_NET_OUT).Value = 1
    ws.Range(Schema.NAME_ENHANCED_MODE).Value = "On"

    ' Hidden mass (zeros initially)
    Dim hiddenVals(1 To 7, 1 To 1) As Double
    ws.Range(Schema.NAME_HIDDEN_MASS).Value = hiddenVals

    ' IR table data
    Dim tbl As ListObject
    Set tbl = ws.ListObjects(Schema.TABLE_IR)
    If Not tbl Is Nothing Then
        EnsureRows tbl, 3
        tbl.DataBodyRange.Rows(1).Value = Array("CB1", 1.5, 350, 60, 2.8, 16, 7, 11, 0.12, Date - 3, "Yes")
        tbl.DataBodyRange.Rows(2).Value = Array("CB2", 0.8, 280, 40, 2.2, 14, 5.5, 9, 0.08, Date - 4, "Yes")
        tbl.DataBodyRange.Rows(3).Value = Array("TNWS", 0.5, 420, 80, 3.1, 18, 8, 12, 0.15, Date - 2, "No")
    End If
End Sub

Private Sub SeedConfigData()
    Dim ws As Worksheet
    Dim tblCatalog As ListObject
    Dim tblTrigger As ListObject

    Set ws = ThisWorkbook.Worksheets(Schema.SHEET_CONFIG)

    ' Catalog (site mappings)
    Set tblCatalog = ws.ListObjects(Schema.TABLE_CATALOG)
    If Not tblCatalog Is Nothing Then
        EnsureRows tblCatalog, 3
        tblCatalog.DataBodyRange.Rows(1).Value = Array("RP1", "CB1", 1.5)
        tblCatalog.DataBodyRange.Rows(2).Value = Array("RP1", "CB2", 0.8)
        tblCatalog.DataBodyRange.Rows(3).Value = Array("GCMBL", "TSFS2", 1.2)
    End If

    ' Trigger presets
    Set tblTrigger = ws.ListObjects(Schema.TABLE_TRIGGER)
    If Not tblTrigger Is Nothing Then
        EnsureRows tblTrigger, 3
        tblTrigger.DataBodyRange.Rows(1).Value = Array("L1", 210, 350, 50, 3.5, 18, 7.5, 12, 0.15)
        tblTrigger.DataBodyRange.Rows(2).Value = Array("L2", 200, 450, 100, 4.2, 22, 9, 15, 0.18)
        tblTrigger.DataBodyRange.Rows(3).Value = Array("L3", 220, 400, 75, 3.8, 20, 8, 13, 0.16)
    End If
End Sub

Private Sub SeedResultsData()
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim baseDate As Date

    Set ws = ThisWorkbook.Worksheets(Schema.SHEET_RESULTS)
    Set tbl = ws.ListObjects(Schema.TABLE_RESULTS)
    baseDate = Date - 10

    If Not tbl Is Nothing Then
        EnsureRows tbl, 5
        tbl.DataBodyRange.Rows(1).Value = Array("RP1", baseDate, "RP1-001", 280, 45, 2.5, 15, 6, 10, 0.1)
        tbl.DataBodyRange.Rows(2).Value = Array("RP1", baseDate + 3, "RP1-002", 290, 48, 2.6, 15.5, 6.2, 10.2, 0.11)
        tbl.DataBodyRange.Rows(3).Value = Array("CB1", baseDate + 1, "CB1-001", 350, 60, 2.8, 16, 7, 11, 0.12)
        tbl.DataBodyRange.Rows(4).Value = Array("CB2", baseDate + 2, "CB2-001", 280, 40, 2.2, 14, 5.5, 9, 0.08)
        tbl.DataBodyRange.Rows(5).Value = Array("GCMBL", baseDate + 4, "GCMBL-001", 320, 55, 2.7, 17, 6.8, 10.8, 0.13)
    End If
End Sub

Private Sub SeedRainData()
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim baseDate As Date
    Dim i As Long
    Dim rainAmounts As Variant

    Set ws = ThisWorkbook.Worksheets(Schema.SHEET_RAIN)
    Set tbl = ws.ListObjects(Schema.TABLE_RAIN)
    baseDate = Date - 14
    rainAmounts = Array(0, 2.5, 0, 8.3, 1.2, 0, 0, 5.6, 3.1, 0, 12.4, 4.2, 0.8, 0)

    If Not tbl Is Nothing Then
        EnsureRows tbl, 14
        For i = 0 To 13
            tbl.DataBodyRange.Rows(i + 1).Value = Array(baseDate + i, rainAmounts(i))
        Next i
    End If
End Sub

' ==== Helper Functions =======================================================

Private Sub AddName(ByVal rangeName As String, ByVal target As Range)
    On Error Resume Next
    ThisWorkbook.Names(rangeName).Delete
    On Error GoTo 0
    ThisWorkbook.Names.Add Name:=rangeName, RefersTo:="=" & target.Address(True, True, xlA1, True)
End Sub

Private Function CreateTable(ByVal ws As Worksheet, ByVal startCell As Range, _
                             ByVal tableName As String, ByVal headers As Variant) As ListObject
    Dim colCount As Long
    Dim tbl As ListObject
    Dim headerRow As Range
    Dim tblRange As Range

    ' Remove existing table if present
    On Error Resume Next
    ws.ListObjects(tableName).Delete
    On Error GoTo 0

    ' Determine column count
    If IsArray(headers) Then
        colCount = UBound(headers) - LBound(headers) + 1
    Else
        colCount = 1
    End If

    ' Write headers
    Set headerRow = startCell.Resize(1, colCount)
    headerRow.Value = headers
    headerRow.Font.Bold = True

    ' Create table with one data row
    Set tblRange = startCell.Resize(2, colCount)
    Set tbl = ws.ListObjects.Add(xlSrcRange, tblRange, , xlYes)
    tbl.Name = tableName

    On Error Resume Next
    tbl.TableStyle = Schema.TABLE_STYLE_DEFAULT
    On Error GoTo 0

    Set CreateTable = tbl
End Function

Private Sub EnsureRows(ByVal tbl As ListObject, ByVal rowCount As Long)
    ' Ensure table has at least rowCount data rows
    Do While tbl.ListRows.Count < rowCount
        tbl.ListRows.Add
    Loop
    ' Clear existing data
    If Not tbl.DataBodyRange Is Nothing Then
        tbl.DataBodyRange.ClearContents
    End If
End Sub
