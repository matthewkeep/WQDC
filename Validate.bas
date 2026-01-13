Attribute VB_Name = "Validate"
Option Explicit
' Validate: Pre-flight checks for workbook structure.
' Purpose: Catch misconfigured workbooks before running simulation.
' Dependencies: Schema (constants only)
'
' Usage:
'   Validate.Check       - Returns True if valid, False if issues found
'   Validate.Report      - Shows detailed report of all issues
'
' This module is standalone and can be removed without affecting core functionality.

Private mIssues As Collection

' ==== Public Entry Points ====================================================

' Check workbook structure, return True if valid
Public Function Check() As Boolean
    Set mIssues = New Collection

    CheckSheets
    CheckNamedRanges
    CheckTables

    Check = (mIssues.Count = 0)

    If mIssues.Count > 0 Then
        Debug.Print "Validation failed: " & mIssues.Count & " issue(s) found"
    Else
        Debug.Print "Validation passed"
    End If
End Function

' Show detailed report of all issues
Public Sub Report()
    Dim i As Long

    If Not Check() Then
        Debug.Print ""
        Debug.Print "=== Validation Issues ==="
        For i = 1 To mIssues.Count
            Debug.Print "  " & i & ". " & mIssues(i)
        Next i
        Debug.Print ""
        Debug.Print "Run Setup.Build to create missing structure"
    End If
End Sub

' Quick check with message box result
Public Sub QuickCheck()
    If Check() Then
        MsgBox "Workbook structure is valid.", vbInformation, "Validate"
    Else
        MsgBox "Found " & mIssues.Count & " issue(s)." & vbNewLine & _
               "Run Validate.Report in Immediate Window for details.", _
               vbExclamation, "Validate"
    End If
End Sub

' ==== Sheet Checks ===========================================================

Private Sub CheckSheets()
    CheckSheet Schema.SHEET_INPUT, "Input sheet"
    CheckSheet Schema.SHEET_CONFIG, "Config sheet"
    CheckSheet Schema.SHEET_RESULTS, "Results sheet"
    CheckSheet Schema.SHEET_RAIN, "Rain sheet"
    CheckSheet Schema.SHEET_HISTORY, "History sheet"
End Sub

Private Sub CheckSheet(ByVal sheetName As String, ByVal desc As String)
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0
    If ws Is Nothing Then
        AddIssue "Missing sheet: " & sheetName & " (" & desc & ")"
    End If
End Sub

' ==== Named Range Checks =====================================================

Private Sub CheckNamedRanges()
    ' Core ranges
    CheckRange Schema.NAME_SITE, "Site selector"
    CheckRange Schema.NAME_INIT_VOL, "Initial volume"
    CheckRange Schema.NAME_TRIGGER_VOL, "Trigger volume"
    CheckRange Schema.NAME_SAMPLE_DATE, "Sample date"
    CheckRange Schema.NAME_RUN_DATE, "Run date"
    CheckRange Schema.NAME_OUTPUT, "Output flow"

    ' Chemistry ranges
    CheckRange Schema.NAME_RES_ROW, "Latest chemistry row"
    CheckRange Schema.NAME_LIMIT_ROW, "Trigger limits row"
    CheckRange Schema.NAME_HIDDEN_MASS, "Hidden mass"

    ' Config ranges
    CheckRange Schema.NAME_TAU, "Tau (mixing constant)"
    CheckRange Schema.NAME_RAIN_FACTOR, "Rain factor"
    CheckRange Schema.NAME_RAIN_MODE, "Rain mode"
    CheckRange Schema.NAME_SURFACE_FRACTION, "Surface fraction"
    CheckRange Schema.NAME_NET_OUT, "Net outflow"
    CheckRange Schema.NAME_ENHANCED_MODE, "Enhanced mode toggle"

    ' Output ranges
    CheckRange Schema.NAME_STD_TRIGGER, "Standard trigger result"
End Sub

Private Sub CheckRange(ByVal rangeName As String, ByVal desc As String)
    Dim rng As Range
    On Error Resume Next
    Set rng = ThisWorkbook.Names(rangeName).RefersToRange
    On Error GoTo 0
    If rng Is Nothing Then
        AddIssue "Missing named range: " & rangeName & " (" & desc & ")"
    End If
End Sub

' ==== Table Checks ===========================================================

Private Sub CheckTables()
    CheckTable Schema.SHEET_INPUT, Schema.TABLE_IR, "Inflow sources table"
    CheckTable Schema.SHEET_CONFIG, Schema.TABLE_CATALOG, "Site catalog table"
    CheckTable Schema.SHEET_CONFIG, Schema.TABLE_TRIGGER, "Trigger presets table"
    CheckTable Schema.SHEET_RESULTS, Schema.TABLE_RESULTS, "Lab results table"
    CheckTable Schema.SHEET_RAIN, Schema.TABLE_RAIN, "Rainfall table"
    CheckTable Schema.SHEET_HISTORY, Schema.TABLE_HISTORY, "Run history table"
End Sub

Private Sub CheckTable(ByVal sheetName As String, ByVal tableName As String, ByVal desc As String)
    Dim ws As Worksheet
    Dim tbl As ListObject

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(sheetName)
    If Not ws Is Nothing Then
        Set tbl = ws.ListObjects(tableName)
    End If
    On Error GoTo 0

    If tbl Is Nothing Then
        AddIssue "Missing table: " & tableName & " on " & sheetName & " (" & desc & ")"
    End If
End Sub

' ==== Helpers ================================================================

Private Sub AddIssue(ByVal msg As String)
    mIssues.Add msg
End Sub
