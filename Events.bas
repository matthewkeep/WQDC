Option Explicit
' Events: Worksheet event handlers.
' Dependencies: Loader, Schema, WQOC, History
'
' NOTE: To enable events, add this code to each sheet module
' (right-click sheet tab > View Code):
'
' === Inputs sheet ===
'   Private Sub Worksheet_Change(ByVal Target As Range)
'       Events.OnInputsChange Target
'   End Sub
'   Private Sub Worksheet_BeforeDoubleClick(ByVal Target As Range, Cancel As Boolean)
'       Events.OnInputsDoubleClick Target, Cancel
'   End Sub
'
' === RunHistory sheet ===
'   Private Sub Worksheet_BeforeDoubleClick(ByVal Target As Range, Cancel As Boolean)
'       Events.OnHistoryDoubleClick Target, Cancel
'   End Sub

' ==== Change Events ============================================================

Public Sub OnInputsChange(ByVal Target As Range)
    ' Called from Inputs sheet Worksheet_Change event
    Dim siteRng As Range
    On Error Resume Next
    Set siteRng = Target.Worksheet.Range(Schema.NAME_SITE)
    On Error GoTo 0
    If siteRng Is Nothing Then Exit Sub
    If Not Intersect(Target, siteRng) Is Nothing Then
        Loader.LoadSiteData CStr(Target.Value)
    End If
End Sub

' ==== Double-Click Events ======================================================

Public Sub OnInputsDoubleClick(ByVal Target As Range, ByRef Cancel As Boolean)
    ' Handle double-clicks on Inputs sheet
    Dim ws As Worksheet, runCell As Range, tbl As ListObject
    Dim actionCol As Long, rowIdx As Long

    Set ws = Target.Worksheet

    ' Check Run Simulation cell
    On Error Resume Next
    Set runCell = ws.Range(Schema.NAME_RUN_CELL)
    On Error GoTo 0
    If Not runCell Is Nothing Then
        If Not Intersect(Target, runCell) Is Nothing Then
            Cancel = True
            WQOC.Run
            Exit Sub
        End If
    End If

    ' Check IR table action column
    Set tbl = GetTable(ws, Schema.TABLE_IR)
    If Not tbl Is Nothing Then
        actionCol = GetColIndex(tbl, Schema.IR_COL_ACTION)
        If actionCol > 0 Then
            ' Check header (Add)
            If Not Intersect(Target, tbl.HeaderRowRange.Cells(1, actionCol)) Is Nothing Then
                Cancel = True
                AddIRRow tbl
                Exit Sub
            End If
            ' Check data rows (Remove)
            If Not tbl.DataBodyRange Is Nothing Then
                If Not Intersect(Target, tbl.DataBodyRange.Columns(actionCol)) Is Nothing Then
                    Cancel = True
                    rowIdx = Target.Row - tbl.DataBodyRange.Row + 1
                    RemoveIRRow tbl, rowIdx
                    Exit Sub
                End If
            End If
        End If
    End If
End Sub

Public Sub OnHistoryDoubleClick(ByVal Target As Range, ByRef Cancel As Boolean)
    ' Handle double-clicks on History sheet (per-site tables)
    Dim ws As Worksheet, tbl As ListObject, lo As ListObject
    Dim actionCol As Long, rowIdx As Long, runId As String, site As String

    Set ws = Target.Worksheet

    ' Find which table was clicked (if any)
    Set tbl = Nothing
    For Each lo In ws.ListObjects
        If Left$(lo.Name, Len(Schema.HISTORY_TABLE_PREFIX)) = Schema.HISTORY_TABLE_PREFIX Then
            If Not lo.DataBodyRange Is Nothing Then
                If Not Intersect(Target, lo.Range) Is Nothing Then
                    Set tbl = lo
                    Exit For
                End If
            End If
        End If
    Next lo
    If tbl Is Nothing Then Exit Sub
    If tbl.DataBodyRange Is Nothing Then Exit Sub

    ' Extract site from table name (e.g., "tblHistory_RP1" -> "RP1")
    site = Mid$(tbl.Name, Len(Schema.HISTORY_TABLE_PREFIX) + 1)

    actionCol = GetColIndex(tbl, Schema.HISTORY_COL_ACTION)
    If actionCol = 0 Then Exit Sub

    ' Check if clicked in action column data area
    If Not Intersect(Target, tbl.DataBodyRange.Columns(actionCol)) Is Nothing Then
        Cancel = True
        rowIdx = Target.Row - tbl.DataBodyRange.Row + 1
        runId = tbl.DataBodyRange.Cells(rowIdx, 1).Value  ' RunId is column 1

        ' Don't rollback the most recent (Current) row
        If rowIdx = tbl.ListRows.Count Then
            MsgBox "This is the current run.", vbInformation, "WQOC"
            Exit Sub
        End If

        If MsgBox("Rollback to run " & runId & "?" & vbNewLine & _
                  "This will remove all runs after this one.", vbYesNo + vbQuestion, "WQOC") = vbYes Then
            History.RollbackTo runId, site
            RefreshHistoryActions tbl
        End If
    End If
End Sub

' ==== IR Table Actions =========================================================

Private Sub AddIRRow(ByVal tbl As ListObject)
    ' Add a new empty row to IR table with "Remove" action
    Dim newRow As ListRow, actionCol As Long
    Set newRow = tbl.ListRows.Add
    actionCol = GetColIndex(tbl, Schema.IR_COL_ACTION)
    If actionCol > 0 Then
        newRow.Range.Cells(1, actionCol).Value = Schema.ACTION_REMOVE
        StyleActionCell newRow.Range.Cells(1, actionCol)
    End If
End Sub

Private Sub RemoveIRRow(ByVal tbl As ListObject, ByVal rowIdx As Long)
    ' Remove a row from IR table
    If tbl.ListRows.Count = 1 Then
        ' Don't delete last row, just clear it
        tbl.DataBodyRange.ClearContents
        tbl.DataBodyRange.Cells(1, GetColIndex(tbl, Schema.IR_COL_ACTION)).Value = Schema.ACTION_REMOVE
    Else
        tbl.ListRows(rowIdx).Delete
    End If
End Sub

Private Sub RefreshHistoryActions(ByVal tbl As ListObject)
    ' Update action column text: "Current" for last row, "Rollback" for others
    Dim i As Long, actionCol As Long
    If tbl.DataBodyRange Is Nothing Then Exit Sub
    actionCol = GetColIndex(tbl, Schema.HISTORY_COL_ACTION)
    If actionCol = 0 Then Exit Sub

    For i = 1 To tbl.ListRows.Count
        If i = tbl.ListRows.Count Then
            tbl.DataBodyRange.Cells(i, actionCol).Value = Schema.ACTION_CURRENT
        Else
            tbl.DataBodyRange.Cells(i, actionCol).Value = Schema.ACTION_ROLLBACK
        End If
        StyleActionCell tbl.DataBodyRange.Cells(i, actionCol)
    Next i
End Sub

' ==== Helpers ==================================================================

Private Function GetTable(ByVal ws As Worksheet, ByVal nm As String) As ListObject
    On Error Resume Next
    Set GetTable = ws.ListObjects(nm)
    On Error GoTo 0
End Function

Private Function GetColIndex(ByVal tbl As ListObject, ByVal colName As String) As Long
    Dim col As ListColumn
    On Error Resume Next
    Set col = tbl.ListColumns(colName)
    If Not col Is Nothing Then GetColIndex = col.Index
    On Error GoTo 0
End Function

Private Sub StyleActionCell(ByVal cell As Range)
    With cell
        .Font.Color = Schema.COLOR_ACTION_FONT
        .Font.Underline = xlUnderlineStyleSingle
    End With
End Sub
