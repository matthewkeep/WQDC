Option Explicit
' Utils: Shared utility functions.
' Purpose: Common helpers used across multiple modules
' Dependencies: None

' ==== Table Helpers ============================================================

Public Function ColIdx(ByVal tbl As ListObject, ByVal colName As String) As Long
    ' Returns column index (1-based) or 0 if not found
    Dim col As ListColumn
    On Error Resume Next
    Set col = tbl.ListColumns(colName)
    If Not col Is Nothing Then ColIdx = col.Index
    On Error GoTo 0
End Function
