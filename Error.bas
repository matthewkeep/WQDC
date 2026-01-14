Option Explicit
' Error: Centralized error handling and diagnostics.
' Dependencies: None

Public Const DEBUG_ON As Boolean = True  ' Toggle all logging

Public Sub Trace(ByVal src As String, ByVal msg As String)
    ' Logs diagnostic message to Immediate window
    If DEBUG_ON Then Debug.Print src & ": " & msg
End Sub

Public Sub TraceErr(ByVal src As String)
    ' Logs current error and clears it
    If Err.Number <> 0 Then
        If DEBUG_ON Then Debug.Print src & ": [" & Err.Number & "] " & Err.Description
        Err.Clear
    End If
End Sub

Public Sub Raise(ByVal src As String, ByVal msg As String)
    ' Logs and raises error for caller to handle
    Trace src, msg
    Err.Raise vbObjectError + 1, src, msg
End Sub
