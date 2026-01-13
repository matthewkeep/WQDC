Option Explicit
' Events: Worksheet event handlers.
' Dependencies: Loader, Schema
'
' NOTE: To enable auto-load on site selection, add this code to the
' Inputs sheet module (right-click sheet tab > View Code):
'
'   Private Sub Worksheet_Change(ByVal Target As Range)
'       Events.OnInputsChange Target
'   End Sub

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
