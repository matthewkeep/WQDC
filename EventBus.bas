Option Explicit
' EventBus: Centralized event dispatch for decoupled components.
' Dependencies: Data, Schema
'
' Usage:
'   EventBus.Notify EVENT_SAMPLE_DATE_CHANGED, site
'
' Benefits:
'   - Producers don't need to know about consumers
'   - Adding new handlers is one line in Dispatch
'   - Programmatic changes trigger same handlers as UI

' ==== Event Constants ===========================================================

Public Const EVENT_SAMPLE_DATE_CHANGED As String = "SampleDateChanged"
Public Const EVENT_SITE_CHANGED As String = "SiteChanged"
Public Const EVENT_ENHANCED_MODE_CHANGED As String = "EnhancedModeChanged"

' ==== Public Entry Point ========================================================

Public Sub Notify(ByVal eventName As String, Optional ByVal data As Variant)
    ' Dispatches event to all registered handlers
    ' data parameter carries context (e.g., site name, date value)
    On Error Resume Next
    Dispatch eventName, data
    On Error GoTo 0
End Sub

' ==== Private Dispatcher ========================================================

Private Sub Dispatch(ByVal eventName As String, ByVal data As Variant)
    ' Routes events to handlers - add new handlers here
    Select Case eventName
        Case EVENT_SAMPLE_DATE_CHANGED
            OnSampleDateChanged data

        Case EVENT_SITE_CHANGED
            OnSiteChanged data

        Case EVENT_ENHANCED_MODE_CHANGED
            OnEnhancedModeChanged data
    End Select
End Sub

' ==== Event Handlers ============================================================

Private Sub OnSampleDateChanged(ByVal site As Variant)
    ' Triggered when Sample Date changes
    ' Loads hidden mass from log for TwoBucket continuity
    Dim sampleDate As Date

    If IsMissing(site) Or Len(CStr(site)) = 0 Then site = Data.GetSite()
    If Len(CStr(site)) = 0 Then Exit Sub

    sampleDate = GetSampleDate()
    If sampleDate = 0 Then Exit Sub

    ' Load hidden state for the new sample date
    Data.LoadHiddenForDate CStr(site), sampleDate
End Sub

Private Sub OnSiteChanged(ByVal newSite As Variant)
    ' Triggered when site selection changes
    ' Could trigger IR population, telemetry column setup, etc.
    ' Currently placeholder for future handlers
    Error.Trace "EventBus", "Site changed to: " & CStr(newSite)
End Sub

Private Sub OnEnhancedModeChanged(ByVal enabled As Variant)
    ' Triggered when Enhanced mode toggled
    ' Could trigger UI updates, recalculation, etc.
    ' Currently placeholder for future handlers
    Error.Trace "EventBus", "Enhanced mode: " & CStr(enabled)
End Sub

' ==== Private Helpers ===========================================================

Private Function GetSampleDate() As Date
    ' Reads sample date from Inputs sheet
    Dim ws As Worksheet
    Set ws = Helpers.GetSheet(Schema.SHEET_INPUT)
    If Not ws Is Nothing Then
        GetSampleDate = Helpers.GetDateVal(ws, Schema.NAME_SAMPLE_DATE)
    End If
End Function
