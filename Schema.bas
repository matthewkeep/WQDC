Option Explicit
' Schema: Constants and configuration.
' Dependencies: None

' ==== Sheet Names ============================================================
Public Const SHEET_INPUT As String = "Inputs"
Public Const SHEET_LOG As String = "Log"
Public Const SHEET_CHART As String = "Chart"
Public Const SHEET_TELEMETRY As String = "Telemetry"
Public Const SHEET_RESULTS As String = "Results"
Public Const SHEET_CONFIG As String = "Config"
Public Const SHEET_HISTORY As String = "RunHistory"

' ==== Named Ranges ===========================================================
Public Const NAME_SITE As String = "RR_Site"
Public Const NAME_OUTPUT As String = "RR_Output"
Public Const NAME_INIT_VOL As String = "RR_InitVol"
Public Const NAME_TRIGGER_VOL As String = "RR_TriggerVol"
Public Const NAME_SAMPLE_DATE As String = "RR_SampleDate"
Public Const NAME_RUN_DATE As String = "Run_Date"
Public Const NAME_TAU As String = "Cfg_Tau"
Public Const NAME_SURFACE_FRACTION As String = "Cfg_SurfaceFrac"
Public Const NAME_LIMIT_ROW As String = "Limit_Row"
Public Const NAME_RES_ROW As String = "Res_Row"
Public Const NAME_ENHANCED_MODE As String = "Cfg_EnhancedMode"
Public Const NAME_STD_TRIGGER As String = "Std_Trigger"
Public Const NAME_ENH_TRIGGER As String = "Enh_Trigger"
Public Const NAME_RESULT_VOL As String = "Result_Vol"
Public Const NAME_PRED_ROW As String = "Pred_Row"
Public Const NAME_HIDDEN_MASS As String = "RR_HiddenMass"
Public Const NAME_MIXING_MODEL As String = "Cfg_MixingModel"
Public Const NAME_RAINFALL_MODE As String = "Cfg_RainfallMode"
Public Const NAME_TELEM_CAL As String = "Cfg_TelemCal"

' ==== Table Names ============================================================
Public Const TABLE_IR As String = "tblIR"
Public Const TABLE_TELEMETRY As String = "tblTelemetry"
Public Const TABLE_RESULTS As String = "tblResults"
Public Const TABLE_CATALOG As String = "tblCatalog"
Public Const TABLE_TRIGGER As String = "tblTrigger"

' Per-site table prefixes (tables created on-demand)
Public Const LIVE_TABLE_PREFIX As String = "tblLive_"
Public Const HISTORY_TABLE_PREFIX As String = "tblHistory_"

' Live table columns (date-centric log with Std/Enh side-by-side)
Public Const LIVE_COL_DATE As String = "Date"
Public Const LIVE_COL_STD_VOL As String = "StdVol"
Public Const LIVE_COL_STD_EC As String = "StdEC"
Public Const LIVE_COL_ENH_VOL As String = "EnhVol"
Public Const LIVE_COL_ENH_EC As String = "EnhEC"
Public Const LIVE_COL_ERR_VOL As String = "ErrVol"
Public Const LIVE_COL_ERR_EC As String = "ErrEC"
Public Const LIVE_COL_RUNID As String = "RunId"
' Note: EnhHid1-7 columns are chemistry-based, built dynamically

' ==== Column Names ===========================================================
' IR table columns
Public Const IR_COL_SOURCE As String = "Source"
Public Const IR_COL_FLOW As String = "Flow (ML/d)"
Public Const IR_COL_ACTIVE As String = "Active"
Public Const IR_COL_SAMPLE_DATE As String = "Sample Date"
Public Const IR_COL_ACTION As String = "Add Input"

' History table columns
Public Const HISTORY_COL_ACTION As String = "Action"

' Telemetry columns (Date and Rain are fixed; EC/Vol are per-site)
Public Const TELEM_COL_DATE As String = "Date"
Public Const TELEM_COL_RAIN As String = "Rain (mm)"

' Volume metric name
Public Const VOLUME_METRIC_NAME As String = "Volume (ML)"

' ==== Action Cell Constants ==================================================
Public Const NAME_RUN_CELL As String = "Run_Simulation"
Public Const ACTION_ADD As String = "Add"
Public Const ACTION_REMOVE As String = "Remove"
Public Const ACTION_ROLLBACK As String = "Rollback"
Public Const ACTION_CURRENT As String = "Current"

' ==== Color Constants ========================================================
' Action/hyperlink colors
Public Const COLOR_ACTION_FONT As Long = &HC16305    ' #0563C1 - Blue hyperlink

' Chart colors (used by WQOC.GenerateCharts)
Public Const COLOR_STD_LINE As Long = &HB3712D       ' #2D71B3 - Standard line
Public Const COLOR_ENH_LINE As Long = &H779900       ' #009977 - Enhanced line
Public Const COLOR_TRIGGER_LINE As Long = &H0000C0   ' #C00000 - Trigger threshold

' Button colors (used by Setup.SetupControls)
Public Const COLOR_BUTTON_ON As Long = &H47AD70      ' #70AD47 - Button active

' Font colors
Public Const COLOR_FONT_WHITE As Long = &HFFFFFF     ' #FFFFFF - White text

' Log row colors
Public Const COLOR_SAMPLE_DATE As Long = &HFFFFCC    ' #CCFFFF - Light cyan for sample date row

' ==== Table Styles ===========================================================
Public Const TABLE_STYLE_DEFAULT As String = "TableStyleMedium2"
Public Const TABLE_GAP_COLS As Long = 2  ' Empty columns between horizontal tables

' ==== Simulation Defaults ====================================================
Public Const MAX_IR As Long = 10  ' Maximum number of IR (inflow) sources
Public Const DEFAULT_FORECAST_DAYS As Long = 100  ' Default forecast horizon (days)
Public Const DEFAULT_SURFACE_FRACTION As Double = 0.8

' ==== Enhanced Mode Options ==================================================
Public Const MIXING_SIMPLE As String = "Simple"
Public Const MIXING_TWOBUCKET As String = "TwoBucket"
Public Const MIXING_MODEL_LIST As String = "Simple,TwoBucket"

Public Const RAINFALL_OFF As String = "Off"
Public Const RAINFALL_HINDCAST As String = "Hindcast"
Public Const RAINFALL_FULL As String = "Hindcast+Forecast"
Public Const RAINFALL_MODE_LIST As String = "Off,Hindcast,Hindcast+Forecast"

Public Const TELEM_CAL_OFF As String = "Off"
Public Const TELEM_CAL_ON As String = "On"
Public Const TELEM_CAL_LIST As String = "Off,On"

' ==== Chart Layout ===========================================================
Public Const CHART_LEFT_POS As Double = 20
Public Const CHART_TOP_START As Double = 20
Public Const CHART_WIDTH As Double = 820
Public Const CHART_HEIGHT_VOLUME As Double = 260
Public Const CHART_HEIGHT_METRIC As Double = 260
Public Const CHART_SPACING As Double = 24

' ==== Chemistry Metrics ======================================================
Private mChemistryNames As Variant

Private Sub EnsureChemistryNames()
    If IsEmpty(mChemistryNames) Then
        ' 7 chemistry metrics (excludes Volume)
        mChemistryNames = Array("EC (uS/cm)", "F_U (ug/L)", "F_Mn (ug/L)", "SO4 (mg/L)", "Mg (mg/L)", "Ca (mg/L)", "TAN (mg/L)")
    End If
End Sub

Public Function ChemistryNames() As Variant
    ' Returns array of chemistry metric names (7 metrics, excludes Volume)
    EnsureChemistryNames
    ChemistryNames = mChemistryNames
End Function

Public Function ChemistryCount() As Long
    ' Returns count of chemistry metrics (7, excludes Volume)
    EnsureChemistryNames
    ChemistryCount = UBound(mChemistryNames) - LBound(mChemistryNames) + 1
End Function

' ==== Helper Functions =======================================================

Public Function ColIdx(ByVal tbl As ListObject, ByVal colName As String) As Long
    ' Returns column index (1-based) or 0 if not found
    Dim col As ListColumn
    On Error Resume Next
    Set col = tbl.ListColumns(colName)
    If Not col Is Nothing Then ColIdx = col.Index
    On Error GoTo 0
End Function

Public Function LiveTableName(ByVal site As String) As String
    ' Returns table name for site's live log table (e.g., "tblLive_RP1")
    LiveTableName = LIVE_TABLE_PREFIX & site
End Function

Public Function EnhHidColName(ByVal idx As Long) As String
    ' Returns hidden layer column name (e.g., "EnhHid1", "EnhHid2", ...)
    EnhHidColName = "EnhHid" & idx
End Function

Public Function HistoryTableName(ByVal site As String) As String
    ' Returns table name for site's history table (e.g., "tblHistory_RP1")
    HistoryTableName = HISTORY_TABLE_PREFIX & site
End Function

Public Function TelemECColName(ByVal site As String) As String
    ' Returns telemetry EC column name for site, e.g., "EC (RP1)"
    TelemECColName = "EC (" & site & ")"
End Function

Public Function TelemVolColName(ByVal site As String) As String
    ' Returns telemetry Volume column name for site, e.g., "Vol (RP1)"
    TelemVolColName = "Vol (" & site & ")"
End Function

Public Function SeasonLogTableName(ByVal site As String) As String
    ' Returns table name for site's season backtest table (e.g., "tblSeasonLog_RP1")
    SeasonLogTableName = "tblSeasonLog_" & site
End Function

' ==== Shared Worksheet/Table Helpers ========================================

Public Function GetSheet(ByVal nm As String) As Worksheet
    ' Returns worksheet by name, or Nothing if not found
    On Error Resume Next
    Set GetSheet = ThisWorkbook.Worksheets(nm)
    On Error GoTo 0
End Function

Public Function GetTable(ByVal sheetName As String, ByVal tableName As String) As ListObject
    ' Returns ListObject by sheet and table name, or Nothing if not found
    Dim ws As Worksheet
    Set ws = GetSheet(sheetName)
    If Not ws Is Nothing Then
        On Error Resume Next
        Set GetTable = ws.ListObjects(tableName)
        On Error GoTo 0
    End If
End Function

Public Function MatchesSite(ByVal v As Variant, ByVal site As String) As Boolean
    ' Case-insensitive site comparison
    MatchesSite = (UCase$(Trim$(CStr(v))) = UCase$(Trim$(site)))
End Function

Public Sub StyleActionCell(ByVal cell As Range)
    ' Applies blue hyperlink style to action cells
    With cell
        .Font.Color = COLOR_ACTION_FONT
        .Font.Underline = xlUnderlineStyleSingle
    End With
End Sub

Public Sub InitIRRowAction(ByVal rowRng As Range, ByVal tbl As ListObject)
    ' Sets action cell value and styling only - no other formatting
    Dim actionCol As Long
    actionCol = ColIdx(tbl, IR_COL_ACTION)
    If actionCol > 0 Then
        rowRng.Cells(1, actionCol).Value = ACTION_REMOVE
        StyleActionCell rowRng.Cells(1, actionCol)
    End If
End Sub
