Option Explicit
' Schema: Constants only (sheet/table/column names, colors, defaults).
' Dependencies: Core (for MetricName)
' Note: Utility functions moved to Helpers.bas

' ==== Sheet Names ============================================================
Public Const SHEET_INPUT As String = "Inputs"
Public Const SHEET_LOG As String = "Log"
Public Const SHEET_CHART As String = "Chart"
Public Const SHEET_RESULTS As String = "Results"
Public Const SHEET_CONFIG As String = "Config"
Public Const SHEET_RECORD As String = "Record"

' ==== Named Ranges ===========================================================
' Reservoir state (Row 3)
Public Const NAME_SITE As String = "RR_Site"
Public Const NAME_INIT_VOL As String = "RR_InitVol"
Public Const NAME_RES_ROW As String = "Res_Row"
Public Const NAME_SAMPLE_DATE As String = "RR_SampleDate"
Public Const NAME_OUTPUT As String = "RR_Output"
Public Const NAME_RUN_DATE As String = "Run_Date"

' Trigger limits (Row 4)
Public Const NAME_TRIGGER_VOL As String = "RR_TriggerVol"
Public Const NAME_LIMIT_ROW As String = "Limit_Row"
Public Const NAME_TRIGGER_PRESET As String = "Trigger_Preset"

' Predicted results (Row 5)
Public Const NAME_RESULT_VOL As String = "Result_Vol"
Public Const NAME_PRED_ROW As String = "Pred_Row"
Public Const NAME_PRED_MODE As String = "Pred_Mode"

' Trigger days (Row 3-4, Col O-P)
Public Const NAME_STD_TRIGGER As String = "Std_Trigger"
Public Const NAME_ENH_TRIGGER As String = "Enh_Trigger"

' Sign Off (N7:O10)
Public Const NAME_SIGN_OFF_NAME As String = "SignOff_Name"

' Enhanced settings (R1:S16)
Public Const NAME_ENHANCED_MODE As String = "Cfg_EnhancedMode"
Public Const NAME_TELEM_CAL As String = "Cfg_TelemCal"
Public Const NAME_RAINFALL_MODE As String = "Cfg_RainfallMode"
Public Const NAME_RAIN_FACTOR As String = "Cfg_RainFactor"
Public Const NAME_MIXING_MODEL As String = "Cfg_MixingModel"
Public Const NAME_TAU As String = "Cfg_Tau"
Public Const NAME_SURFACE_FRACTION As String = "Cfg_SurfaceFrac"
Public Const NAME_HIDDEN_MASS As String = "RR_HiddenMass"

' Action buttons
Public Const NAME_RUN_CELL As String = "Run_Simulation"
Public Const NAME_LOAD_CELL As String = "Load_Latest"

' ==== Table Names ============================================================
Public Const TABLE_IR As String = "tblIR"
Public Const TABLE_TELEMETRY As String = "tblTelemetry"
Public Const TABLE_RESULTS As String = "tblResults"
Public Const TABLE_INDEX As String = "tblIndex"
Public Const TABLE_TRIGGERS As String = "tblTriggers"
Public Const TABLE_SIGN As String = "tblSign"

' Per-site table prefixes (tables created on-demand)
Public Const LIVE_TABLE_PREFIX As String = "tblLive_"
Public Const HISTORY_TABLE_PREFIX As String = "tblHistory_"

' Live table columns (date-centric log with Std/Enh side-by-side)
Public Const LIVE_COL_DATE As String = "Date"
Public Const LIVE_COL_DAYS As String = "Days"
Public Const LIVE_COL_STD_VOL As String = "StdVol"
Public Const LIVE_COL_STD_EC As String = "StdEC"
Public Const LIVE_COL_ENH_VOL As String = "EnhVol"
Public Const LIVE_COL_ENH_EC As String = "EnhEC"
Public Const LIVE_COL_ERR_VOL As String = "ErrVol"
Public Const LIVE_COL_ERR_EC As String = "ErrEC"
Public Const LIVE_COL_RUNID As String = "RunId"
' Note: Std/Enh chemistry columns built via StdChemColName, EnhChemColName, EnhHidColName

' ==== Column Names ===========================================================
' IR table columns
Public Const IR_COL_SOURCE As String = "Source"
Public Const IR_COL_FLOW As String = "Flow (ML/d)"
Public Const IR_COL_ACTIVE As String = "Active"
Public Const IR_COL_SAMPLE_DATE As String = "Sample Date"
Public Const IR_COL_ACTION As String = "Add Input"

' History table columns
Public Const HISTORY_COL_RUNID As String = "RunId"
Public Const HISTORY_COL_TIMESTAMP As String = "Timestamp"
Public Const HISTORY_COL_RUNDATE As String = "RunDate"
Public Const HISTORY_COL_DAYS As String = "Days"
Public Const HISTORY_COL_ACTION As String = "Action"
Public Const HISTORY_COL_LOAD As String = "Load"

' History snapshot columns (for accurate rollback)
Public Const HISTORY_COL_SAMPLE_DATE As String = "SampleDate"
Public Const HISTORY_COL_TRIGGER_VOL As String = "TriggerVol"
Public Const HISTORY_COL_RES_CHEM As String = "ResChemistry"
Public Const HISTORY_COL_TRIGGER_CHEM As String = "TriggerChemistry"
Public Const HISTORY_COL_HIDDEN_MASS As String = "HiddenMass"
Public Const HISTORY_COL_IR_SNAPSHOT As String = "IRSnapshot"

' Telemetry columns (Date and Rain are fixed; EC/Vol are per-site)
Public Const TELEM_COL_DATE As String = "Date"
Public Const TELEM_COL_RAIN As String = "Rain (mm)"

' Volume metric name
Public Const VOLUME_METRIC_NAME As String = "Volume (ML)"

' ==== Action Values ==========================================================
Public Const ACTION_ADD As String = "Add"
Public Const ACTION_REMOVE As String = "Remove"
Public Const ACTION_ROLLBACK As String = "Rollback"
Public Const ACTION_CURRENT As String = "Current"

' ==== Color Constants ========================================================
' Chart colors (used by WQOC.GenerateCharts)
Public Const COLOR_STD_LINE As Long = &HB3712D       ' #2D71B3 - Standard line
Public Const COLOR_ENH_LINE As Long = &H779900       ' #009977 - Enhanced line
Public Const COLOR_TRIGGER_LINE As Long = &H0000C0   ' #C00000 - Trigger threshold

' Button colors (used by Setup.SetupControls)
Public Const COLOR_BUTTON_ON As Long = &H47AD70      ' #70AD47 - Button active
Public Const COLOR_BUTTON_LOAD As Long = &HDAEFE2   ' #E2EFDA - Load Latest button

' Log row colors
Public Const COLOR_SAMPLE_DATE As Long = &HFFFFCC    ' #CCFFFF - Light cyan for sample date row
Public Const COLOR_RUN_DATE As Long = &HCCFFCC       ' #CCFFCC - Light green for run date row

' Trigger formatting
Public Const COLOR_TRIGGER_FONT As Long = &H0000C0   ' #C00000 - Red for triggered values

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

Public Const TELEM_CAL_ON As String = "On"

' ==== Chart Layout ===========================================================
Public Const CHART_LEFT_POS As Double = 20
Public Const CHART_TOP_START As Double = 20
Public Const CHART_WIDTH As Double = 820
Public Const CHART_HEIGHT As Double = 260
Public Const CHART_SPACING As Double = 24

' ==== Chart Styling ==========================================================
Public Const CHART_LINE_WEIGHT As Double = 2
Public Const CHART_TRIGGER_WEIGHT As Double = 1.5

' ==== Chemistry Metrics (Private) =============================================

Private mChemistryNames As Variant

Private Sub EnsureChemistryNames()
    If IsEmpty(mChemistryNames) Then
        ' 7 chemistry metrics (excludes Volume) - full names with units
        mChemistryNames = Array("EC (uS/cm)", "F_U (ug/L)", "F_Mn (ug/L)", "SO4 (mg/L)", "Mg (mg/L)", "Ca (mg/L)", "TAN (mg/L)")
    End If
End Sub

Public Function ChemistryNames() As Variant
    ' Returns array of chemistry metric names (7 metrics, excludes Volume)
    EnsureChemistryNames
    ChemistryNames = mChemistryNames
End Function

' ==== Chemistry Metrics (Public) ==============================================

Public Function ChemistryCount() As Long
    ' Returns count of chemistry metrics (7, excludes Volume)
    ChemistryCount = Core.METRIC_COUNT
End Function

Public Function ChemShortName(ByVal idx As Long) As String
    ' Returns short name for chemistry index (1-based): EC, F_U, F_Mn, SO4, Mg, Ca, TAN
    ' Delegates to Core.MetricName (single source of truth)
    ChemShortName = Core.MetricName(idx)
End Function

Public Function StdChemColName(ByVal idx As Long) As String
    ' Standard chemistry column name: StdEC, StdF_U, etc.
    StdChemColName = "Std" & ChemShortName(idx)
End Function

Public Function EnhChemColName(ByVal idx As Long) As String
    ' Enhanced visible chemistry column name: EnhEC, EnhF_U, etc.
    EnhChemColName = "Enh" & ChemShortName(idx)
End Function

Public Function EnhHidColName(ByVal idx As Long) As String
    ' Enhanced hidden mass column name: EnhHidEC, EnhHidF_U, etc.
    EnhHidColName = "EnhHid" & ChemShortName(idx)
End Function

