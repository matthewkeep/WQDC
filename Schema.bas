Option Explicit
' Schema: Constants and configuration.
' Dependencies: None

Public Const SHEET_INPUT As String = "Inputs"
Public Const SHEET_LOG As String = "Log"
Public Const SHEET_CHART As String = "Chart"
Public Const SHEET_RAIN As String = "Rain"
Public Const SHEET_RESULTS As String = "Results"
Public Const SHEET_CONFIG As String = "Config"
Public Const SHEET_HISTORY As String = "RunHistory"

Public Const NAME_SITE As String = "RR_Site"
Public Const NAME_OUTPUT As String = "RR_Output"
Public Const NAME_INIT_VOL As String = "RR_InitVol"
Public Const NAME_TRIGGER_VOL As String = "RR_TriggerVol"
Public Const NAME_TRIGGER_RESULT_VOL As String = "RR_TriggerResultVol"
Public Const NAME_SAMPLE_DATE As String = "RR_SampleDate"
Public Const NAME_RUN_DATE As String = "Run_Date"
Public Const NAME_TAU As String = "Cfg_Tau"
Public Const NAME_RAIN_FACTOR As String = "Cfg_RainFactor"
Public Const NAME_SURFACE_FRACTION As String = "Cfg_SurfaceFrac"
Public Const NAME_RAIN_MODE As String = "Cfg_RainMode"
Public Const NAME_NET_OUT As String = "Cfg_NetOut"
Public Const NAME_TELEM_VOL As String = "Telem_Vol"
Public Const NAME_TELEM_EC As String = "Telem_EC"
Public Const NAME_LIMIT_ROW As String = "Limit_Row"
Public Const NAME_RES_ROW As String = "Res_Row"
Public Const NAME_ENHANCED_MODE As String = "Cfg_EnhancedMode"
Public Const NAME_STD_TRIGGER As String = "Std_Trigger"
Public Const NAME_ENH_TRIGGER As String = "Enh_Trigger"
Public Const NAME_HIDDEN_MASS As String = "RR_HiddenMass"
 
Public Const NAME_TRIGGER_PRESET As String = "Cfg_TriggerPreset"
Public Const NAME_CATALOG_RR_LIST As String = "CatalogRRList"
Public Const NAME_TRIGGER_PRESET_LIST As String = "TriggerPresetList"
Public Const IR_COL_SOURCE As String = "Source"
Public Const IR_COL_FLOW As String = "Flow (ML/d)"
Public Const IR_COL_ACTIVE As String = "Active"
Public Const IR_COL_SAMPLE_DATE As String = "Sample Date"
Public Const IR_COL_ACTION As String = "Add Source"
Public Const TABLE_IR As String = "tblIR"
Public Const TABLE_RAIN As String = "tblRainTotals"
Public Const TABLE_RAIN_WQ As String = "tblRainWQ"
Public Const TABLE_RESULTS As String = "tblResults"
Public Const TABLE_CATALOG As String = "tblCatalog"
Public Const TABLE_TRIGGER As String = "tblTrigger"
Public Const TABLE_LOG_DAILY As String = "tblLogDaily"
Public Const TABLE_HISTORY As String = "tblHistory"

Public Const HISTORY_STATUS_ACTIVE As String = "Active"
Public Const HISTORY_STATUS_ROLLEDBACK As String = "RolledBack"

' ==== Color Constants (HEX) ================================
' Sheet/Table colors
Public Const COLOR_STD_HEADER As Long = &HF7EBD3         ' #DDE7EB - Std header fill
Public Const COLOR_ENH_HEADER As Long = &HDAEFDA         ' #E2EFDA - Enhanced header fill
Public Const COLOR_RUN_HEADER As Long = &HCCF2FF         ' #FFF2CC - Run header fill
Public Const COLOR_METADATA_FILL As Long = &HCFCFCF      ' #CFCFCF - Metadata panel
Public Const COLOR_TRIGGER_FONT As Long = &H0000C0       ' #C00000 - Trigger breach font
Public Const COLOR_ROW_HINDCAST As Long = &HE6E6E6       ' #E6E6E6 - Hindcast row
Public Const COLOR_ROW_SAMPLE As Long = &HFFFFDB         ' #DBFFFF - Sample day row
Public Const COLOR_ROW_FORECAST As Long = &HFFFFFF       ' #FFFFFF - Forecast rows

' Chart colors
Public Const COLOR_STD_LINE As Long = &HB3712D           ' #2D71B3 - Standard line
Public Const COLOR_ENH_LINE As Long = &H779900           ' #009977 - Enhanced line
Public Const COLOR_STD_VOLUME As Long = &H404040         ' #404040 - Standard volume
Public Const COLOR_ENH_VOLUME As Long = &H404040         ' #404040 - Enhanced volume
Public Const COLOR_TRIGGER_1 As Long = &H0000C0          ' #C00000 - Trigger line 1
Public Const COLOR_TRIGGER_2 As Long = &H000000          ' #000000 - Trigger line 2
Public Const COLOR_GRIDLINE As Long = &HD6D6D6           ' #D6D6D6 - Chart gridlines

' Input sheet colors
Public Const COLOR_INPUT_RUN_FILL As Long = &HE3FAFF     ' #FFFAE3 - Light yellow
Public Const COLOR_INPUT_CALIB_FILL As Long = &HEBF1EB   ' #EBF1EB - Light green
Public Const COLOR_INPUT_HIDDEN_FILL As Long = &HE9EDF9  ' #F9EDE9 - Light orange
Public Const COLOR_INPUT_RES_FILL As Long = &HD3D3D3     ' #D3D3D3 - Light grey
Public Const COLOR_INPUT_VALUE_FILL As Long = &HFFFFFF   ' #FFFFFF - White
Public Const COLOR_LATEST_FILL As Long = &HCCE4FF        ' #CCE4FF  - Latest WQ fill
Public Const COLOR_TRIGGER_FILL As Long = &HFFB1AC      ' #FDDFDF  - Trigger fill
Public Const COLOR_PREDICTED_FILL As Long = &HE3FAFF      ' #E3FAFF - Predicted WQ fill

' Button colors
Public Const COLOR_BUTTON_OFF As Long = &HBFBFBF         ' #BFBFBF - Button inactive
Public Const COLOR_BUTTON_ON As Long = &H47AD70          ' #70AD47 - Button active

' Font colors
Public Const COLOR_FONT_WHITE As Long = &HFFFFFF         ' #FFFFFF - White text
Public Const COLOR_FONT_BLACK As Long = &H000000         ' #000000 - Black text

' ==== Table Styles ============================
Public Const TABLE_STYLE_DEFAULT As String = "TableStyleMedium2"
Public Const TABLE_STYLE_CONFIG As String = "TableStyleMedium4"
Public Const TABLE_STYLE_LOG As String = "TableStyleLight1"
Public Const TABLE_GAP_COLS As Long = 2 ' Empty columns between horizontal tables

' ==== Input Sheet Block Positions ==============
' Block positions match desired layout per 2025-11-24 image
Public Const INPUT_BLOCK_RES_ROW As Long = 2
Public Const INPUT_BLOCK_RES_COL As Long = 1        ' A2: Reservoir Summary
Public Const INPUT_BLOCK_RUN_ROW As Long = 2
Public Const INPUT_BLOCK_RUN_COL As Long = 10       ' J2: Run Info
Public Const INPUT_BLOCK_RESULTS_ROW As Long = 2
Public Const INPUT_BLOCK_RESULTS_COL As Long = 14   ' N2: Last Run Results
Public Const INPUT_BLOCK_TELEM_ROW As Long = 2
Public Const INPUT_BLOCK_TELEM_COL As Long = 17     ' Q2: Telemetry Snapshot (hideable)
Public Const INPUT_BLOCK_CALIB_ROW As Long = 6
Public Const INPUT_BLOCK_CALIB_COL As Long = 14     ' N6: Calibration Inputs
Public Const INPUT_BLOCK_HIDDEN_ROW As Long = 13
Public Const INPUT_BLOCK_HIDDEN_COL As Long = 17    ' Q13: Hidden Loads (hideable)
Public Const INPUT_BLOCK_INSTRUCT_ROW As Long = 1
Public Const INPUT_BLOCK_INSTRUCT_COL As Long = 20  ' T1: Instructions
Public Const INPUT_BLOCK_IR_TABLE_ROW As Long = 8
Public Const INPUT_BLOCK_IR_TABLE_COL As Long = 1   ' A8: Input Summary table

' ==== UI Titles ============================
Public Const TITLE_RUN_SIM As String = "Run Simulation"
Public Const TITLE_MULTI_RUN As String = "Multi-Run"

Public Const MAX_IR As Long = 10  ' Maximum number of IR (inflow) sources
Public Const DEFAULT_FORECAST_DAYS As Long = 100  ' Default simulation forecast horizon (days)

' ==== UI Defaults ============================
Public Const DEFAULT_SURFACE_FRACTION As Double = 0.8

' ==== Log Table Layout ========================
Public Const LOG_TABLE_START_ROW As Long = 7  ' Row where tblLogDaily header starts (after metadata)
Public Const LOG_METADATA_ROWS As Long = 5  ' Number of rows in log metadata block (A1:B5)

' ==== Simulation Constants ====================
Public Const SIM_EPS_VOL As Double = 0.000001  ' Epsilon for volume comparisons
Public Const TRIGGER_NONE As Long = -1  ' Sentinel value: no trigger occurred during forecast

' ==== Rain Mode Constants =====================
Public Const RAIN_MODE_CONSERVATIVE As String = "Conservative"
Public Const RAIN_MODE_TYPICAL As String = "Typical"
Public Const RAIN_MODE_LIST As String = "Conservative,Typical"  ' For validation dropdowns

' ==== Chart Layout Constants ==================
Public Const CHART_LEFT_POS As Double = 20
Public Const CHART_TOP_START As Double = 20
Public Const CHART_WIDTH As Double = 820
Public Const CHART_HEIGHT_VOLUME As Double = 260
Public Const CHART_HEIGHT_METRIC As Double = 260
Public Const CHART_SPACING As Double = 24
Public Const CHART_SITE_GAP As Double = 60 ' Gap between site chart columns (px)


Private mChemistryNames As Variant


' ==== Volume Metric (Separate) ============================================
Public Const VOLUME_METRIC_NAME As String = "Volume (ML)"


' ==== Chemistry Metrics (Array) ===========================================
Private Sub EnsureChemistryNames()
    If IsEmpty(mChemistryNames) Then
        ' 7 chemistry metrics (NO Volume)
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


'' All legacy MetricNames/MetricCount code removed. Use ChemistryNames/ChemistryCount and VOLUME_METRIC_NAME only.
