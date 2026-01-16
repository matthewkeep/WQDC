# WQOC Architecture Reference

Detailed technical documentation from codebase verification (Jan 2026).

## Module Summary

| Module | Lines | Purpose |
|--------|-------|---------|
| Core.bas | ~60 | Types (State, Config, Result), constants |
| Schema.bas | ~200 | Sheet/table/column names, colors |
| Helpers.bas | ~180 | Utilities (ColIdx, GetSheet, GetTable, FindRowByDate) |
| Modes.bas | ~120 | Mixing models (StepSimple, StepTwoBucket) |
| Sim.bas | ~100 | Simulation loop, trigger detection |
| Data.bas | ~360 | Inputs sheet I/O, state loading |
| Telemetry.bas | ~100 | Rain/EC/Vol lookups (O(1) via MATCH) |
| SimLog.bas | ~340 | Date-centric live log (UPSERT) |
| History.bas | ~280 | Audit trail, rollback |
| Loader.bas | ~250 | Site selection, IR population |
| Events.bas | ~290 | Worksheet handlers |
| WQOC.bas | ~240 | Entry point, orchestration |
| Setup.bas | ~950 | Table creation, scaffolding |
| Error.bas | ~20 | Trace, TraceErr |

**Total:** ~16 modules, ~4,500 lines

---

## Pipeline Traces

### P1: Run Pipeline

```
WQOC.Run (WQOC.bas:5)
  │
  ├─► Data.LoadState (Data.bas:31)
  │     Returns: State UDT
  │     - Vol from NAME_INIT_VOL
  │     - Chem(1-7) from NAME_RES_ROW
  │     - Hidden(1-7) from NAME_HIDDEN_MASS
  │
  ├─► Data.LoadConfig (Data.bas:98)
  │     Returns: Config UDT
  │     - Site, Days, StartDate, Tau, Outflow, SurfaceFrac
  │     - Inflow, InflowChem(1-7) from IR table
  │     - TriggerVol, TriggerChem(1-7) from trigger row
  │     - Mode, RainfallMode, RainFactor (Enhanced only)
  │
  ├─► Sim.Run (Sim.bas:5)
  │     Input: State, Config
  │     Returns: Result UDT
  │     - Snaps(0 to Days) - daily state snapshots
  │     - TriggerDay, TriggerDate, TriggerMetric
  │     - FinalState
  │     │
  │     └─► Modes.StepSimple/StepTwoBucket (Modes.bas)
  │           Called for each day d = 1 to Days
  │
  ├─► SimLog.WriteLog (SimLog.bas:11)
  │     Input: Result, Config, RunId, Site
  │     - Detects STD/ENH from Left$(runId, 3)
  │     - UPSERT rows by date (EnsureRowForDate)
  │     - Writes Vol + Chem(1-7) to Std or Enh columns
  │     - Writes Hidden(1-7) for Enhanced
  │     - Calculates ErrVol/ErrEC from telemetry
  │
  ├─► History.RecordRun (History.bas:8)
  │     Input: cfgStd, rStd, cfgEnh, rEnh, hasEnhanced, runId, site
  │     - Single row per run (both Std and Enh results)
  │     - Enh columns blank when hasEnhanced=False
  │     - Updates older rows to "Rollback" action
  │
  └─► Data.SaveResult (Data.bas:198)
        Input: Result, runType ("Standard"/"Enhanced")
        - Writes to NAME_STD_TRIGGER or NAME_ENH_TRIGGER
        - Updates Pred_Row if Pred_Mode matches runType
        - Formats triggered cell red + bold
```

### P2: Rollback Pipeline

```
Events.OnHistoryDoubleClick (Events.bas:137)
  │
  ├─► Guard: Is in Action column? (line 181)
  │     - If rowIdx == ListRows.Count → "This is current run" → Exit
  │
  ├─► History.RollbackTo (History.bas:198)
  │     Input: runId, site
  │     - Finds target row by runId
  │     - Gets targetStartDate from HISTORY_COL_RUNDATE
  │     │
  │     └─► SimLog.DeleteAfterDate (SimLog.bas:294)
  │           Input: cutoffDate, site
  │           - Deletes rows where Date > cutoffDate
  │           - Backward loop to avoid index issues
  │
  │     - Deletes history rows after target (backward loop)
  │     - Updates new last row to "Current"
  │
  ├─► RefreshHistoryActions (Events.bas:241)
  │     - Updates action column text after deletion
  │
  ├─► History.LoadSettings (History.bas:108)
  │     Input: runId, site
  │     - Restores 7 config fields to Inputs sheet
  │     - Sample Date write triggers OnInputsChange
  │
  └─► WQOC.Run (auto re-run)
```

### P3: Sample Date Pipeline

```
Events.OnInputsChange (Events.bas:23)
  │
  ├─► Intersect with NAME_SAMPLE_DATE? (line 43)
  │
  └─► Data.LoadHiddenForDate (Data.bas:339)
        Input: site, sampleDate
        │
        └─► Data.LoadHiddenFromLog (Data.bas:305)
              - Gets tblLive_{site}
              - Finds row by date (Helpers.FindRowByDate)
              - Reads EnhHidColName(1-7) columns
              - Returns State with Hidden(1-7) populated
        │
        └─► Writes to NAME_HIDDEN_MASS cells (O16:O22)
```

### P4: Telemetry Pipeline

```
Sim.GetRainForDay (Sim.bas:85)
  │
  └─► Telemetry.GetRain (Telemetry.bas:21)
        Input: targetDate
        - Uses Application.Match for O(1) lookup
        - Returns rain value in mm

SimLog.WriteDiscrepancy (SimLog.bas:138)
  │
  ├─► Telemetry.GetLatestEC (Telemetry.bas:44)
  │     Input: beforeDate, site
  │     - Application.Match + backward scan
  │     - Returns most recent EC <= beforeDate
  │
  └─► Telemetry.GetLatestVol (Telemetry.bas:71)
        Input: beforeDate, site
        - Same pattern as GetLatestEC
```

---

## Wiring Verification

### Call Site → Target Verification

| Call Site | Line | Target | Status |
|-----------|------|--------|--------|
| WQOC.Run | 38 | Data.LoadState | ✓ State UDT returned |
| WQOC.Run | 39 | Data.LoadConfig | ✓ Config UDT returned |
| WQOC.Run | 54 | Sim.Run | ✓ Result UDT returned |
| WQOC.Run | 55 | SimLog.WriteLog | ✓ Correct 4 params |
| WQOC.Run | 92 | History.RecordRun | ✓ Correct 7 params |
| WQOC.Run | 56 | Data.SaveResult | ✓ Correct 2 params |
| Events | 49 | Data.LoadHiddenForDate | ✓ Direct call |
| Events | 192 | History.RollbackTo | ✓ Correct params |

### Parameter Signatures

```vba
' SimLog.WriteLog
Public Sub WriteLog(ByRef r As Result, ByRef cfg As Config, _
                    ByVal runId As String, ByVal site As String)

' History.RecordRun
Public Sub RecordRun(ByRef cfgStd As Config, ByRef rStd As Result, _
                     ByRef cfgEnh As Config, ByRef rEnh As Result, _
                     ByVal hasEnhanced As Boolean, ByVal runId As String, _
                     ByVal site As String)

' Data.SaveResult
Public Sub SaveResult(ByRef r As Result, ByVal runType As String)
```

---

## Table Column Mappings

### tblLive_{site} (28 columns)

| Index | Column | Written By | Read By |
|-------|--------|------------|---------|
| 1 | Date | SimLog | SimLog, Data |
| 2 | Days | SimLog | - |
| 3 | StdVol | SimLog.WriteLiveStandard | WQOC.Charts |
| 4-10 | StdEC..StdTAN | SimLog.WriteLiveStandard | WQOC.Charts |
| 11 | EnhVol | SimLog.WriteLiveEnhanced | WQOC.Charts |
| 12-18 | EnhEC..EnhTAN | SimLog.WriteLiveEnhanced | WQOC.Charts |
| 19-25 | EnhHidEC..EnhHidTAN | SimLog.WriteLiveEnhanced | Data.LoadHiddenFromLog |
| 26 | ErrVol | SimLog.WriteDiscrepancy | - |
| 27 | ErrEC | SimLog.WriteDiscrepancy | - |
| 28 | RunId | SimLog | - |

### tblHistory_{site} (17 columns)

| Column | Written By | Read By |
|--------|------------|---------|
| RunId | History.RecordRun | History.LoadSettings |
| Timestamp | History.RecordRun | - |
| RunDate | History.RecordRun | History.RollbackTo |
| Days | History.RecordRun | - |
| RainfallMode | History.RecordRun | History.LoadSettings |
| TelemCal | History.RecordRun | History.LoadSettings |
| Tau | History.RecordRun | History.LoadSettings |
| SurfaceFrac | History.RecordRun | History.LoadSettings |
| RainFactor | History.RecordRun | History.LoadSettings |
| StdMode | History.RecordRun | - |
| StdTriggerDay | History.RecordRun | - |
| StdTriggerMetric | History.RecordRun | - |
| EnhMode | History.RecordRun | History.LoadSettings |
| EnhTriggerDay | History.RecordRun | - |
| EnhTriggerMetric | History.RecordRun | - |
| Action | History.RecordRun | Events.OnHistoryDoubleClick |
| Load | History.RecordRun | Events.OnHistoryDoubleClick |

---

## Integration Points

### RunId Format

```
Base:    {site}-{yyyymmdd}-{seq}     e.g., RP1-20260115-001
Full:    STD-{base} or ENH-{base}    e.g., STD-RP1-20260115-001

Generated: WQOC.MakeRunId (line 126)
Prefixed:  WQOC.Run (lines 55, 87)
Parsed:    SimLog.WriteLog (line 16) - Left$(runId, 3)
```

### Date Handling

```vba
' Type: Always Date
StartDate As Date                      ' Core.bas:30

' Parsing
GetDateVal = CDate(v)                  ' Helpers.bas:150

' MATCH (requires Double)
Application.Match(CDbl(targetDate), ...) ' Helpers.bas:94

' Arithmetic
logDate = cfg.StartDate + i            ' SimLog.bas:47
```

### Column Name Generation

```vba
' Schema.bas functions
Schema.StdChemColName(j)   ' Returns "Std" & ChemShortName(j)
Schema.EnhChemColName(j)   ' Returns "Enh" & ChemShortName(j)
Schema.EnhHidColName(j)    ' Returns "EnhHid" & ChemShortName(j)

' ChemShortName returns: EC, F_U, F_Mn, SO4, Mg, Ca, TAN
```

---

## Chart System

### Behavior
- Charts created once per site, never deleted
- Multiple sites stack horizontally (each site gets its own column)
- Data bound directly to table column ranges (auto-update)
- Named `cht_{site}_{metric}` (e.g., `cht_RP1_EC`)

### Layout
```
┌─────────────┐  ┌─────────────┐  ┌─────────────┐
│   Site 1    │  │   Site 2    │  │   Site 3    │
│  EC Chart   │  │  EC Chart   │  │  EC Chart   │
├─────────────┤  ├─────────────┤  ├─────────────┤
│  F_U Chart  │  │  F_U Chart  │  │  F_U Chart  │
├─────────────┤  ├─────────────┤  ├─────────────┤
│    ...      │  │    ...      │  │    ...      │
└─────────────┘  └─────────────┘  └─────────────┘
     X=20          X=20+W+24        X=20+2*(W+24)
```

### Chart Functions (WQOC.bas)

| Function | Purpose |
|----------|---------|
| `GenerateCharts` | Main entry - creates or updates charts for site |
| `GetOrCreateChart` | Gets existing chart or creates new one |
| `ChartNeedsSeries` | Checks if chart needs series (empty) |
| `GetSiteChartLeft` | Calculates X position (existing or new stack) |
| `BuildChartSeries` | Creates all series for a chart from table columns |
| `AddDataSeries` | Adds a data series with styling |
| `AddTriggerLine` | Adds constant trigger threshold line |
| `FormatChart` | Applies title, axes, legend formatting |
| `UpdateChartRanges` | Updates existing chart data sources |
| `GetColRange` | Returns `ListColumn.DataBodyRange` for binding |

### Series per Chart
**EC Chart (dual-axis):** up to 6 series
1. Std EC - Left Y-axis, solid blue
2. Enh EC - Left Y-axis, solid teal (if Enhanced)
3. EC Trigger - Left Y-axis, dash-dot red (if threshold set)
4. Std Vol - Right Y-axis, dashed blue
5. Enh Vol - Right Y-axis, dashed teal (if Enhanced)
6. Vol Trigger - Right Y-axis, dash-dot red (if threshold set)

**Other Charts (single-axis):** up to 3 series
1. Std {Metric} - Left Y-axis, solid blue
2. Enh {Metric} - Left Y-axis, solid teal (if Enhanced)
3. {Metric} Trigger - Left Y-axis, dash-dot red (if threshold set)

### Key Design Decisions
- **No array allocation**: Bind to table columns directly to avoid overflow
- **No delete/recreate**: Causes Excel state inconsistency on consecutive runs
- **Persistent charts**: Site charts remain when switching sites
- **Horizontal stacking**: New sites add columns to the right

---

## Refactoring History (Jan 2026)

### Deleted Modules
- `Storage.bas` - 66 lines, never integrated
- `EventBus.bas` - 60 lines, over-engineered 5-level indirection

### Removed Functions
- `SimLog.FindRowByDate` - wrapper for Helpers
- `SimLog.FindTelemRowByDate` - wrapper for Helpers
- `Data.FindLogRowByDate` - wrapper for Helpers
- `SimLog.ClearSiteLog` - unused
- `SimLog.GetLatestLogDate` - unused
- `History.GetLastRun` - unused
- `History.GetRunHistory` - unused
- `Error.Raise` - unused
- `Setup.StyleActionColumn` - unused

### Consolidated Functions
- `Setup.MakeTbl` + `MakeTblLight` → single `MakeTbl` with optional autofit
- `Loader.GetLatestLabData` + `GetLatestLabDataFiltered` → single function with optional cutoff

### Optimizations
- `Telemetry.GetLatestEC/Vol` - O(n) loops → O(1) MATCH + backward scan
- `GenerateCharts` - Array allocation → table column binding (fixes overflow)

### Chart System Refactor
- **Problem**: Delete/recreate every run caused overflow on consecutive runs
- **Solution**: Create once, bind to table columns, stack sites horizontally
- **Functions**: `GetOrCreateChart`, `BuildChartSeries`, `AddDataSeries`, `AddTriggerLine`, `FormatChart`, `UpdateChartRanges`, `GetColRange`
- **EC-only volume**: Only EC chart has dual-axis with volume; other charts single-analyte
- **Styling constants**: `CHART_LINE_WEIGHT`, `CHART_TRIGGER_WEIGHT` in Schema.bas

### Metrics

| Metric | Before | After |
|--------|--------|-------|
| Modules | 18 | 16 |
| Lines | ~4,800 | ~4,500 |
| Wrapper functions | 7 | 0 |
| Dead functions | 5 | 0 |
