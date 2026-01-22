# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

WQOC (Water Quality Optimisation Calculator) is an Excel/VBA simulation tool for mining wastewater treatment. Models reservoir inflows, mixing, and trigger-based releases.

**Platform:** Excel VBA (Windows & Mac via DictionaryShim)

## Quick Start

```vba
Setup.BuildAll           ' Create sheets, buttons, seed data
Setup.Initialize         ' Create per-site tables/columns from Index
WQOC.Run                 ' Run simulation (Standard + Enhanced if enabled)
WQOC.Rollback            ' Undo last run for current site
Tests.RunSmokeSuite      ' 10 smoke tests
Scenarios.RunAll         ' 6 regression scenarios
```

## Architecture

```
WQOC.bas ─┬─ Data.bas ──────── Helpers.bas ── Schema.bas
          ├─ Telemetry.bas ─── Helpers.bas ── Schema.bas
          ├─ Sim.bas ───────── Modes.bas ──── Core.bas, Schema.bas
          ├─ History.bas ───── SimLog.bas ─── Helpers.bas
          ├─ Loader.bas ────── Helpers.bas
          ├─ Events.bas ────── Helpers.bas
          └─ (Charts)
```

### Module Layers

```
┌─────────────────────────────────────────────────────────────┐
│ Entry Points: WQOC.bas, Events.bas                          │
├─────────────────────────────────────────────────────────────┤
│ Business Logic: Sim.bas, Modes.bas, History.bas, SimLog.bas │
├─────────────────────────────────────────────────────────────┤
│ Data Access: Data.bas, Loader.bas, Telemetry.bas            │
├─────────────────────────────────────────────────────────────┤
│ Infrastructure: Helpers.bas, Schema.bas, Core.bas           │
└─────────────────────────────────────────────────────────────┘
```

### Modules

| Module | Purpose |
|--------|---------|
| **Core.bas** | Types (State, Config, Result), constants, pure functions |
| **Modes.bas** | Mixing models: StepSimple, StepTwoBucket |
| **Sim.bas** | Simulation loop, trigger detection, rainfall integration |
| **Data.bas** | Inputs sheet I/O, state loading, result saving |
| **Telemetry.bas** | Telemetry data access (Rain, EC, Vol) |
| **History.bas** | Audit trail, rollback, LoadSettings (restore config) |
| **SimLog.bas** | Date-centric live log (UPSERT to tblLive) |
| **Loader.bas** | Site selection, IR/chemistry population |
| **Events.bas** | Worksheet handlers, double-click toggles, date validation |
| **WQOC.bas** | Entry point, orchestration, chart generation |
| **Schema.bas** | Constants only (names, colors, defaults) |
| **Helpers.bas** | Utilities (ColIdx, GetSheet, GetTable, serialization, range access) |
| **Error.bas** | Centralized error handling (Trace, TraceErr, DEBUG_ON toggle) |
| **Setup.bas** | Scaffolding, table creation, dropdowns, conditional formatting |
| **Backtest.bas** | Season replay for A/B comparison |
| **Tests.bas** | Smoke tests |
| **Scenarios.bas** | Regression scenarios |
| **Validate.bas** | Structure validation, date format checks |
| **DictionaryShim.cls** | Mac compatibility |

### Core Types

```vba
Type State    ' Vol, Chem(1-7), Hidden(1-7) - UDT, defaults to zeros
Type Config   ' Site, Mode, Days, Tau, Inflow, Outflow, Triggers, RainfallMode, RainFactor, SurfaceFrac
Type Result   ' TriggerDay, TriggerMetric, TriggerDate, Snaps(), FinalState

Enum Metric   ' Use instead of magic numbers: s.Chem(mEC), cfg.TriggerChem(mSO4)
    mEC = 1, mF_U = 2, mF_Mn = 3, mSO4 = 4, mMg = 5, mCa = 6, mTAN = 7
End Enum
```

### Key Flows

**Run Simulation:**
```
WQOC.Run → Data.LoadState/LoadConfig → Sim.Run → Modes.Step
         → SimLog.WriteLog → History.RecordRun → Data.SaveResult → GenerateCharts
```

**Rollback (from History table):**
```
Events.OnHistoryDoubleClick → History.RollbackTo → SimLog.DeleteAfterDate
                            → History.LoadSettings → WQOC.Run (auto re-run)
```

**Load Settings (from History table):**
```
Events.OnHistoryDoubleClick → History.LoadSettings → Writes to Inputs sheet
                            → Sample Date change triggers Events.OnInputsChange
                            → Data.LoadHiddenForDate (auto-loads hidden mass)
```

## Conventions

**Type Variables:**
- `s` = State (input)
- `n` = State (output/next in step functions)
- `cfg` = Config
- `r` = Result
- `ws` = Worksheet
- `tbl` = ListObject
- `row` = ListRow
- `rng` = Range

**Loop/Index Variables:** `i`, `j`, `d` (day), `col` (column index)

**Prefixes:** `p` = previous (e.g., `pVol`), `mix` = after mixing (e.g., `mixVol`)

**Headers:** `Option Explicit` + `' Module: desc` + `' Dependencies: X, Y`

**Helpers:** Use `Helpers.ColIdx`, `Helpers.GetSheet`, `Helpers.GetTable` (not Schema)

**Error Handling:**
```vba
Public Sub DoWork()
    On Error GoTo Fail
    ' ... code ...
    Exit Sub
Fail:
    Error.TraceErr "Module.DoWork"
End Sub
```
- Use `Error.Trace "src", "msg"` for diagnostic logging
- Use `Error.TraceErr "src"` in Fail handlers
- Set `Error.DEBUG_ON = False` for production (silences all logging)

## Extending

- **Add mode:** New `StepX` in Modes.bas
- **Add metric:** Update `METRIC_COUNT` in Core.bas, update `ChemistryNames()` in Schema.bas
- **Add trigger:** Update `ChkTriggers` in Sim.bas
- **Add helper:** Put in Helpers.bas (not Schema.bas)

## Working Style

- **Smallest effective action** - do less, not more
- **Fix, don't improve** - solve the problem, stop there
- **Silence is approval** - don't ask, just do (within scope)
- Bullets over paragraphs, code over explanation

## Gotchas

See `.claude/agents/_gotchas.md` for full list. Key ones:

| Issue | Fix |
|-------|-----|
| `Log` is reserved | Use `SimLog`, `AuditLog` |
| Mac compatibility | Use `DictionaryShim` not `Scripting.Dictionary` |
| Table access | Check `tbl.DataBodyRange Is Nothing` before access |
| History/SimLog | Share RunId for rollback coordination |
| Helper functions | Use `Helpers.*` not `Schema.*` |
| Conditional format order | Apply Enhanced last (highest priority) |
| Mixing mode strings | Use `Schema.MIXING_SIMPLE`, `Schema.MIXING_TWOBUCKET` |
| Dropdown validation | Uses `INDIRECT("tblName[colName]")` - column names must match exactly |
| Table column names | Use Schema constants for both table creation AND column lookups |

## Per-Site Architecture

- **Live tables per site:** `tblLive_RP1`, `tblHistory_RP1`
- **Live table structure:** Date-centric with Std/Enh side-by-side (28 columns)
  - Date, Days, StdVol, Std[7 chem], EnhVol, Enh[7 chem], EnhHid[7 chem], ErrVol, ErrEC, RunId
  - Days: Relative to run date (0 = today, negative = past, positive = forecast)
  - Column names: `StdEC`, `StdF_U`, `EnhEC`, `EnhF_U`, `EnhHidEC`, `EnhHidF_U`, etc.
  - One row per date (UPSERT on run, not append)
  - All 7 chemistry metrics logged for Standard and Enhanced
  - Hidden layer stored for TwoBucket continuity between runs
  - Discrepancy columns (ErrVol/ErrEC) compare prediction vs telemetry
  - Row shading: Sample date = light cyan, Run date = light green
  - Triggered values: Red + bold formatting on triggered metric cell
- **History table structure:**
  - Columns (15): RunId, Timestamp, RunDate, SampleDate, Outflow, ResChemistry, IRSnapshot, Triggers, StdResult, EnhResult, EnhSettings, HiddenMass, SignName, Action, Load
  - Bundled columns: Triggers=`Vol|EC|...|TAN|Preset`, StdResult/EnhResult=`Days|Metric`, EnhSettings=`Enabled|TelemCal|RainfallMode|RainFactor|MixingModel|Tau|SurfaceFrac`
  - One row per run (captures both Std and Enh results; Enh columns blank when disabled)
  - Action column: "Current" (latest) or "Rollback" (older runs)
  - Load column: Click to restore settings without rollback
- **RRState table structure (aligned with History):**
  - Columns (12): Site, RunDate, SampleDate, Outflow, ResChemistry, IRSnapshot, Triggers, EnhSettings, HiddenMass, SignName, PredView, LastModified
  - Bundled columns: ResChemistry=`Vol|EC|...|TAN`, Triggers=`Vol|EC|...|TAN|Preset`, PredView=`Vol|EC|...|TAN|Mode`
  - One row per site (upsert on site switch)
- **Telemetry table:** Located on Results sheet at column L (tblTelemetry)
- **Telemetry columns per site:** `EC (RP1)`, `Vol (RP1)` (Rain is global)
- **RunId format:** `{site}_{seq}` (e.g., `RP1_001`)
- **Rollback:** Deletes future data, loads settings, auto-runs simulation
- **Load Settings:** Restores config to Inputs (no deletion, no run)
- **Charts:** 7 charts per site (1 dual-axis, 6 single-axis), stacked vertically
  - Named `cht_{site}_{metric}` (e.g., `cht_RP1_EC`)
  - Created once, never deleted - data bound to table columns
  - EC chart (first): Dual-axis with Volume on right Y-axis
  - Other charts: Single-analyte only (no volume)
  - Styling: Std=Blue, Enh=Teal, Trigger=Red dotted
  - Multiple sites stack horizontally (new site → new column to right)
- Tables created on-demand (first run) or via `Setup.Initialize`

## Enhanced Mode

- **Rainfall:** Telemetry in mm/day × RainFactor = volume added (ML)
- **Hidden mass:** Auto-loads from tblLive when Sample Date changes
- **Conditional UI:**
  - Enhanced Off → greys all Enhanced settings (R3:S16)
  - Rainfall Off → greys Rain Factor (R5:S5)
  - Mixing Model Simple → greys Tau, Surface Fraction, Hidden Mass (R7:S16)
- **Double-click toggles:** Enabled, Telemetry Cal (On/Off), Pred Mode (Standard/Enhanced)

## Inputs Sheet Layout

**J5:** Pred_Mode toggle (Standard/Enhanced) - controls which result displays in Predicted row (Row 5)

**N4:P4:** Enhanced results row (greyed when Enhanced=Off)

**N7:O10 Sign Off Block:**
```
N7:  Sign Off (header)
N8:  Name label       O8: Name dropdown (linked to tblUsers)
N9:  Signed label     O9: Signed value
N10: Position label   O10: Position (VLOOKUP from tblUsers)
```

**Column R-S (Enhanced Settings):**

```
R1:  Enhanced (header)
R2:  Enabled           S2: On/Off (double-click toggle)
R3:  Telemetry Cal     S3: On/Off (double-click toggle)
R4:  Rainfall          S4: Off/Hindcast/Hindcast+Forecast
R5:  Rain Factor       S5: number (greyed when Rainfall=Off)
R6:  Mixing Model      S6: Simple/TwoBucket
R7:  Tau (days)        S7: number (greyed when Model=Simple)
R8:  Surface Fraction  S8: number (greyed when Model=Simple)
R9:  Hidden Mass       (header, greyed when Model=Simple)
R10-R16: Chemistry labels (greyed when Model=Simple)
S10-S16: Hidden mass values
```

**Predicted Row (Row 5):**
- B5: Volume, C5-I5: Chemistry metrics (Pred_Row named range)
- J5: Pred_Mode toggle (Standard/Enhanced) - double-click to switch
- Triggered metric displays red + bold formatting

## History Table Actions

| Column | Click Action |
|--------|--------------|
| Action ("Rollback") | Delete future runs, load settings, re-run simulation |
| Action ("Current") | Shows message (can't rollback current) |
| Load | Restore settings to Inputs sheet (no deletion, no run) |

## Validation & Testing

```vba
Validate.Check           ' Returns True if structure valid (used before WQOC.Run)
Validate.Report          ' Detailed validation report to Immediate window
Tests.RunSmokeSuite      ' 10 smoke tests for core math
Scenarios.RunAll         ' 6 regression scenarios
Backtest.RunSeason       ' Season replay with A/B comparison (Std vs Enh)
```

**Date validation:** Run Date (K3) and Sample Date (L3) are validated on entry (Events.bas) and before run (Validate.bas). Invalid dates are cleared with user message.

## Verified Patterns (Required)

These patterns have been verified across the entire codebase. Follow them exactly.

### Data Access (Mandatory)

| Operation | Required Pattern | Never Do |
|-----------|------------------|----------|
| Get table | `Helpers.GetTable(sheetName, tableName)` | `ws.ListObjects()` |
| Get column | `Helpers.ColIdx(tbl, colName)` | Hardcoded indices |
| Get sheet | `Helpers.GetSheet(sheetName)` | `Worksheets()` |
| Find row | `Helpers.FindRowByDate(tbl, date)` | Local loops |
| Get date | `Helpers.GetDateVal(ws, rangeName)` | Raw `CDate()` |

### Guards (Mandatory)

```vba
' Table access
Set tbl = Helpers.GetTable(...)
If tbl Is Nothing Then Exit Sub

' Data body access (preferred)
If Not Helpers.HasData(tbl) Then Exit Sub

' Named range access
On Error Resume Next
Set rng = ws.Range(nm)
On Error GoTo 0
If rng Is Nothing Then Exit Sub
```

### Error Handling

```vba
' Public subs - always use Fail block
Public Sub DoWork()
    On Error GoTo Fail
    ' ... code ...
    Exit Sub
Fail:
    Error.TraceErr "Module.DoWork"
End Sub

' Private subs - inline guards or propagate
' Functions - return Empty/Nothing on failure
```

### Column Names

- Live table: Use `Schema.StdChemColName(j)`, `Schema.EnhChemColName(j)`, `Schema.EnhHidColName(j)`
- Constants: Use `Schema.LIVE_COL_*`, `Schema.HISTORY_COL_*`
- Never hardcode column name strings in business logic

### Date Handling

- Type: Always `Date` (not Double)
- Parsing: `Helpers.GetDateVal()` or `CDate()`
- MATCH lookup: `CDbl(targetDate)` for Application.Match
- Arithmetic: Direct addition works (`cfg.StartDate + i`)

## Architecture Reference

See `.claude/docs/architecture.md` for:
- Complete pipeline traces with line numbers
- Wiring verification details
- Integration point documentation
- Table column mappings
