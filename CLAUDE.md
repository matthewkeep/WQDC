# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

WQOC (Water Quality Optimisation Calculator) is an Excel/VBA simulation tool for mining wastewater treatment. Models reservoir inflows, mixing, and trigger-based releases.

**Platform:** Excel VBA (Windows & Mac via DictionaryShim)

## Quick Start

```vba
Setup.BuildAll           ' Create sheets, buttons, seed data
Setup.Initialize         ' Create per-site tables/columns from Catalog
WQOC.Run                 ' Run simulation (Standard + Enhanced if enabled)
WQOC.Rollback            ' Undo last run for current site
Tests.RunSmokeSuite      ' 10 smoke tests
Scenarios.RunAll         ' 6 regression scenarios
```

## Architecture

```
WQOC.bas ─┬─ Data.bas ──────── Helpers.bas ── Schema.bas
          ├─ Telemetry.bas ─── Helpers.bas ── Schema.bas
          ├─ Sim.bas ───────── Modes.bas ──── Core.bas
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
| **Events.bas** | Worksheet handlers, double-click toggles, action dispatching |
| **WQOC.bas** | Entry point, orchestration, chart generation |
| **Schema.bas** | Constants only (names, colors, defaults) |
| **Helpers.bas** | Utilities (ColIdx, GetSheet, GetTable, styling, range access) |
| **Setup.bas** | Scaffolding, table creation, dropdowns, conditional formatting |
| **Backtest.bas** | Season replay for A/B comparison |
| **Tests.bas** | Smoke tests |
| **Scenarios.bas** | Regression scenarios |
| **Validate.bas** | Structure validation |
| **DictionaryShim.cls** | Mac compatibility |

### Core Types

```vba
Type State    ' Vol, Chem(1-7), Hidden(1-7), HidVol
Type Config   ' Site, Mode, Days, Tau, Inflow, Outflow, Triggers, RainfallMode, RainFactor, SurfaceFrac
Type Result   ' TriggerDay, TriggerMetric, Snaps(), FinalState
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

**Variables:** `s` = State, `cfg` = Config, `r` = Result, `ws` = Worksheet, `tbl` = ListObject

**Headers:** `Option Explicit` + `' Module: desc` + `' Dependencies: X, Y`

**Helpers:** Use `Helpers.ColIdx`, `Helpers.GetSheet`, `Helpers.GetTable` (not Schema)

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

## Per-Site Architecture

- **Live tables per site:** `tblLive_RP1`, `tblHistory_RP1`
- **Live table structure:** Date-centric with Std/Enh side-by-side
  - Columns: Date, StdVol, StdEC, EnhVol, EnhEC, EnhHid1-7, ErrVol, ErrEC, RunId
  - One row per date (UPSERT on run, not append)
  - Hidden layer stored for TwoBucket continuity between runs
  - Discrepancy columns (ErrVol/ErrEC) compare prediction vs telemetry
- **History table structure:**
  - Columns: RunId, Timestamp, RunDate, Days, Mode, RainfallMode, TelemCal, Tau, SurfaceFrac, RainFactor, TriggerDay, TriggerMetric, Action, Load
  - Action column: "Current" (latest) or "Rollback" (older runs)
  - Load column: Click to restore settings without rollback
- **Telemetry columns per site:** `EC (RP1)`, `Vol (RP1)` (Rain is global)
- **RunId format:** `STD-{site}-{date}-{seq}`, `ENH-{site}-{date}-{seq}`
- **Rollback:** Deletes future data, loads settings, auto-runs simulation
- **Load Settings:** Restores config to Inputs (no deletion, no run)
- **Charts:** Read from tblLive for full season view
- Tables created on-demand (first run) or via `Setup.Initialize`

## Enhanced Mode

- **Rainfall:** Telemetry in mm/day × RainFactor = volume added (ML)
- **Hidden mass:** Auto-loads from tblLive when Sample Date changes
- **Conditional UI:**
  - Enhanced Off → greys all Enhanced settings (N9:O22)
  - Rainfall Off → greys Rain Factor (N11:O11)
  - Mixing Model Simple → greys Tau, Surface Fraction, Hidden Mass (N13:O22)
- **Double-click toggles:** Enabled and Telemetry Cal cells toggle On/Off

## Inputs Sheet Layout (Column N-O)

```
N7:  Enhanced (header)
N8:  Enabled           O8: On/Off (double-click toggle)
N9:  Telemetry Cal     O9: On/Off (double-click toggle)
N10: Rainfall          O10: Off/Hindcast/Hindcast+Forecast
N11: Rain Factor       O11: number (greyed when Rainfall=Off)
N12: Mixing Model      O12: Simple/TwoBucket
N13: Tau (days)        O13: number (greyed when Model=Simple)
N14: Surface Fraction  O14: number (greyed when Model=Simple)
N15: Hidden Mass       (header, greyed when Model=Simple)
N16-N22: Chemistry labels (greyed when Model=Simple)
O16-O22: Hidden mass values
```

## History Table Actions

| Column | Click Action |
|--------|--------------|
| Action ("Rollback") | Delete future runs, load settings, re-run simulation |
| Action ("Current") | Shows message (can't rollback current) |
| Load | Restore settings to Inputs sheet (no deletion, no run) |
