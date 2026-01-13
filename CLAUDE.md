# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

WQOC (Water Quality Dilution Calculator) is an Excel/VBA-based simulation tool for mining wastewater treatment modeling. It models reservoir inflows, rainfall, mixing, and trigger-based releases to help operators plan dilution events and track water quality outcomes.

**Platform:** Excel VBA (Windows & Mac compatible via DictionaryShim)

## Testing

Run in the VBA Immediate Window:
```
Setup.BuildAll           ' Create workbook structure + seed test data
Tests.RunSmokeSuite      ' Run all 10 smoke tests (pure, no worksheet I/O)
WQOC.Run                 ' Run full simulation with worksheet data
WQOC.TestCore            ' Quick core test (no worksheet I/O)
```

Setup commands (standalone, can be removed after testing):
```
Setup.Build              ' Create sheets, tables, named ranges
Setup.Seed               ' Populate test data
Setup.Clean              ' Remove all WQOC sheets (reset)
```

Validation and regression testing:
```
Validate.Check           ' Returns True if workbook structure is valid
Validate.Report          ' Detailed list of missing sheets/ranges/tables
Scenarios.RunAll         ' Run 6 regression scenarios, verify math
Scenarios.RunOne 3       ' Run single scenario by index
```

There is no external build system or CI/CD. Tests are executed manually within Excel.

## Architecture

### Module Structure (6 Core Modules, ~770 lines)

```
ENTRY POINT
    WQOC.bas (~130 lines)     ' One button: WQOC.Run
         │
         ├── Data.bas         ' Load state/config, save results
         ├── Sim.bas          ' Core simulation loop
         ├── Modes.bas        ' Pluggable step functions
         ├── History.bas      ' Audit trail
         └── Types.bas        ' 3 types: State, Config, Result
```

### Core Modules

| Module | Lines | Purpose |
|--------|-------|---------|
| `Types.bas` | ~100 | Type definitions: State, Config, Result |
| `Modes.bas` | ~165 | Pluggable simulation modes: StepSimple, StepTwoBucket |
| `Sim.bas` | ~90 | Core loop: Run(), trigger detection, snapshots |
| `Data.bas` | ~250 | Worksheet I/O: LoadState, LoadConfig, SaveResult |
| `History.bas` | ~165 | Audit trail: RecordRun, RollbackLast |
| `WQOC.bas` | ~170 | Entry point: Run(), Rollback(), TestCore() |

### Supporting Modules

| Module | Purpose |
|--------|---------|
| `Schema.bas` | Constants: sheet names, named ranges, table names |
| `DictionaryShim.cls` | Mac/Windows Dictionary compatibility (reserved) |

### Core Types (Types.bas)

```vba
Type State                    ' Current reservoir state
    Vol As Double             '   Total volume (ML)
    Chem(1 To 7) As Double    '   Concentrations by metric
    Hidden(1 To 7) As Double  '   Hidden mass (two-bucket mode)
    HidVol As Double          '   Hidden volume (two-bucket mode)
End Type

Type Config                   ' Simulation configuration
    Mode As String            '   "Simple" or "TwoBucket"
    Days As Long              '   Forecast days
    StartDate As Date         '   Sample date
    Tau As Double             '   Mixing time constant
    Inflow, Outflow As Double '   Daily flows (ML/d)
    RainVol As Double         '   Daily rain (ML/d)
    InflowChem(1 To 7)        '   Inflow concentrations
    TriggerVol As Double      '   Volume trigger (ML)
    TriggerChem(1 To 7)       '   Chemistry triggers
End Type

Type Result                   ' Simulation output
    TriggerDay As Long        '   Day triggered (-1 if none)
    TriggerDate As Date       '   Date triggered
    TriggerMetric As String   '   Which metric triggered
    Snaps() As State          '   Daily snapshots
    FinalState As State       '   End state
End Type
```

### Chemistry Metrics

7 fixed metrics (indices 1-7): EC, F_U, F_Mn, SO4, Mg, Ca, TAN

### Key Entry Points

- `WQOC.Run` - Main simulation entry (loads from worksheet, runs sim, saves result)
- `WQOC.Rollback` - Undo most recent run
- `WQOC.TestCore` - Quick test without worksheet I/O
- `Tests.RunSmokeSuite` - Run all smoke tests

### Simulation Flow

```
WQOC.Run()
    ├── Data.LoadState()      ' Read initial state from worksheet
    ├── Data.LoadConfig()     ' Read config from worksheet
    ├── Sim.Run(state, cfg)   ' Core simulation loop
    │       └── Modes.Step()  ' Daily step (Simple or TwoBucket)
    │       └── CheckTriggers ' Detect threshold breach
    ├── Data.SaveResult()     ' Write result to worksheet
    └── History.RecordRun()   ' Append to history table
```

## Conventions

### Variable Naming
- `s` = State
- `cfg` = Config
- `r` = Result

### Module Headers
All modules use:
```vba
Option Explicit
' ModuleName: Brief description
' Purpose: What this module does
' Dependencies: What it depends on
```

### Section Headers
Internal organization uses: `' ==== SectionName ====`

### Error Handling
- `On Error GoTo HandleError` with cleanup
- Capture/restore Application state (Calculation, ScreenUpdating, EnableEvents)

### Mac Compatibility
- Use `DictionaryShim.cls` instead of `Scripting.Dictionary`
- Avoid Windows-specific API calls

## Extending the Codebase

1. **Add Simulation Mode:** Add `StepNewMode` function in `Modes.bas`, update dispatcher
2. **Add Metrics:** Update `Types.METRIC_COUNT` and `Types.MetricName()`
3. **Add Config Options:** Update `Config` type and `Data.LoadConfig()`
4. **Add Triggers:** Update `Sim.CheckTriggers()`

## Key Named Ranges (Schema.bas)

| Range | Purpose |
|-------|---------|
| `RR_InitVol` | Initial reservoir volume |
| `RR_TriggerVol` | Volume trigger threshold |
| `Cfg_Tau` | Mixing time constant |
| `Cfg_RainFactor` | Rain volume factor |
| `RR_ResRow` | Latest chemistry concentrations |
| `RR_LimitRow` | Chemistry trigger thresholds |
| `RR_HiddenMass` | Hidden mass for two-bucket mode |

## Key Tables

| Table | Purpose |
|-------|---------|
| `tblIR` | Inflow sources with flow rates and chemistry |
| `tblHistory` | Run history for audit/rollback |
