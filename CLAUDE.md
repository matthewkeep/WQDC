# CLAUDE.md

## Project Overview

WQOC (Water Quality Optimisation Calculator) is an Excel/VBA simulation tool for mining wastewater treatment. Models reservoir inflows, mixing, and trigger-based releases.

**Platform:** Excel VBA (Windows & Mac via DictionaryShim)

## Testing

```vba
Setup.BuildAll           ' Create structure + seed data
Tests.RunSmokeSuite      ' 10 smoke tests (pure, no I/O)
Scenarios.RunAll         ' 6 regression scenarios
Validate.Check           ' Structure validation
WQOC.Run                 ' Full simulation
WQOC.TestCore            ' Quick test (no I/O)
```

## Architecture

```
WQOC.bas ─┬─ Data.bas ──── Schema.bas
          ├─ Sim.bas ───── Modes.bas
          ├─ History.bas
          └─ Types.bas
```

### Modules (~1,400 lines total)

| Module | Lines | Purpose |
|--------|-------|---------|
| Types.bas | ~56 | State, Config, Result types |
| Modes.bas | ~82 | StepSimple, StepTwoBucket |
| Sim.bas | ~43 | Run loop, trigger detection |
| Data.bas | ~157 | Worksheet I/O |
| History.bas | ~106 | Audit trail, rollback |
| WQOC.bas | ~97 | Entry point |
| Schema.bas | ~180 | Constants |
| Tests.bas | ~124 | Smoke tests |
| Setup.bas | ~293 | Workbook scaffolding |
| Validate.bas | ~92 | Structure checks |
| Scenarios.bas | ~100 | Regression tests |
| DictionaryShim.cls | ~265 | Mac compatibility |

### Types

```vba
Type State
    Vol As Double
    Chem(1 To 7) As Double
    Hidden(1 To 7) As Double
    HidVol As Double
End Type

Type Config
    Mode As String              ' "Simple" or "TwoBucket"
    Days As Long
    StartDate As Date
    Tau, Inflow, Outflow As Double
    RainVol, SurfaceFrac As Double
    InflowChem(1 To 7) As Double
    TriggerVol As Double
    TriggerChem(1 To 7) As Double
End Type

Type Result
    TriggerDay As Long          ' -1 = no trigger
    TriggerDate As Date
    TriggerMetric As String
    Snaps() As State
    FinalState As State
End Type
```

### Chemistry Metrics

7 metrics (1-7): EC, F_U, F_Mn, SO4, Mg, Ca, TAN

### Flow

```
WQOC.Run()
    Data.LoadState()
    Data.LoadConfig()
    Sim.Run(s, cfg)
        Modes.Step()        ' Simple or TwoBucket
        ChkTriggers()
    Data.SaveResult()
    History.RecordRun()
```

## Conventions

**Variables:** `s` = State, `cfg` = Config, `r` = Result, `ws` = Worksheet, `tbl` = ListObject

**Headers:**
```vba
Option Explicit
' Module: Brief desc.
' Dependencies: X, Y
```

**Sections:** `' ==== Name ====`

**Error handling:** `On Error GoTo Cleanup` with Application state restore

## Extending

- **Add mode:** New `StepX` in Modes.bas, update dispatcher
- **Add metric:** Update `METRIC_COUNT`, `MetricName()`
- **Add config:** Update Config type and Data.LoadConfig
- **Add trigger:** Update Sim.ChkTriggers

## Key Ranges

| Range | Purpose |
|-------|---------|
| RR_InitVol | Initial volume |
| RR_TriggerVol | Volume trigger |
| Res_Row | Latest chemistry |
| Limit_Row | Chemistry triggers |
| RR_HiddenMass | Hidden mass (two-bucket) |
| Cfg_Tau | Mixing time constant |

## Key Tables

| Table | Purpose |
|-------|---------|
| tblIR | Inflow sources |
| tblHistory | Run history |
