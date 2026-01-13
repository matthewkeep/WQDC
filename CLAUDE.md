# CLAUDE.md

## Project Overview

WQOC (Water Quality Optimisation Calculator) is an Excel/VBA simulation tool for mining wastewater treatment. Models reservoir inflows, mixing, and trigger-based releases.

**Platform:** Excel VBA (Windows & Mac via DictionaryShim)

## Testing

```vba
Tests.RunSmokeSuite      ' 10 smoke tests
Scenarios.RunAll         ' 6 regression scenarios
WQOC.Run                 ' Full simulation
WQOC.Rollback            ' Undo last run
```

## Architecture

```
WQOC.bas ─┬─ Data.bas ──── Schema.bas
          ├─ Sim.bas ───── Modes.bas ──── Core.bas
          └─ History.bas
```

### Modules (~1,300 lines)

| Module | Lines | Purpose |
|--------|-------|---------|
| Core.bas | 55 | Types: State, Config, Result |
| Modes.bas | 81 | StepSimple, StepTwoBucket |
| Sim.bas | 42 | Run loop, trigger detection |
| Data.bas | 156 | Worksheet I/O |
| History.bas | 105 | Audit trail, rollback |
| WQOC.bas | 96 | Entry point |
| Schema.bas | 178 | Constants |
| Tests.bas | 125 | Smoke tests |
| Setup.bas | 291 | Workbook scaffolding |
| Validate.bas | 91 | Structure checks |
| Scenarios.bas | 98 | Regression tests |
| DictionaryShim.cls | 265 | Mac compatibility |

### Core Types

```vba
Type State                  ' Reservoir state
    Vol As Double
    Chem(1 To 7) As Double
    Hidden(1 To 7) As Double
    HidVol As Double
End Type

Type Config                 ' Simulation config
    Mode As String          ' "Simple" or "TwoBucket"
    Days, Tau, Inflow, Outflow As Double
    TriggerVol, TriggerChem(1 To 7) As Double
End Type

Type Result                 ' Simulation output
    TriggerDay As Long      ' -1 = no trigger
    TriggerMetric As String
    Snaps() As State
End Type
```

### Chemistry Metrics

7 metrics: EC, F_U, F_Mn, SO4, Mg, Ca, TAN

### Flow

```
WQOC.Run → Data.Load → Sim.Run → Modes.Step → Data.Save → History.Record
```

## Conventions

**Variables:** `s` = State, `cfg` = Config, `r` = Result, `ws` = Worksheet

**Headers:**
```vba
Option Explicit
' Module: Brief desc.
' Dependencies: X, Y
```

## Extending

- **Add mode:** New `StepX` in Modes.bas
- **Add metric:** Update `METRIC_COUNT` in Core.bas
- **Add trigger:** Update `ChkTriggers` in Sim.bas
