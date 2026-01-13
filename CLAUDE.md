# CLAUDE.md

## Session Context

**Last updated:** 2026-01-13

**Recent changes:**
- Charts + buttons added (Setup.bas, WQOC.bas)
- Agent system complete (6 agents + foundation)
- All tests passing (10/10 smoke, 6/6 scenarios)

**Paused/Pending:**
- Test new Excel features (buttons, charts)

**Key decisions:**
- Core.bas (not Types/AAATypes) for VBA compile order
- Mass conservation test for TwoBucket (not gradient test)

---

## Project Overview

WQOC (Water Quality Optimisation Calculator) is an Excel/VBA simulation tool for mining wastewater treatment. Models reservoir inflows, mixing, and trigger-based releases.

**Platform:** Excel VBA (Windows & Mac via DictionaryShim)

## Quick Start

```vba
Setup.BuildAll           ' Create sheets, buttons, seed data
WQOC.Run                 ' Run simulation (or click Run button)
WQOC.Rollback            ' Undo last run
Tests.RunSmokeSuite      ' 10 smoke tests
Scenarios.RunAll         ' 6 regression scenarios
```

## Architecture

```
WQOC.bas ─┬─ Data.bas ──── Schema.bas
          ├─ Sim.bas ───── Modes.bas ──── Core.bas
          ├─ History.bas
          └─ (Charts)
```

### Modules (~1,500 lines)

| Module | Purpose |
|--------|---------|
| Core.bas | Types: State, Config, Result |
| Modes.bas | StepSimple, StepTwoBucket |
| Sim.bas | Run loop, trigger detection |
| Data.bas | Worksheet I/O |
| History.bas | Audit trail, rollback |
| WQOC.bas | Entry point + chart generation |
| Schema.bas | Constants, sheet/table names |
| Setup.bas | Scaffolding, buttons, dropdowns |
| Tests.bas | Smoke tests |
| Scenarios.bas | Regression tests |
| Validate.bas | Structure checks |
| DictionaryShim.cls | Mac compatibility |

### Core Types

```vba
Type State    ' Vol, Chem(1-7), Hidden(1-7), HidVol
Type Config   ' Mode, Days, Tau, Inflow, Outflow, Triggers
Type Result   ' TriggerDay, TriggerMetric, Snaps(), FinalState
```

### Flow

```
WQOC.Run → Data.Load → Sim.Run → Modes.Step → Data.Save → History.Record → Charts
```

## Conventions

**Variables:** `s` = State, `cfg` = Config, `r` = Result, `ws` = Worksheet

**Headers:** `Option Explicit` + `' Module: desc` + `' Dependencies: X, Y`

## Extending

- **Add mode:** New `StepX` in Modes.bas
- **Add metric:** Update `METRIC_COUNT` in Core.bas
- **Add trigger:** Update `ChkTriggers` in Sim.bas

## Agent System

See `.claude/agents/` for Claude Code agents:
- **navigator** (`next`) - direction
- **overseer** (`see/plan/repo`) - orchestration
- **cleaner** (`clean`) - code hygiene
- **steward** (`verify`) - integrity checks
- **scout** (`find`) - reconnaissance
- **fixer** (`fix/error`) - debugging
