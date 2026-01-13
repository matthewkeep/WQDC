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
WQOC.bas ─┬─ Data.bas ──────── Schema.bas
          ├─ Telemetry.bas ─── Schema.bas
          ├─ Sim.bas ───────── Modes.bas ── Core.bas
          ├─ History.bas ───── SimLog.bas
          ├─ Loader.bas
          └─ (Charts)
```

### Modules

| Module | Purpose |
|--------|---------|
| Core.bas | Types: State, Config, Result |
| Modes.bas | StepSimple, StepTwoBucket |
| Sim.bas | Run loop, trigger detection |
| Data.bas | Worksheet I/O (Input/Config/Results) |
| Telemetry.bas | Telemetry data access (Rain, EC, Vol) |
| History.bas | Audit trail, Jenga rollback |
| SimLog.bas | Persistent daily snapshots |
| Loader.bas | Site selection, IR/chemistry population |
| Events.bas | Worksheet change handlers |
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

## Per-Site Architecture

- Log/History tables per site: `tblLog_RP1`, `tblHistory_RP1`
- Telemetry columns per site: `EC (RP1)`, `Vol (RP1)` (Rain is global)
- RunId format: `STD-{site}-{date}-{seq}`, `ENH-{site}-{date}-{seq}`
- Tables created on-demand (first run) or via `Setup.Initialize`
