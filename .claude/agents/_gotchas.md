# Project Gotchas

Accumulated learnings. All agents should reference this before making changes.

## VBA Language

| Issue | Fix |
|-------|-----|
| `Log` is reserved (math function) | Use `SimLog`, `AuditLog`, etc. |
| `hhnnss` format typo | Use `hhmmss` for minutes |
| `Scripting.Dictionary` not on Mac | Use `DictionaryShim` class |
| Module name = function name | VBA allows it but causes confusion |

## Excel/VBA Quirks

| Issue | Fix |
|-------|-----|
| Chart SetSourceData with string range | Use `Union()` or separate `.Values`/`.XValues` |
| ListObject column by name case-sensitive | It's not, but be consistent |
| `On Error Resume Next` scope | Always `On Error GoTo 0` after |
| FormatConditions.Delete on subrange | Deletes from overlapping ranges too |

## This Project

| Issue | Fix |
|-------|-----|
| Chemistry column names | `Schema.ChemistryNames()` returns full names like "EC (uS/cm)" |
| History/SimLog coordination | Share RunId between both for rollback |
| Helper functions location | Put in `Helpers.bas`, not `Schema.bas` (Schema = constants only) |
| Table column lookup | Use `Helpers.ColIdx()` not private copies |
| Conditional format priority | Apply Enhanced rule LAST (highest priority) |
| Range access utilities | Use `Helpers.GetRng`, `Helpers.WriteToRange`, `Helpers.ReadFromRange` |

## Module Responsibilities

| Module | Contains | Does NOT contain |
|--------|----------|------------------|
| Schema.bas | Constants, ChemistryNames() | Helper functions, utilities |
| Helpers.bas | ColIdx, GetSheet, GetTable, styling | Business logic, constants |
| Events.bas | Event dispatch, toggle helpers | IR table operations (use Helpers) |
| History.bas | Audit trail, LoadSettings | UI updates (triggers Events) |

## Patterns That Work

- **Error handling**: `On Error GoTo Cleanup` with state restoration
- **Performance**: `Application.ScreenUpdating = False` bracket
- **Tables**: Always check `tbl.DataBodyRange Is Nothing` before access
- **Rollback**: Delete from bottom up to avoid index shift
- **Conditional formatting**: Clear all first, then add rules, Enhanced last
- **Double-click toggle**: Return Boolean, check Intersect, call `ToggleOnOff`
- **LoadSettings**: Write to named ranges, Sample Date change auto-triggers hidden mass load

## Anti-Patterns Found

- Writing same value to two columns (copy-paste error)
- Creating helper functions in multiple modules
- Using Windows-only APIs without checking
- FormatConditions.Delete in each helper (clears overlapping rules)
- Dropdowns for On/Off fields (use double-click toggle instead)

## Conditional Formatting Order

```vba
' Setup.bas - Correct order (Enhanced has highest priority)
ws.Range("N9:O22").FormatConditions.Delete  ' Clear all first
ApplyRainFactorConditionalFormat ...        ' Lowest priority
ApplyMixingConditionalFormat ...            ' Medium priority
ApplyEnhancedConditionalFormat ...          ' Highest priority (applied last)
```

## History Table Columns

```
RunId, Timestamp, RunDate, Days, Mode, RainfallMode, TelemCal,
Tau, SurfaceFrac, RainFactor, TriggerDay, TriggerMetric, Action, Load
```

- Action: "Current" or "Rollback"
- Load: Always "Load" (clickable to restore settings)

## Live Table Columns

```
Date, StdVol, StdEC, EnhVol, EnhEC, EnhHid1-7, ErrVol, ErrEC, RunId
```

- EnhHid1-7: Hidden layer for TwoBucket continuity
- ErrVol/ErrEC: Prediction vs telemetry discrepancy

## Inputs Sheet Layout (N-O)

```
Row 7:  Enhanced header
Row 8:  Enabled (toggle)
Row 9:  Telemetry Cal (toggle)
Row 10: Rainfall (dropdown)
Row 11: Rain Factor (greyed when Rainfall=Off)
Row 12: Mixing Model (dropdown)
Row 13: Tau (greyed when Model=Simple)
Row 14: Surface Fraction (greyed when Model=Simple)
Row 15: Hidden Mass header (greyed when Model=Simple)
Row 16-22: Hidden mass values (greyed when Model=Simple)
```
