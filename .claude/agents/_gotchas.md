# Project Gotchas

Accumulated learnings. All agents should reference this before making changes.

## Deleted Modules (Do Not Recreate)

| Module | Reason | Alternative |
|--------|--------|-------------|
| Storage.bas | Never integrated, dead code | Use SimLog + History directly |
| EventBus.bas | Over-engineered 5-level indirection | Direct calls in Events.bas |

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
| Chemistry column names | `Schema.ChemistryNames()` = full names, `Schema.ChemShortName(idx)` = short (delegates to `Core.MetricName`) |
| Live table column names | Use `Schema.StdChemColName(idx)`, `Schema.EnhChemColName(idx)`, `Schema.EnhHidColName(idx)` |
| History/SimLog coordination | Share RunId between both for rollback |
| Helper functions location | Put in `Helpers.bas`, not `Schema.bas` (Schema = constants only) |
| Table column lookup | Use `Helpers.ColIdx()` not private copies |
| Date row lookup | Use `Helpers.FindRowByDate()` - O(1) via Application.Match |
| Conditional format priority | Apply Enhanced rule LAST (highest priority) |
| Range access utilities | Use `Helpers.GetRng`, `Helpers.WriteToRange`, `Helpers.ReadFromRange`, `Helpers.GetDateVal` |

## Module Responsibilities

| Module | Contains | Does NOT contain |
|--------|----------|------------------|
| Schema.bas | Constants, ChemistryNames(), column name builders | Helper functions, utilities |
| Helpers.bas | ColIdx, GetSheet, GetTable, GetRng, WriteToRange, ReadFromRange, GetDateVal, FindRowByDate, styling | Business logic, constants |
| Events.bas | Event dispatch, toggle helpers | IR table operations (use Helpers) |
| History.bas | Audit trail, LoadSettings | UI updates (triggers Events) |
| Data.bas | Worksheet I/O, state loading/saving | Duplicate helpers (use Helpers.bas) |

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
- Thin wrapper functions that just delegate to Helpers (deleted Jan 2026)
- EventBus-style indirection for simple event handling (deleted Jan 2026)
- O(n) loops for date lookups when MATCH gives O(1) (fixed Jan 2026)
- Delete/recreate charts every run (causes overflow on consecutive runs)
- Large array allocation for chart data (bind to table columns instead)

## Chart System Gotchas

| Issue | Fix |
|-------|-----|
| Overflow on consecutive chart runs | Don't delete/recreate - update existing charts |
| Chart data doesn't auto-update | Bind to `ListColumn.DataBodyRange` not arrays |
| Charts disappear on site change | Don't delete charts - stack horizontally |
| Chart naming conflicts | Use `cht_{site}_{metric}` pattern |
| Series.Values with large array | Use table column range reference instead |

## Conditional Formatting Order

```vba
' Setup.bas - Correct order (Enhanced has highest priority)
ws.Range("N4:P4").FormatConditions.Delete   ' Clear Enhanced results row
ws.Range("R3:S16").FormatConditions.Delete  ' Clear Enhanced settings
ApplyGreyoutFormat ws.Range("N4:P4"), ...   ' Enhanced results (greyed when Off)
ApplyGreyoutFormat ws.Range("R5:S5"), ...   ' Rain Factor (greyed when Rainfall=Off)
ApplyGreyoutFormat ws.Range("R7:S16"), ...  ' Mixing settings (greyed when Simple)
ApplyGreyoutFormat ws.Range("R3:S16"), ...  ' Highest priority - whole block when Off
```

## History Table Columns (17 total)

```
RunId, Timestamp, RunDate, Days, RainfallMode, TelemCal,
Tau, SurfaceFrac, RainFactor,
StdMode, StdTriggerDay, StdTriggerMetric,
EnhMode, EnhTriggerDay, EnhTriggerMetric,
Action, Load
```

- One row per run (captures both Std and Enh results)
- Enh columns blank when Enhanced disabled
- Action: "Current" or "Rollback"
- Load: Always "Load" (clickable to restore settings)

## Live Table Columns (28 total)

```
Date, Days,
StdVol, StdEC, StdF_U, StdF_Mn, StdSO4, StdMg, StdCa, StdTAN,
EnhVol, EnhEC, EnhF_U, EnhF_Mn, EnhSO4, EnhMg, EnhCa, EnhTAN,
EnhHidEC, EnhHidF_U, EnhHidF_Mn, EnhHidSO4, EnhHidMg, EnhHidCa, EnhHidTAN,
ErrVol, ErrEC, RunId
```

- Days: Relative to run date (0 = today, negative = past, positive = forecast)
- Std[chem]: Standard mode predictions (all 7 chemistry metrics)
- Enh[chem]: Enhanced mode visible layer predictions
- EnhHid[chem]: Enhanced hidden layer mass (TwoBucket continuity)
- ErrVol/ErrEC: Prediction vs telemetry discrepancy
- Row shading: Sample date = light cyan, Run date = light green
- Triggered values: Red + bold formatting

## Inputs Sheet Layout

**J5 Toggle:** Pred_Mode (Std/Enh) - controls which result displays in Predicted row

**N7:O10 Sign Off Block:**
```
N7:  Sign Off header
N8:  Name label       O8: Name dropdown (linked to tblSign)
N9:  Signed label     O9: Signed value
N10: Position label   O10: Position (VLOOKUP from tblSign)
```

**R1:S16 Enhanced Settings Block (greyed when Enhanced=Off):**
```
R1:  Enhanced header
R2:  Enabled          S2: On/Off (toggle)
R3:  Telemetry Cal    S3: On/Off (toggle)
R4:  Rainfall         S4: dropdown (greyed when Enhanced=Off)
R5:  Rain Factor      S5: number (greyed when Rainfall=Off)
R6:  Mixing Model     S6: dropdown
R7:  Tau (days)       S7: number (greyed when Model=Simple)
R8:  Surface Frac     S8: number (greyed when Model=Simple)
R9:  Hidden Mass header (greyed when Model=Simple)
R10-R16: Chemistry labels (greyed when Model=Simple)
S10-S16: Hidden mass values
```
