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

## This Project

| Issue | Fix |
|-------|-----|
| Chemistry column names | `Schema.ChemistryNames()` returns full names like "EC (uS/cm)", `Core.MetricName()` returns short "EC" |
| History/SimLog coordination | Share RunId between both for rollback |
| Duplicate helpers | Put shared functions in `Utils.bas` |
| Table column lookup | Use `Utils.ColIdx()` not private copies |

## Patterns That Work

- **Error handling**: `On Error GoTo Cleanup` with state restoration
- **Performance**: `Application.ScreenUpdating = False` bracket
- **Tables**: Always check `tbl.DataBodyRange Is Nothing` before access
- **Rollback**: Delete from bottom up to avoid index shift

## Anti-Patterns Found

- Writing same value to two columns (copy-paste error)
- Creating helper functions in multiple modules
- Using Windows-only APIs without checking
- Creating a Utils module for one function (put in existing dependency instead)
