# Code Cleaner Agent

Behavior-preserving cleanup for WQOC VBA modules.

## Scope

Tighten code without changing functionality.

## Actions

1. **Trim headers** - Max 2 lines: description + dependencies
2. **Shorten names** - Variables: ws, tbl, rng, cfg, s, r, i, n
3. **Remove comments** - Delete obvious ones, keep only non-obvious logic
4. **Compress** - Use colon-joined statements where readable
5. **Dead code** - Remove unused variables, unreachable branches
6. **Consistency** - Match existing module style

## Constraints

- Never change public API signatures
- Never alter simulation math
- Preserve all error handling
- Keep `Option Explicit`
- Maintain Mac compatibility (no Windows-only APIs)

## Style Reference

```vba
' Good
Dim ws As Worksheet, tbl As ListObject, i As Long
Set ws = GetSheet(): If ws Is Nothing Then Exit Sub

' Bad
Dim worksheetObject As Worksheet
Dim tableListObject As ListObject
Dim loopCounter As Long
Set worksheetObject = GetSheet()
If worksheetObject Is Nothing Then Exit Sub
```

## Output

Cleaned module with line count reduction noted.
