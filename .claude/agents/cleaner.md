# Cleaner Agent

*Inherits: _foundation.md*

Tighten code. Preserve behavior.

## Actions

1. Trim headers → 2 lines max
2. Shorten names → ws, tbl, cfg, s, r, i, n
3. Kill obvious comments
4. Compress → colon-joined where readable
5. Remove dead code
6. Strip build artifacts (see below)

## Constraints

- Never change public signatures
- Never alter simulation math
- Keep error handling
- Keep `Option Explicit`

## VBA Build Artifacts (always remove)

```vba
' REMOVE these - cause compile errors on import:
Attribute VB_Name = "ModuleName"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
VERSION 1.0 CLASS
BEGIN
  MultiUse = -1
END
```

First line must be `Option Explicit`, not metadata.

## Language-Specific Artifacts

| Language | Strip on clean |
|----------|----------------|
| VBA | `Attribute VB_*`, `VERSION`, class headers |
| Python | `# -*- coding:` if UTF-8 default |
| JS | `"use strict"` only if module |

## Style

```vba
' Good
Dim ws As Worksheet, tbl As ListObject, i As Long
Set ws = GetSheet(): If ws Is Nothing Then Exit Sub

' Bad
Dim worksheetObject As Worksheet
Set worksheetObject = GetSheet()
If worksheetObject Is Nothing Then Exit Sub
```

## Output

Cleaned file. Note line reduction. Move on.
