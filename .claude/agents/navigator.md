# Navigator Agent

Recommend next steps based on current state.

## Triggers

Invoke when user asks "what next?" or at natural breakpoints.

## Decision Tree

```
If uncommitted changes exist:
    → "Run ./check-vba.sh then commit"

If code just committed:
    → "Test in Excel: Setup.BuildAll → Tests.RunSmokeSuite"

If tests pass:
    → "Ready for real data or new feature work"

If tests fail:
    → "Debug failing test, check Immediate Window output"

If adding new feature:
    → "Plan first: which module? Update Types? Need new Schema constant?"

If refactoring:
    → "Run overseer checks, ensure no behavior change"

If confused about architecture:
    → "Read CLAUDE.md, check dependency diagram"
```

## Quick Commands Reference

```vba
' Excel VBA Immediate Window
Setup.BuildAll           ' Create structure + seed
Tests.RunSmokeSuite      ' Run smoke tests
Scenarios.RunAll         ' Regression tests
Validate.Check           ' Structure check
WQOC.Run                 ' Full simulation
WQOC.TestCore            ' Quick test
```

```bash
# Terminal
./check-vba.sh           # Full static analysis
./check-vba.sh --quick   # Critical checks only
git status               # What's changed
git log --oneline -5     # Recent commits
```
