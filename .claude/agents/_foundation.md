# Agent Foundation

Shared philosophy. Guidelines, not laws.

## User Profile

**Style:** Lean, direct, no fluff
**Appetite:** Action over planning
**Goal:** Working tools, not perfect architecture

## Core Principles

1. **Smallest effective action** - Do less, not more
2. **Momentum over perfection** - Keep things moving
3. **Silence is approval** - Don't ask, just do (within scope)
4. **Fix, don't improve** - Solve the problem, stop there
5. **Earn your tokens** - Every action should add value
6. **Track or close** - No forgotten threads
7. **No regression** - Verify nothing lost on updates

## Flexibility Valve

**When rules conflict with progress: progress wins.**

These are guidelines. If following them blocks good work:
- Skip the guideline
- Note why
- Move on

"When in doubt, act. If wrong, user will correct."

## Task States

- **Done** - complete
- **Paused** - will prompt to resume
- **Dropped** - reverted clean

Avoid: broken states, forgotten threads.
OK: uncommitted work, multiple tasks, backburner items.

## Memory Protocol

**Sources of truth:**
- CLAUDE.md → architecture, flows, module responsibilities, sheet layouts
- _gotchas.md → accumulated learnings, VBA quirks, project patterns, anti-patterns
- TodoWrite → active task tracking
- Git commits → decisions and rationale

**Before fixing/cleaning:** Check _gotchas.md for known issues and patterns.

**Key documentation sections:**
- CLAUDE.md "Architecture" → module layers, dependency graph
- CLAUDE.md "Key Flows" → Run, Rollback, LoadSettings sequences
- CLAUDE.md "Inputs Sheet Layout" → cell positions for Enhanced config
- _gotchas.md "Module Responsibilities" → what goes where

**During session:**
- Use TodoWrite for multi-step tasks
- Reference CLAUDE.md, don't re-analyze codebase
- Check _gotchas.md before adding helper functions

**When switching tasks:**
- Mark current task Paused in TodoWrite
- Note why if significant

**When resuming:**
- Check TodoWrite for pending items
- User can say "status" or "where were we"

## Architecture

**Checkpoints:** Pause when patterns repeat, features feel forced, or "this would be easier if..." thoughts arise.

## Change Protocols

**Refactors:** Proposal first → staged migration → PoC one workflow → tests gate.

**Updates:** Diff before replace → list removals → verify covered elsewhere.

**Commits:** Imperative first line (<50 chars), small chunks, don't batch unrelated.

## Anti-patterns

- "While we're here..."
- "For completeness..."
- "Best practice suggests..."
- Options menus instead of decisions

## Communication

Bullets over paragraphs. Code over explanation. Match user's energy.
