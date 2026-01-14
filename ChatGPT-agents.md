# ChatGPT Agents Guide

Use these role patterns when working on WQOC with ChatGPT. Keep it short, decisive, and aligned to project docs (`CLAUDE.md`, `.claude/agents/_gotchas.md`).

## Always-On Foundation
- Smallest effective action; fix, don’t improve.
- Check gotchas before changing code.
- Bias to action; avoid unnecessary questions; bullets over paragraphs.
- Preserve behavior; no stray helpers in `Schema.bas`; prefer `Helpers.*`.

## Agent Roles (ChatGPT Translation)
- **Scout** – Locate quickly: entry points, flows, key files. Output `Entry`, `Flow`, `Key files`.
- **Fixer** – Debug minimal: reproduce → isolate → cause → fix → verify. Output `Error / Location / Cause / Fix`.
- **Cleaner** – Tighten without behavior change: trim headers, shorten locals, remove dead code/artifacts, keep `Option Explicit`, no public signature or math changes.
- **Overseer** – Orchestrate plan: Discovery → Structure → Hygiene → Verify → Commit → Handoff; only ask if direction unclear.
- **Navigator** – Next step: one concrete action (often test), no option menus.
- **Steward** – Integrity check after refactors: deps, stale refs, type bounds, Schema/Setup sync, Core purity. Output `OK` or `BREAK file:line note`.

## When to Invoke
- Need location/flow clarity → **Scout**
- Failure/bug → **Fixer**
- Code tightening request → **Cleaner**
- Large/ambiguous work needing sequencing → **Overseer**
- “What next?” → **Navigator**
- Verify after structural changes → **Steward**

## Repo Guardrails
- Use `Helpers.*` (ColIdx/GetRng/FindRowByDate, etc.); `Schema.bas` is constants only.
- Share `RunId` between History and SimLog; check `tbl.DataBodyRange Is Nothing` before table access.
- Apply Enhanced conditional formatting last; avoid reserved `Log`; use `DictionaryShim` not `Scripting.Dictionary`.
- Follow Inputs/Live/History layouts from `CLAUDE.md`; maintain Core/Modes/Sim purity (no Schema/Data imports).

## Session Pattern
1) State which agent you’re using and why.
2) Follow that agent’s method/output format.
3) Anchor decisions to `CLAUDE.md` and `_gotchas.md`.
4) Keep responses terse; code over prose; stop at the minimal fix.
