# Code Cleanup Plan

**Created:** 2026-01-15
**Status:** In Progress
**Estimated Impact:** ~100 lines reduction + performance improvements
**Actual Result:** Focused on high-value, low-risk changes

---

## Phase 1: Dead Code Investigation & Removal

Before deleting, verify each item is truly dead (not unwired).

### 1.1 Data.HasLogDataForDate() [Lines 360-375]
- **Status:** [x] Investigated [x] Confirmed dead [x] Deleted
- **Investigation notes:**
  - Utility function to check if tblLive has data for a date
  - Uses `Helpers.FindRowByDate()` internally - was likely a helper for some planned feature
  - **Zero callers found** - never wired into any pipeline
  - Related function `LoadHiddenFromLog()` already returns empty state if no data, making this check redundant
- **Decision:** DELETE - truly dead, redundant with LoadHiddenFromLog behavior

### 1.2 Core.Config.RainVol field [Line 34]
- **Status:** [x] Investigated [x] Confirmed dead [x] Deleted
- **Investigation notes:**
  - Exists in Config type but NEVER assigned or read anywhere
  - Sim.bas uses `cfg.RainFactor` (line 71) for rain→volume conversion
  - Comment in Sim.bas:45 explains: "RainFactor converts mm to ML"
  - RainVol was likely an earlier design, replaced by RainFactor approach
- **Decision:** DELETE - superseded by RainFactor

### 1.3 Setup.SeedFullSeason() [Lines 504-529]
- **Status:** [x] Investigated [ ] NOT DEAD - keep
- **Investigation notes:**
  - **Developer utility** for generating 90 days of test data
  - Public so it can be called from VBA Immediate window: `Setup.SeedFullSeason`
  - Comment at Setup.bas:426 says "Use SeedFullSeason for comprehensive backtest data"
  - Required for testing `Backtest.RunSeason()` with realistic data
- **Decision:** KEEP - developer/test utility, not dead code

### 1.4 Setup.SeedFullCatalog() [Lines 531-553]
- **Status:** [x] Investigated [ ] NOT DEAD - keep
- **Investigation notes:**
  - Private helper called by SeedFullSeason
  - Creates 2 RR sites with 4 IR sources each for realistic testing
- **Decision:** KEEP - part of SeedFullSeason utility chain

### 1.5 Setup.SeedFullResults() [Lines 555-678]
- **Status:** [x] Investigated [ ] NOT DEAD - keep
- **Investigation notes:**
  - Private helper called by SeedFullSeason
  - Generates 13 weeks of sample data with realistic chemistry trends
- **Decision:** KEEP - part of SeedFullSeason utility chain

### 1.6 Setup.SeedFullTelemetry() [Lines 680-715]
- **Status:** [x] Investigated [ ] NOT DEAD - keep
- **Investigation notes:**
  - Private helper called by SeedFullSeason
  - Seeds 90 days of rain + per-site EC/Vol telemetry
- **Decision:** KEEP - part of SeedFullSeason utility chain

### 1.7 Setup.SeedSiteTelemFull() [Lines 717-771]
- **Status:** [x] Investigated [ ] NOT DEAD - keep
- **Investigation notes:**
  - Public but never called from code - **intentionally manual**
  - NOT duplicate of SeedFullTelemetry - they're complementary:
    - SeedFullTelemetry → seeds global Date + Rain columns
    - SeedSiteTelemFull → seeds per-site EC + Vol columns (requires site columns to exist)
  - Workflow: BuildAll → SeedFullSeason → Initialize → then manually `SeedSiteTelemFull("RP1")`
  - Seeds realistic water balance behavior (vol responds to rain, EC dilutes)
- **Decision:** KEEP - post-Initialize utility for seeding site telemetry test data

---

## Phase 2: Duplicate Code Consolidation

### 2.1 Telemetry.bas - GetLatestEC/GetLatestVol [Lines 42-94]
- **Status:** [x] Consolidated
- **Pattern:** 95% identical, only column name differs
- **Target:** Single `GetLatestValue(site, colNameFunc)` function
- **Lines saved:** ~25

### 2.2 Loader.bas - PopulateRRLatest variants [Lines 85-224]
- **Status:** [x] Consolidated
- **Pattern:** 60% identical, filtered version adds cutoffDate
- **Target:** Single function with optional `cutoffDate` parameter
- **Lines saved:** ~30

### 2.3 SimLog.bas - WriteLiveStandard/WriteLiveEnhanced [Lines 27-136]
- **Status:** [ ] SKIPPED - functions differ significantly (column names, hidden layer, post-processing)
- **Pattern:** 70% identical structure but meaningful differences
- **Notes:** Consolidation would require many flags/parameters, reducing clarity

### 2.4 History.bas - Sort table logic [Lines 120-124, 311-315]
- **Status:** [x] Consolidated - extracted SortHistoryTable helper
- **Pattern:** Identical 4-line sort block repeated 2x
- **Target:** Extract `SortHistoryTable(tbl)` private helper
- **Lines saved:** ~8

### 2.5 Events.bas - ToggleOnOff/ToggleStdEnh [Lines 262-290]
- **Status:** [ ] SKIPPED - ToggleStdEnh has side effect (RefreshPredictedRow)
- **Notes:** Different behavior, not just different values

### 2.6 Setup.bas - Conditional format functions [Lines 1059-1090]
- **Status:** [x] Consolidated - single ApplyGreyoutFormat helper
- **Pattern:** 3 functions with identical RGB/pattern
- **Target:** Single `ApplyGreyoutFormat(rng, condition)` helper
- **Lines saved:** ~20

### 2.7 Data.bas - Getter functions [Lines 7-96]
- **Status:** [ ] SKIPPED - different return types (String, Boolean) and post-processing
- **Notes:** GetPredMode has default logic, GetTelemCalEnabled returns Boolean

### 2.8 Validate.bas - Check functions [Lines 40-93]
- **Status:** [ ] SKIPPED - already minimal (6-8 lines each), check different object types
- **Notes:** Consolidation would add branching complexity, not simplify

### 2.9 Backtest.bas - Table helpers [Lines 424-447]
- **Status:** [ ] SKIPPED - helpers called multiple times, improve readability at call sites
- **Notes:** Good encapsulation, not duplication - replacing with Helpers.GetTable would be more verbose

---

## Phase 3: Verbose Code Simplification

### 3.1 Data.RefreshPredictedRow() [Lines 377-421]
- **Status:** [ ] DEFERRED - complex refactor, higher risk
- **Issue:** 45 lines, 4+ responsibilities
- **Notes:** Works correctly, refactor would be cosmetic

### 3.2 Data.SaveResult() [Lines 198-258]
- **Status:** [ ] DEFERRED - complex refactor, higher risk
- **Issue:** 61 lines, extract offset calculation
- **Notes:** Works correctly, refactor would be cosmetic

### 3.3 WQOC.RunCore() hidden layer init [Lines 93-105]
- **Status:** [ ] SKIPPED - well-commented, clear in context
- **Notes:** Priority logic documented, extraction would reduce readability

### 3.4 WQOC.GetSiteChartLeft() [Lines 234-264]
- **Status:** [ ] SKIPPED - clean encapsulation, inlining would reduce readability
- **Notes:** Separate function handles chart positioning logic clearly

### 3.5 Events.OnInputsDoubleClick() [Lines 65-142]
- **Status:** [ ] DEFERRED - complex refactor, higher risk
- **Issue:** 78 lines, 29 If statements
- **Notes:** Works correctly, refactor would be cosmetic

### 3.6 SimLog column index pre-fetching [Lines 59-62, 116-122]
- **Status:** [x] Optimized
- **Issue:** ColIdx called in inner loop (O(n*7) calls)
- **Target:** Pre-fetch chemistry columns outside loop
- **Result:** Both WriteLiveStandard and WriteLiveEnhanced now pre-fetch all column indices

---

## Phase 4: Pattern Standardization

### 4.1 Create Helpers.GetNamedRange()
- **Status:** [ ] DEFERRED - would add code, not remove
- **Notes:** 5 instances not enough to justify new helper

### 4.2 Standardize table access to Helpers.GetTable()
- **Status:** [ ] SKIPPED - Backtest helpers improve readability at call sites
- **Notes:** See 2.9

### 4.3 Remove unused GenerateCharts parameters
- **Status:** [x] Removed
- **Location:** WQOC.bas line 158 (rStd, rEnh unused)

### 4.4 Consolidate site validation pattern
- **Status:** [ ] SKIPPED - only 3-4 lines each, extraction adds complexity
- **Locations:** WQOC.bas lines 45-48, 136-139

---

## Progress Log

| Date | Action | Result |
|------|--------|--------|
| 2026-01-15 | Created cleanup plan | - |
| 2026-01-15 | Investigated Phase 1 dead code | 5 of 7 items are dev utilities (KEEP), 2 confirmed dead |
| 2026-01-15 | Deleted Data.HasLogDataForDate() | -17 lines |
| 2026-01-15 | Deleted Core.Config.RainVol field | -1 line |
| 2026-01-15 | Consolidated Telemetry GetLatestEC/GetLatestVol | -16 lines |
| 2026-01-15 | Consolidated Loader PopulateRRLatest variants | -33 lines |
| 2026-01-15 | Extracted History.SortHistoryTable helper | -8 lines |
| 2026-01-15 | Consolidated Setup conditional format functions | -22 lines |
| 2026-01-15 | Removed unused GenerateCharts params (rStd, rEnh) | cleaner signature |
| 2026-01-15 | Pre-fetched SimLog column indices | performance: O(n) → O(1) lookups |
| | | |

---

## Notes

- Always run `Tests.RunSmokeSuite` and `Scenarios.RunAll` after each change
- Check that `WQOC.Run`, `WQOC.Rollback` still work
- Verify `Setup.BuildAll` and `Setup.Initialize` still function
