# Upstream PR Analysis Plan for @witan/exceljs Fork

## Executive Summary

This document provides a systematic framework for evaluating 135 open pull requests in the upstream exceljs repository for potential inclusion in our fork. Our fork has already adopted 15+ upstream PRs and developed 3 original features.

## PRs Already Merged into Our Fork

Based on git history, we have already adopted:

### Bug Fixes (Already Merged)

| PR    | Description                                                   | Status    |
| ----- | ------------------------------------------------------------- | --------- |
| #2851 | Fix boolean read val error                                    | ✅ Merged |
| #2956 | Fix dateToExcel() return value for non-numeric values         | ✅ Merged |
| #2973 | Fix parsing error for dynamicFilter nodes in Excel tables     | ✅ Merged |
| #2915 | Fix WorkbookReader sharedString interpretation                | ✅ Merged |
| #2978 | Fix undefined column assignment autofilter                    | ✅ Merged |
| #2803 | Fix corrupted file with conditional formatting and hyperlinks | ✅ Merged |
| #2736 | Improve conditional formatting (stopIfTrue, new operators)    | ✅ Merged |
| #2876 | Fix image reference when same image added non-consecutively   | ✅ Merged |
| #1767 | Add expression support for x14:cfRule                         | ✅ Merged |

### Features (Already Merged)

| PR    | Description                         | Status    |
| ----- | ----------------------------------- | --------- |
| #2885 | Add 'count' metric for pivot tables | ✅ Merged |

### Documentation (Already Merged)

| PR    | Description                                   | Status    |
| ----- | --------------------------------------------- | --------- |
| #2783 | Fix image embedding documentation errors      | ✅ Merged |
| #2733 | Fix broken example code                       | ✅ Merged |
| #2577 | Fix tabColor example and document ARGB format | ✅ Merged |
| #2912 | Fix Chinese docs variable naming              | ✅ Merged |

### From Other Forks (Already Merged)

| Source                 | Description                      | Status    |
| ---------------------- | -------------------------------- | --------- |
| rmartin93/exceljs-fork | Fix getTable().addRow() workflow | ✅ Merged |

---

## Evaluation Framework

### Phase 1: Initial Screening

For each PR, gather:

1. **PR Number and Title**
2. **Type**: Bug Fix / Feature / Enhancement / Maintenance
3. **Complexity**: Low (<50 lines, <3 files) / Medium (50-200 lines, 3-6 files) / High (>200 lines, >6 files)
4. **Streaming Impact**: Does it affect streaming code paths? (Skip if yes)
5. **Test Coverage**: Complete / Partial / None
6. **Community Signal**: Comment count, thumbs up reactions, linked issues

### Phase 2: Deep Evaluation

For PRs that pass Phase 1:

1. **Code Review Checklist**

   - [ ] Code follows existing patterns and style
   - [ ] No security vulnerabilities introduced
   - [ ] No breaking changes to public API
   - [ ] Error handling is appropriate
   - [ ] Edge cases are considered

2. **Test Requirements**

   - [ ] Has unit tests
   - [ ] Has integration tests
   - [ ] Tests cover happy path
   - [ ] Tests cover error cases
   - [ ] Tests are deterministic

3. **Compatibility Check**

   - [ ] Works with our existing modifications
   - [ ] No conflicts with pivot table enhancements
   - [ ] No conflicts with XML escaping fixes

4. **Performance Validation (MANDATORY for all PRs)**
   - [ ] Run `npm run benchmark:nonstreaming` before applying PR
   - [ ] Run benchmark again after applying PR
   - [ ] Compare all 12 benchmarks against baseline
   - [ ] No benchmark shows >10% time regression
   - [ ] No benchmark shows >20% memory regression
   - [ ] Document results in evaluation record

### Phase 3: Integration

1. Cherry-pick or manually apply changes
2. Run full test suite: `npm test`
3. Run specific integration tests
4. Manual testing with sample Excel files
5. Document in FORK.md with attribution

---

## Detailed Evaluation Process

### For Each Candidate PR:

#### Step 1: Fetch PR Details

```bash
gh pr view {PR_NUMBER} --repo exceljs/exceljs
gh pr diff {PR_NUMBER} --repo exceljs/exceljs > pr-{PR_NUMBER}.diff
```

#### Step 2: Analyze Changes

1. Count files changed and lines modified
2. Identify affected components (lib/, spec/, types/)
3. Check for streaming-specific code paths
4. Review test coverage

#### Step 3: Test Locally

```bash
# 0. Establish baseline benchmark (on master)
git checkout master
npm run benchmark:nonstreaming

# 1. Create branch for testing
git checkout -b test-pr-{PR_NUMBER}

# 2. Apply patch
git apply pr-{PR_NUMBER}.diff

# 3. Run tests
npm test

# 4. Run specific integration tests if applicable
npm run test:integration

# 5. Run benchmark (MANDATORY - catches performance regressions)
npm run benchmark:nonstreaming
# Review the "COMPARISON WITH PREVIOUS RUN" output
# Flag any benchmark with >10% time increase or >20% memory increase
```

#### Step 4: Document Findings

Create evaluation record:

```markdown
## PR #{NUMBER}: {Title}

**Decision**: Accept / Reject / Needs Work
**Reason**:

### Changes Summary

- Files: X
- Lines: +Y / -Z

### Test Coverage

- Unit: Yes/No
- Integration: Yes/No
- Missing coverage:

### Benchmark Results

| Benchmark        | Before (ms) | After (ms) | Change | Status   |
| ---------------- | ----------- | ---------- | ------ | -------- |
| read_small_xlsx  | X           | Y          | +Z%    | ✅/⚠️/❌ |
| read_large_xlsx  | X           | Y          | +Z%    | ✅/⚠️/❌ |
| write_large_xlsx | X           | Y          | +Z%    | ✅/⚠️/❌ |
| ...              | ...         | ...        | ...    | ...      |

**Benchmark Status Legend:**

- ✅ = <10% time increase, <20% memory increase
- ⚠️ = 10-25% time increase OR 20-50% memory increase (needs justification)
- ❌ = >25% time increase OR >50% memory increase (reject or fix)

### Risks

-

### Required Modifications

-

### Attribution

Original author: @{username}
```

---

## Risks and Mitigations

### Risk: Breaking Changes

- **Mitigation**: Run full test suite, test with real-world Excel files

### Risk: Performance Regression

- **Mitigation**: Benchmark before/after, especially for #2867

### Risk: Conflicts with Our Modifications

- **Mitigation**: Test pivot table functionality after each merge

### Risk: Incomplete PRs

- **Mitigation**: Write missing tests before merging, attribute work properly

---

## Test Requirements Checklist

Before merging ANY PR:

- [ ] All existing tests pass (`npm test`)
- [ ] New/modified functionality has tests
- [ ] Integration tests with real Excel files work
- [ ] No console errors or warnings
- [ ] Generated Excel files open correctly in:
  - [ ] Microsoft Excel
  - [ ] Google Sheets
  - [ ] LibreOffice Calc

---

## Performance Benchmarking (MANDATORY)

**IMPORTANT**: Run benchmarks for ALL PRs, not just performance-related ones. Any code change can introduce performance regressions.

### Running the Benchmark

```bash
# Via npm script (recommended)
npm run benchmark:nonstreaming

# Or directly
node --expose-gc benchmark-nonstreaming.js
```

The benchmark saves results to `benchmark-results.json` and automatically compares with the previous run.

**Note**: The existing `npm run benchmark` tests STREAMING performance only. For non-streaming API evaluation, always use `benchmark:nonstreaming`.

### Baseline Performance (November 2025)

Established on Node v24.9.0, darwin arm64:

| Benchmark                  | Description                       | Baseline (ms) | Baseline Mem (MB) |
| -------------------------- | --------------------------------- | ------------- | ----------------- |
| `read_small_xlsx`          | Read gold.xlsx (~9KB)             | **6.7**       | 1.1               |
| `read_medium_xlsx`         | Read images.xlsx (~23KB)          | **9.0**       | 1.3               |
| `read_large_xlsx`          | Read huge.xlsx (~14MB, 150K rows) | **4241**      | 905               |
| `write_small_xlsx`         | Write 100 rows, basic             | **13**        | 2.6               |
| `write_medium_styled_xlsx` | Write 1000 rows with styles       | **46**        | 29                |
| `write_large_xlsx`         | Write 10000 rows, 10 columns      | **556**       | 101               |
| `cell_operations`          | 5000 individual cell writes       | **56**        | 31                |
| `merged_cells`             | 500 merge operations              | **80**        | 51                |
| `conditional_formatting`   | 100 CF rules                      | **17**        | 9                 |
| `tables`                   | Create table with 500 rows        | **19**        | 10                |
| `round_trip`               | Read + modify + write             | **18**        | 3                 |
| `formulas`                 | 1000 cells with formulas          | **33**        | 19                |

### Regression Thresholds

A PR should be **flagged for review** if any benchmark shows:

- **Time increase > 10%** from baseline
- **Memory increase > 20%** from baseline

A PR should be **rejected or requires justification** if:

- **Time increase > 25%** on any benchmark
- **Memory increase > 50%** on any benchmark
- **Any new operation is O(n²) or worse** when O(n) is possible

### Benchmark Coverage

**What the benchmark covers:**

- Read operations (small, medium, large files)
- Write operations (various sizes and complexity)
- Cell-level operations with styles
- Merged cells
- Conditional formatting
- Tables with totals/filters
- Formula cells
- Round-trip (read → modify → write)

**What the benchmark does NOT cover (manual testing needed):**

- Pivot tables (test separately with existing pivot table tests)
- Images (covered by integration tests)
- Data validation
- Charts (not supported by library)
- Comments/notes
- Print settings

### Benchmark Process for Each PR

```bash
# 1. Run benchmark on master (baseline)
git checkout master
node --expose-gc benchmark-nonstreaming.js

# 2. Apply PR and run benchmark
git checkout -b test-pr-{NUMBER}
git apply pr-{NUMBER}.diff
node --expose-gc benchmark-nonstreaming.js

# 3. Check comparison output - look for regressions
# The benchmark automatically compares with previous run
```

### Recording Results

Include benchmark comparison in PR evaluation:

```markdown
### Benchmark Results

| Benchmark        | Before (ms) | After (ms) | Change    |
| ---------------- | ----------- | ---------- | --------- |
| read_large_xlsx  | 4241        | 4300       | +1.4% ✅  |
| write_large_xlsx | 556         | 580        | +4.3% ✅  |
| merged_cells     | 80          | 95         | +18.8% ⚠️ |
```

---

## Attribution Template

When merging PRs, use this commit message format:

```
{type}: Adopt upstream PR #{number} - {short description}

{longer description if needed}

Original PR: https://github.com/exceljs/exceljs/pull/{number}
Original author: @{username}

Co-Authored-By: {Full Name} <{email}>
```

---

_Last updated: November 2025_
_Maintainer: @witan/exceljs team_
