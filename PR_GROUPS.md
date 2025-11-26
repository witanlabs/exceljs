# Upstream PR Groups for @witan/exceljs Fork

This document groups related PRs that implement the same or overlapping functionality, allowing us to evaluate them together and take the best approach from each.

---

## Legend

- **⭐ Recommended**: Best candidate in the group
- **🔄 Alternative**: Valid alternative approach
- **⚠️ Has Issues**: Known problems to address
- **📅 Stale**: No activity >2 years, may need updating

---

## Group 1: Cell Style Mutation / Cloning ✅ DONE

**Problem**: When cells share style objects, modifying one cell's style affects others.

| PR           | Title                                               | Tests   | Notes                             |
| ------------ | --------------------------------------------------- | ------- | --------------------------------- |
| **#2781** ⭐ | Fix cell style cloning (setters clone style object) | ✅ Yes  | Cleanest fix - clones on setter   |
| #1378        | Fix cell style modification affecting other cells   | Partial | Adds setStyle() helper method     |
| #1573        | Fix cell.style.fill problems                        | ✅ Yes  | Similar to #2781, focused on fill |

**Decision**: ✅ **ADOPTED** - Implemented PR #2781's approach (style cloning on setters).

**Analysis Summary**:
- PR #2781 is the cleanest solution - modifies all 6 style setters to clone the style before mutation
- PR #1378 only provides a workaround (`setStyle()` method) requiring users to change their code
- PR #1573 only addresses the model setter, not the property setters

**Implementation**: Added `_applyStyle()` helper method to Cell class that clones style via existing `copyStyle` utility before setting property. All 6 setters (numFmt, font, alignment, border, fill, protection) now use this method.

**Tests**: Added 7 unit tests verifying style cloning behavior for all setters.

**Benchmark**: No regression - all benchmarks within ±6% (natural variance).

---

## Group 2: Merged Cells Performance ✅ DONE

**Problem**: Checking merged cell conflicts is O(n²), extremely slow with many merges.

| PR           | Title                                   | Tests  | Notes                                         |
| ------------ | --------------------------------------- | ------ | --------------------------------------------- |
| #2691        | Fix inefficient merge check             | Manual | ⚠️ Has syntax errors (isMerged() vs isMerged) |
| **#2920** ⭐ | Fix inefficient merge check (corrected) | Manual | Fixes #2691's syntax errors                   |

**Decision**: ✅ **ADOPTED** - Implemented PR #2920's approach (O(n) cell-based merge checking).

**Analysis Summary**:
- The original implementation iterated through ALL existing merges for each new merge operation - O(n²) complexity
- PR #2920 changes this to check individual cells directly via `cell.isMerged` property - O(m) where m is merge area
- PR #2691 had syntax errors (`isMerged()` vs `isMerged`), which PR #2920 corrects

**Performance Results** (10,000 merges stress test):
| Merges | Old O(n²) | New O(n) | Speedup |
|--------|-----------|----------|---------|
| 500    | 14ms      | 3ms      | 4.7x    |
| 1000   | 84ms      | 8ms      | 10.5x   |
| 2000   | 220ms     | 9ms      | 24x     |
| 5000   | 1499ms    | 14ms     | 107x    |
| 10000  | 6464ms    | 34ms     | **190x** |

**Benchmark**: Standard merged_cells benchmark shows -10% time improvement (137ms → 123ms) and -7% memory improvement.

**Tests**: All 4 merge-related tests pass (overlapping merges prevention, merge/unmerge, style merging).

**Implementation**: Modified `_mergeCellsInternal()` in `lib/doc/worksheet.js` to check `cell.isMerged` property directly instead of iterating through all existing merges.

---

## Group 3: Shapes and Drawing Objects

**Problem**: No support for shapes, text boxes in ExcelJS.

| PR           | Title                            | Tests   | Notes                                          |
| ------------ | -------------------------------- | ------- | ---------------------------------------------- |
| #2077        | Basic shape support              | Partial | 3 approvals, 11 comments, original impl        |
| **#2601** ⭐ | Shapes and text boxes (enhanced) | ✅ Yes  | Built on #2077, adds TypeScript, more features |

**Recommendation**: Evaluate #2601 - it's more complete with better test coverage and TypeScript types. May want to compare implementations for any features #2077 has that #2601 doesn't.

---

## Group 4: Table Loading and Corruption Fixes ✅ DONE

**Problem**: Tables don't load correctly from disk, cause corruption when modified.

| PR           | Title                                            | Tests   | Notes                               |
| ------------ | ------------------------------------------------ | ------- | ----------------------------------- |
| **#1222** ⭐ | Load tables correctly when reading from disk     | Partial | ⚠️ Needs multi-char column fix      |
| #1345        | Improved table handling                          | Partial | ⚠️ Has failing tests, broader scope |
| #1938        | Fix loading of table rows and ref from xlsx      | Unknown | Similar to #1222                    |
| #1936        | Fix loading of tables with calculated columns    | Unknown | Calculated column specific          |
| **#2089** ⭐ | Fix table corruption when reading/modifying XLSX | ✅ Yes  | Focuses on corruption prevention    |
| #2090        | Tables/media update on row/column insert/delete  | ✅ Yes  | Reference updates during splice     |

**Decision**: ✅ **ADOPTED** - Combined approach from PRs #1936, #2089, #1222, #1938

**Analysis Summary**:
- PR #1936: Simple fix for `parseClose()` in table-column-xform.js to properly handle calculated columns
- PRs #2089/#1222/#1938: All address the same core issue - tables loaded from xlsx have empty `rows` arrays
- PR #2090: Larger scope enhancement for splice operations - **DEFERRED** for separate evaluation
- PR #1345: Too broad scope with failing tests - **REJECTED**

**Implementation**:
1. Fixed `parseClose()` in `table-column-xform.js` to return `name !== this.tag` instead of `false`
2. Added `_loadRowsFromWorksheet()` method to Table class that populates `rows` from worksheet cells
3. Modified Table constructor to accept `{isLoading: true}` option that triggers row loading instead of `store()`
4. Updated worksheet.js to pass `{isLoading: true}` when loading tables from xlsx

**Results**:
- Tables loaded from xlsx now have their `rows` array properly populated from worksheet cell data
- `table.addRow()` works correctly on loaded tables (appends to existing rows)
- Worksheet cell data is preserved (not overwritten with empty data)
- Tables with header rows and/or totals rows handled correctly

**Tests**: Added 4 integration tests verifying table loading behavior.

**Benchmark**: No regression - all benchmarks within normal variance (±5%)

---

## Group 5: spliceRows / spliceColumns Fixes

**Problem**: spliceRows() doesn't work correctly for deletions, especially at end of worksheet.

| PR           | Title                                           | Tests   | Notes                                   |
| ------------ | ----------------------------------------------- | ------- | --------------------------------------- |
| **#1600** ⭐ | Fix spliceRows() issues with delete             | ✅ Yes  | 12 comments, well-documented            |
| #2625        | Fix spliceColumn mergedcells                    | Unknown | Merged cells during splice              |
| #2090        | Tables/media update on row/column insert/delete | ✅ Yes  | Overlaps - handles tables during splice |

**Recommendation**: #1600 first for core fix, then #2090 for table awareness during splice.

---

## Group 6: Rich Text / Shared Strings

**Problem**: Rich text parsing issues, especially with shared strings.

| PR    | Title                                   | Tests   | Notes                                |
| ----- | --------------------------------------- | ------- | ------------------------------------ |
| #2737 | Don't render empty rich text substrings | Unknown | Simple correctness fix               |
| #2588 | Fix shared strings and richText         | Unknown | SharedStrings + richText interaction |
| #2001 | Fix rich text tags not closed           | Unknown | Parser compatibility                 |

**Recommendation**: Evaluate all three - they may fix different aspects of the same subsystem.

---

## Group 7: Image Handling

**Problem**: Various image-related bugs and missing features.

| PR           | Title                                                | Tests   | Notes                            |
| ------------ | ---------------------------------------------------- | ------- | -------------------------------- |
| **#2983** ⭐ | Add ImageEditAs type ('twoCell' option)              | ✅ Yes  | Image positioning with filtering |
| #2614        | Fix addImage position is wrong                       | Unknown | Position calculation bug         |
| #2924        | Fix Anchor Column/Row Positioning                    | Unknown | Similar positioning issue        |
| #2782        | Add option for note height/width                     | Unknown | Comment box sizing               |
| #1789        | Add image accessibility requirements                 | Unknown | Alt text, etc.                   |
| #1448        | Add addImage method validation                       | Unknown | Input validation                 |
| #2049        | Improved media properties (rotation, extent, offset) | Unknown | More image options               |
| #2903        | Add support for removing images                      | Unknown | removeImage() method             |
| #2630        | Add additional Media types for worksheet/workbook    | Unknown | More media type support          |

**Recommendation**:

1. #2983 first (well-tested feature)
2. Evaluate #2614 vs #2924 for positioning fixes (may be same issue)
3. #2903 for removeImage if needed

---

## Group 8: Conditional Formatting

**Problem**: Various CF edge cases and missing features.

| PR    | Title                                           | Tests   | Notes           |
| ----- | ----------------------------------------------- | ------- | --------------- |
| #2095 | Fix saving worksheet with CF breaks font styles | Unknown | 📅 Stale (2022) |
| #2655 | Add color field in data bar CF (TypeScript)     | N/A     | Types only      |

**Note**: We already merged #2803, #2736, #1767 for major CF fixes.

---

## Group 9: Date Handling

**Problem**: Date parsing issues in various formats.

| PR           | Title                                            | Tests   | Notes                   |
| ------------ | ------------------------------------------------ | ------- | ----------------------- |
| **#2702** ⭐ | Fix date parsing for Strict OpenXML spreadsheets | Partial | Important compatibility |
| #1796        | Make date1904 property optional                  | Unknown | Type/options fix        |

**Recommendation**: #2702 first for Strict OpenXML, then evaluate #1796.

---

## Group 10: Data Validation

**Problem**: Data validation edge cases.

| PR    | Title                                              | Tests   | Notes           |
| ----- | -------------------------------------------------- | ------- | --------------- |
| #2977 | Fix large validation ranges (clamp to actual data) | Unknown | Range handling  |
| #2697 | Add DataValidationType type                        | N/A     | TypeScript only |

---

## Group 12: AutoFilter

**Problem**: AutoFilter bugs and missing features.

| PR    | Title                                        | Tests   | Notes            |
| ----- | -------------------------------------------- | ------- | ---------------- |
| #1383 | Add support for excluding AutoFilter columns | Unknown | Feature addition |

**Note**: We already merged #2978 for autofilter fix.

---

## Group 13: Row/Column Operations

**Problem**: Various row/column manipulation issues.

| PR    | Title                                           | Tests   | Notes               |
| ----- | ----------------------------------------------- | ------- | ------------------- |
| #2116 | Duplicate multiple rows feature                 | Unknown | Utility feature     |
| #1563 | Col listed in different order breaks props load | Unknown | Column ordering bug |

---

## Group 14: Worksheet Properties

**Problem**: Missing or incorrect worksheet properties.

| PR    | Title                                  | Tests   | Notes                 |
| ----- | -------------------------------------- | ------- | --------------------- |
| #2800 | Fix worksheet-reader hidden prop       | Unknown | Hidden sheet property |
| #2102 | Fix outlineProperties doesn't exist    | Unknown | Outline settings      |
| #1971 | Fix multiple print area functionality  | Unknown | Print areas           |
| #1516 | Fix multiple print areas serialization | Unknown | Similar to #1971      |
| #2807 | Fix pageSetUpPr order (broken xlsx)    | Unknown | XML ordering          |

**Recommendation**: #2807 first (likely simple), then evaluate others.

---

## Group 15: Comments/Notes

**Problem**: Comment/note handling issues.

| PR    | Title                              | Tests   | Notes                    |
| ----- | ---------------------------------- | ------- | ------------------------ |
| #2079 | Fix note in cell within table      | Unknown | Table + note interaction |
| #1746 | Add removeNote method              | Unknown | Note removal             |
| #1933 | Add automatic size for comment box | Unknown | Auto-sizing              |

---

## Group 16: Formulas

**Problem**: Formula handling issues.

| PR    | Title                                          | Tests   | Notes               |
| ----- | ---------------------------------------------- | ------- | ------------------- |
| #2883 | Make to work with expressions with no formulae | Unknown | Expression handling |
| #2264 | Fix displaying >255 characters in formula      | Unknown | Long formula bug    |

---

## Group 17: Hyperlinks

**Problem**: Hyperlink handling issues.

| PR    | Title              | Tests   | Notes                   |
| ----- | ------------------ | ------- | ----------------------- |
| #2002 | Fix hyperlink hash | Unknown | Hash character handling |

**Note**: We already merged #2803 which fixes CF + hyperlinks corruption.

---

## Group 18: XML/Parser Issues

**Problem**: XML parsing edge cases and compatibility.

| PR           | Title                                          | Tests   | Notes                           |
| ------------ | ---------------------------------------------- | ------- | ------------------------------- |
| **#2962** ⭐ | Fix missing `r` attribute in row/cell elements | ✅ Yes  | DataGrip/IntelliJ compatibility |
| #2894        | Fix parse-sax.js broke utf8 string             | Unknown | UTF-8 parsing                   |
| #2081        | Fix utf-8 multibyte character garbled          | Unknown | Similar UTF-8 issue             |
| #1665        | Fix parse-sax destructs multi-byte char        | Unknown | Similar UTF-8 issue             |
| #2846        | Update xlsx.js for non-office generated files  | Unknown | Compatibility                   |
| #2185        | Fix file not opening because of wrong defaults | Unknown | Default values                  |
| #2852        | Fix empty target on worksheet-xform reconcile  | Unknown | Empty target check              |
| #2341        | Add test with third party exported excel file  | Unknown | Compatibility testing           |

**Recommendation**: #2962 first (well-tested), then evaluate UTF-8 fixes together (#2894, #2081, #1665 may be duplicates).

---

## Group 19: Performance Optimizations ✅ DONE

**Problem**: Performance bottlenecks.

| PR           | Title                                 | Tests   | Notes                  |
| ------------ | ------------------------------------- | ------- | ---------------------- |
| **#2867** ⭐ | styleCacheMode - Up to 3x performance | ✅ Yes  | Major perf improvement |
| #1929        | Release large values early for GC     | Unknown | Memory optimization    |
| #2691/#2920  | Efficient merge check                 | Manual  | See Group 2            |

**Decision**: ✅ **ADOPTED** - Implemented both PR #2867 (styleCacheMode) and PR #1929 (GC optimization)

**Analysis Summary**:
- PR #2867 (styleCacheMode): Introduces configurable style caching with 4 modes (WEAK_MAP, JSON_MAP, FAST_MAP, NO_CACHE)
  - FAST_MAP provides ~33% performance improvement for style-heavy workbooks
  - Works by comparing styles by value (string serialization) instead of object reference
  - Backward compatible - defaults to WEAK_MAP (original behavior)
- PR #1929 (GC optimization): Simple memory optimization that releases large buffers early
  - Nullifies `chunks`, `buffer`, and `zip` references after they're no longer needed
  - Helps garbage collector reclaim memory sooner during large file processing
  - Low risk, minimal code changes
- PR #2691/#2920: Already adopted in Group 2

**Implementation**:
1. Created `lib/utils/style-fast-serialize.js` with fast style serialization functions
2. Modified `lib/xlsx/xform/style/styles-xform.js` to support 4 cache modes
3. Added `stylesCacheMode` option to workbook write operations
4. Exported `StyleCacheMode` from main ExcelJS module
5. Added memory release in `lib/xlsx/xlsx.js` for GC optimization

**Benchmark Results** (heavy_styles benchmark - 10000 cells):
| Mode | Time (ms) | Change |
|------|-----------|--------|
| WEAK_MAP (default) | 186 | baseline |
| FAST_MAP | 126 | -32% ✅ |

**Tests**: Added 18 unit tests for style-fast-serialize module. All 910 unit tests pass.

---

## Group 20: TypeScript Types

**Problem**: Type definition improvements.

| PR    | Title                                  | Tests | Notes            |
| ----- | -------------------------------------- | ----- | ---------------- |
| #2562 | Fix typescript and intellisense        | N/A   | 10 comments      |
| #2697 | Add DataValidationType                 | N/A   | Validation types |
| #2655 | Add color field in data bar CF         | N/A   | CF types         |
| #2664 | Worksheet protect() fix types          | N/A   | Protection types |
| #2596 | Fix dimensions type                    | N/A   | Dimensions       |
| #2587 | Fix Row/Column values types            | N/A   | Row/Column       |
| #2720 | Fix type mismatch in Address interface | N/A   | Address          |
| #1886 | Fix Row.values type definition         | N/A   | Row values       |
| #1901 | Fix worksheet reader type definitions  | N/A   | Reader types     |
| #1922 | WorkbookReader options optional type   | N/A   | Options          |

**Recommendation**: Bundle TypeScript fixes together in one evaluation pass.

---

## Group 21: Pivot Tables

**Problem**: Pivot table enhancements.

| PR    | Title                              | Tests  | Notes                 |
| ----- | ---------------------------------- | ------ | --------------------- |
| #2578 | PivotTable multiple values support | ✅ Yes | Multiple value fields |

**Note**: We have our own pivot table enhancements (#2995, #2996, #2997). Evaluate #2578 for potential merge.

---

## Group 22: Header/Footer

**Problem**: Header/footer features.

| PR    | Title                           | Tests   | Notes                                |
| ----- | ------------------------------- | ------- | ------------------------------------ |
| #2563 | Header and footer support image | Partial | ⚠️ Compatibility issues, 24 comments |

---

## Group 23: Page Breaks

**Problem**: Page break handling.

| PR    | Title             | Tests   | Notes            |
| ----- | ----------------- | ------- | ---------------- |
| #2602 | Parse page breaks | Unknown | Read page breaks |

---

## Group 24: Table Styling

**Problem**: Table style issues.

| PR    | Title                                     | Tests   | Notes            |
| ----- | ----------------------------------------- | ------- | ---------------- |
| #2061 | 'None' is another theme style for a table | Unknown | Theme handling   |
| #1907 | Fix assignStyle on Table.store            | Unknown | Style assignment |
| #2767 | Table creation accepts invalid names      | Unknown | Validation       |
| #2680 | Table creation allows empty array of rows | Unknown | Validation       |

---

## Group 25: Nested/Hierarchical Columns

**Problem**: No support for nested column headers.

| PR    | Title                          | Tests   | Notes                      |
| ----- | ------------------------------ | ------- | -------------------------- |
| #1889 | Support nested columns feature | Partial | 18 comments, high interest |

---

## Group 26: Themes

**Problem**: Theme handling.

| PR    | Title         | Tests   | Notes         |
| ----- | ------------- | ------- | ------------- |
| #2009 | Fix addThemes | Unknown | Theme loading |

---

## Group 27: Security/Dependencies

**Problem**: Dependency vulnerabilities and issues.

| PR    | Title                                 | Tests | Notes                     |
| ----- | ------------------------------------- | ----- | ------------------------- |
| #2744 | Bump unzipper (license issue)         | N/A   | 20 comments, pre-released |
| #2869 | Bump unzipper (duplicate)             | N/A   | Duplicate of #2744        |
| #2999 | Remove critical vulnerabilities       | N/A   | Security fixes            |
| #2989 | Fix Snyk transitive dependencies      | N/A   | Archiver issue            |
| #2891 | Dependencies bump and code fix        | N/A   | Mixed                     |
| #2687 | Replace unzipper with yauzl-promise   | N/A   | ⚠️ Breaking change        |
| #2672 | Fix unsafe-eval CSP issue             | N/A   | Uses patch-package        |
| #2278 | Downgrade regenerator-runtime for CSP | N/A   | CSP compliance            |
| #2710 | Add proper version control to deps    | N/A   | Version pinning           |
| #2812 | Update dependency version             | N/A   | Version update            |

**Recommendation**: Evaluate security PRs together. #2744 seems most active.

---

## Group 28: Race Conditions / Async Issues

**Problem**: Race conditions in async code.

| PR    | Title                                       | Tests   | Notes          |
| ----- | ------------------------------------------- | ------- | -------------- |
| #2874 | Fix error-prone race conditions             | Unknown | Async fixes    |
| #2698 | style.xml has [Object object] as formatCode | Unknown | May be related |

---

## Group 29: Cell Formatting Features

**Problem**: Missing cell formatting options.

| PR    | Title                                     | Tests   | Notes                     |
| ----- | ----------------------------------------- | ------- | ------------------------- |
| #2809 | Add quote prefix feature                  | Unknown | Cell quote prefix support |
| #1061 | Add ability to set fill style for a range | Unknown | Range fill helper         |

---

## Group 30: Miscellaneous Bug Fixes

**Problem**: Various standalone fixes.

| PR    | Title           | Tests   | Notes            |
| ----- | --------------- | ------- | ---------------- |
| #2651 | Fix issue #2547 | Unknown | Specific bug fix |
| #1688 | Fix issue #894  | Unknown | Specific bug fix |

---

## Excluded: Streaming-Only PRs

These PRs only affect streaming code paths (we use non-streaming):

- #2849 - Web-native streams
- #3002 - Merged cells in stream processing
- #2558 - Fix file writing using streams
- #1570 - Fix streaming parser bugs
- #2685 - Fix streaming with autofilter/sheet protection
- #2791 - xlsx stream missing worksheets
- #1457 - Browser streaming xlsx reader
- #1885 - Table support for streaming mode
- #2633, #2201 - WorksheetWriter.addImage
- #2148 - stream.xlsx cleanup
- #2671 - iterate-stream resilience
- #1596 - StreamBuf piping
- #1531 - Fix default options for streaming xlsx reader (docs)

---

## Excluded: Already Merged

- #2998 - Fix getTable().addRow() ✅ (from rmartin93 fork)
- #2851 - Boolean read fix ✅
- #2956 - dateToExcel fix ✅
- #2973 - dynamicFilter fix ✅
- #2915 - WorkbookReader fix ✅
- #2978 - Autofilter fix ✅
- #2803 - CF + hyperlinks fix ✅
- #2736 - CF improvements ✅
- #2876 - Image fix ✅
- #1767 - x14:cfRule ✅
- #2885 - Pivot count metric ✅
- #2783, #2733, #2577, #2912 - Docs ✅

---

## Excluded: CSV-Only (Not Used)

- #2752 - Fix CSV cells with spaces converted to 0
- #1743 - Fix CSV reading large numbers
- #2080 - Fix reading CSV files with headers
- #2127 - Fix FastCsvParserOptionsArgs type

---

## Excluded: Unnecessary Syntactic Sugar

- #2991 - Add getFirstWorksheet() method

---

## Excluded: Trivial/Garbage

- #3003 - Fix typo in comment
- #2847 - Unclear patch
- #2930 - content-types.01.xml update
- #2779 - Add debug logs
- #1869 - Lint fixes only
- #1346 - Rename "slave" to "aligned" (terminology)
- #1664 - Remove 'use strict' (questionable)
- #2688 - Update package name

---

_Last updated: November 2025_
