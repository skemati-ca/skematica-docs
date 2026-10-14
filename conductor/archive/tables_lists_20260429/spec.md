# Specification: Tables & Lists

## Overview

Add table creation and modification tools, expose list/numbering manipulation that is partly available internally today, and refactor `xml-utils` to use `preserveOrder:true` to fix body/table interleaving fidelity.

Tables are the largest remaining gap for commercial and academic documents (price tables, schedules, bibliographies, comparative analyses). Lists already have internal infrastructure (`getParagraphListInfo`) but no MCP-level tool.

## Technical Approach

- **`preserveOrder:true` migration:** the current parser config loses ordering between body paragraphs and tables. This is documented as a limitation and must be fixed before tables can be inserted at arbitrary positions. Affects every existing tool's read path.
- **Tables in DXA only:** Per `docs/reference/docx/SKILL.md`, `WidthType.PERCENTAGE` breaks Google Docs — never use it. Always DXA (1440 = 1 inch).
- **Numbering reuse:** Reuse existing `numId` when possible to keep list continuity; only clone an `abstractNum` when a different format is needed.
- **Reference:** `docs/reference/docx/SKILL.md` — table sections (column widths, cell widths, shading, borders, cell margins, gridSpan, vMerge), list sections (BULLET vs DECIMAL formats, shared vs independent reference), tables-as-dividers anti-pattern.

## Scope

### In Scope

1. **xml-utils refactor to preserveOrder:true**
   - Migrate `parseXml` / `buildXml` to ordered mode.
   - Update every existing read helper (`extractParagraphs`, `extractParagraphStyles`, `extractCommentRange`, `parseComments`, etc.) to handle ordered structure.
   - Update every existing write tool to read/write ordered nodes.
   - This is the largest single piece of work in the track.

2. **`word_insert_table`** — create a new table at a specified position:
   - Inputs: filePath, position (after paragraph index | at start | at end | after section), rows, cols, columnWidthsDxa[], optional headerRow, optional named tableStyle.
   - Generates `w:tbl` with `w:tblGrid`, `w:tr`, `w:tc`, optional `w:tblPr` referencing the style.
   - Cell content is empty paragraphs by default.

3. **`word_set_cell_text`** — replace cell content:
   - Inputs: filePath, tableIndex, row, col, text, optional preserveCellPr (default true).
   - Replaces the cell's paragraph content while keeping `w:tcPr`.

4. **`word_modify_table`** — structural changes:
   - Operations: addRow (with index), deleteRow, addColumn (with index), deleteColumn, mergeCells (range), setColumnWidth, setRowHeight.
   - Handles `w:gridSpan` for horizontal merges, `w:vMerge` for vertical merges.

5. **`word_apply_list`** — assign list to paragraph(s):
   - Inputs: filePath, paragraphIndex | paragraphRange, listFormat (`bullet` | `decimal`), optional level (default 0), optional reuseNumId.
   - Adds `w:numPr` with `w:ilvl` and `w:numId`. Reuses an existing list when available; otherwise creates one in `word/numbering.xml`.

6. **`word_set_list_level`** — change indent level of list paragraph:
   - Inputs: filePath, paragraphIndex, level.
   - Updates `w:ilvl` only (preserves the `numId`).

### Out of Scope

- Table-of-tables / list-of-figures fields (covered by `references_navigation` if `\h` switch is implemented)
- Nested tables (defer)
- Charts / SmartArt embedded in tables (defer)

## Audience

User-agnostic. Top use cases by audience:
- **Academic**: bibliography table, comparative methodology table, numbered chapter outlines.
- **Commercial**: pricing tables, schedule tables, comparison matrices.
- **Personal**: any structured data — recipes, budgets, comparison shopping.
- **Educational**: numbered exercises, step-by-step instructions, taxonomy tables.

## Acceptance Criteria

- All 5 tools registered (plus the xml-utils refactor itself, which is plumbing).
- `preserveOrder:true` migration: all existing tests pass after the refactor.
- New tables open in Word and Google Docs without warnings.
- Merged cells render correctly across both editors.
- Lists use existing `numId` when format matches (no proliferation of duplicate abstractNums).
- New fixture: `tests/fixtures/with-tables.docx` covering a doc with several tables in mixed positions for read/modify tests.

## Risks

- **Largest track risk:** the `preserveOrder` migration touches every existing tool. Mitigation: do it as Phase 1 with the existing test suite as the regression net; commit before proceeding to new tool work.
- **Numbering ID conflicts:** reusing `numId` requires careful detection of compatible existing definitions. Mitigation: defensive — when in doubt, clone.
- **Cell merge semantics in Word vs Google Docs:** `vMerge` can render differently. Mitigation: integration test rendering manually in both.

## Dependencies

- Tracks `tracked_changes_v2_20260429`, `style_portability_20260429`, and `references_navigation_20260429` should land first to minimize churn during the `preserveOrder` migration (fewer in-flight changes to update).
