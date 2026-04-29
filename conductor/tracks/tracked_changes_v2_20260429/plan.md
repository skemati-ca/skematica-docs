# Implementation Plan: Tracked Changes Closure + Correctness Hardening

## Phase 1: OOXML Correctness Hardening (Foundational)

Apply correctness invariants from `docs/reference/docx/SKILL.md` across all writes.

- [x] Task: Write tests for ooxml-normalize module d0ac133
    - [x] Test xml:space="preserve" added to whitespace-bearing w:t
    - [x] Test smart quotes preserved through round-trip
    - [x] Test w:pPr children reordered to (pStyle, numPr, spacing, ind, jc, rPr)
    - [x] Test paraId regeneration for values >= 0x7FFFFFFF
    - [x] Test RSID 8-digit hex validation/coercion
- [x] Task: Implement src/ooxml-normalize.ts d0ac133
    - [x] Whitespace normalization function
    - [x] Smart-quote entity encoding
    - [x] w:pPr child reordering
    - [x] ID regeneration helpers
    - [x] RSID validator
- [x] Task: Wire normalize into DocxDocument.setXmlPart 3e97436
    - [x] Apply on every write to word/document.xml
    - [x] Verify all existing tests still pass (no regressions)
- [~] Task: Add round-trip integration test
    - [ ] Open → save → open → assert structural equivalence

## Phase 2: Cross-Run Tracked Change Insertion

Extend `word_insert_tracked_change` to handle text spanning multiple runs.

- [ ] Task: Write tests for cross-run tracked change
    - [ ] Test text spanning 2 runs with same rPr produces single tracked change
    - [ ] Test text spanning 3+ runs with different rPr preserves each segment's rPr
    - [ ] Test cross-run match at paragraph boundaries
    - [ ] Test backward compatibility: single-run cases still work
- [ ] Task: Extract run-splitting utility
    - [ ] New helper: splitRunAtTextOffset(run, offset)
    - [ ] New helper: collectRunsCoveringRange(paragraph, start, end)
- [ ] Task: Refactor word_insert_tracked_change
    - [ ] Use new utilities for cross-run matches
    - [ ] Wrap each spanned run segment in w:ins/w:del
    - [ ] Preserve rPr per segment
- [ ] Task: Verify existing tests pass

## Phase 3: word_accept_changes Tool

- [ ] Task: Write tests for word_accept_changes
    - [ ] Test accepting w:ins promotes children, removes wrapper
    - [ ] Test accepting w:del drops the content entirely
    - [ ] Test accepting w:pPrChange / w:rPrChange / w:numberingChange
    - [ ] Test filter by author limits scope
    - [ ] Test filter by date range
    - [ ] Test filter by change ID
    - [ ] Test response includes _suggestions
- [ ] Task: Implement word_accept_changes tool
    - [ ] Walk document tree, identify all change elements
    - [ ] Apply filters
    - [ ] Mutate XML per accept rules
    - [ ] Atomic write
    - [ ] Register schema and handler

## Phase 4: word_reject_changes Tool

- [ ] Task: Write tests for word_reject_changes
    - [ ] Test rejecting w:ins drops content entirely
    - [ ] Test rejecting w:del restores w:delText as w:t
    - [ ] Test filter by author / date / id
    - [ ] Test response includes _suggestions
- [ ] Task: Implement word_reject_changes tool
    - [ ] Walk + filter (shared with accept logic)
    - [ ] Apply reject rules
    - [ ] Atomic write
    - [ ] Register schema and handler

## Phase 5: word_set_run_format Tool

- [ ] Task: Write tests for word_set_run_format
    - [ ] Test bold/italic/underline applied to single-run match
    - [ ] Test color (hex) and highlight (named) applied
    - [ ] Test fontSize in half-points
    - [ ] Test cross-run formatting splits runs correctly
    - [ ] Test removing format (passing false explicitly)
    - [ ] Test response includes _suggestions
- [ ] Task: Implement word_set_run_format tool
    - [ ] Reuse run-splitting utility from Phase 2
    - [ ] Build w:rPr with strict child ordering
    - [ ] Atomic write
    - [ ] Register schema and handler

## Phase 6: Conductor - User Manual Verification

- [ ] Task: Conductor - User Manual Verification 'Track tracked_changes_v2' (Protocol in workflow.md)
