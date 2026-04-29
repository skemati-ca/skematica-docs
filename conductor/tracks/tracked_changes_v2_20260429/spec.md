# Specification: Tracked Changes Closure + Correctness Hardening

## Overview

Close the tracked-changes review loop and apply correctness hardening across the OOXML write paths. Today the MCP can insert tracked changes but cannot accept or reject them, blocking the round-trip review workflow. Additionally, several OOXML invariants flagged in `docs/reference/docx/SKILL.md` are not enforced and can silently corrupt documents on round-trip.

## Technical Approach

- **Runtime:** Existing Node.js v20+ / TypeScript stack.
- **OOXML normalization:** New shared module `src/ooxml-normalize.ts` applied in `DocxDocument.setXmlPart` so every write benefits.
- **Run splitting:** Reusable utility (extracted from `word-insert-tracked-change.ts`) that handles OOXML run fragmentation for tracked changes and direct formatting alike.
- **Reference:** `docs/reference/docx/SKILL.md` — sections on tracked-change patterns, invariants, and `xml:space="preserve"`.

## Scope

### In Scope

1. **OOXML correctness pass** (foundational — applied to all writes):
   - Auto-add `xml:space="preserve"` to every `w:t` containing leading/trailing whitespace.
   - Smart-quote round-trip preservation (entity encoding `&#x2018;` `&#x2019;` `&#x201C;` `&#x201D;`).
   - Enforce `w:pPr` child element ordering (`pStyle`, `numPr`, `spacing`, `ind`, `jc`, `rPr` last).
   - Regenerate `paraId` / `durableId` values when `>= 0x7FFFFFFF`.
   - Validate RSIDs are 8-digit hex; coerce when invalid.

2. **`word_insert_tracked_change` v2** — cross-run support:
   - Replace whole `w:r` blocks per SKILL.md (never splice tags inside a run).
   - Split surrounding runs as needed to wrap matched runs in `w:ins` / `w:del`.
   - Preserve original `w:rPr` on both `w:del` and `w:ins` sides.

3. **`word_accept_changes`** — apply tracked changes:
   - Unwrap `w:ins` (promote children to siblings).
   - Drop `w:del` content entirely.
   - Convert `w:delText` to `w:t` only when restoring (not when accepting).
   - Handle `w:pPrChange`, `w:rPrChange`, `w:numberingChange`.
   - Filters: by author, by date range, by change ID, by section.

4. **`word_reject_changes`** — discard tracked changes:
   - Drop `w:ins` content entirely (the proposed text disappears).
   - Convert `w:delText` back to `w:t`.
   - Same filter set as accept.

5. **`word_set_run_format`** — direct run formatting without a named style:
   - Properties: `bold`, `italic`, `underline`, `strikethrough`, `color` (hex), `highlight` (Word color name), `fontSize` (half-points).
   - Splits runs at the boundaries of the matched text.
   - Writes `w:rPr` with strict child ordering.
   - Operates on a search text or paragraph + character range.

### Out of Scope (Other Tracks)

- Style import/copy (`style_portability` track)
- Hyperlinks, bookmarks, fields (`references_navigation` track)
- Tables (`tables_lists` track)

## Audience

User-agnostic. The tracked-changes loop matters equally to a lawyer reviewing a contract, a manager approving a report, an academic responding to thesis revisions, or a domestic user finalizing a personal letter.

## Acceptance Criteria

- All 5 tools registered in `src/tool-registry.ts` and `src/config.ts`.
- Accept/reject tools restore the document to a clean (no-tracked-changes) state without corrupting structure.
- `word_insert_tracked_change` succeeds on cross-run matches that previously failed.
- `word_set_run_format` applies formatting on text spanning multiple runs.
- Round-trip test: insert tracked change → accept → file is byte-for-byte equivalent in semantic structure to one where the change was applied directly.
- All existing tests still pass after `ooxml-normalize` is applied to writes.
- New integration tests cover each tool with `with-comments.docx`, `cv.docx`, and a new fixture with pre-existing tracked changes.

## Risks

- Normalization pass could subtly change byte output of existing tools' writes. Mitigation: extensive integration tests + diff against pre-change output.
- Cross-run editing logic is the most error-prone area in OOXML. Mitigation: TDD with edge cases (whitespace, multi-paragraph, tracked change inside another tracked change).
