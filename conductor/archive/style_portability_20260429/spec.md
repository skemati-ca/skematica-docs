# Specification: Template & Style Portability

## Overview

Enable copying styles, numbering definitions, theme, and document defaults from a source DOCX file into a target DOCX file. This unlocks template portability — applying a template's "look and feel" to an existing document, sharing personal style libraries, aligning multi-author documents, and standardizing organizational templates.

This is the flagship feature for this phase: it offers immediate, high-leverage value to every user category (academic templates, corporate templates, personal style sets, multi-author chapter alignment).

## Technical Approach

- **Cross-document XML part transfer:** Extend `DocxDocument` with `copyXmlPartFrom(otherDoc, partPath)` for safe cross-document XML reads.
- **ID remapping:** Numbering and style cross-references must be remapped when merging into a target with existing entries (avoid `numId` / `abstractNumId` collisions).
- **Content Types coordination:** When copying parts that don't yet exist in the target (e.g., theme), `[Content_Types].xml` must be updated.
- **Reference:** `docs/reference/docx/SKILL.md` — sections on style overrides (`outlineLevel`, `quickFormat`, `basedOn`, `next`), numbering (`abstractNum`/`num`), and theme.

## Scope

### In Scope

1. **`word_import_styles`** — copy `w:style` entries from source `word/styles.xml` to target. Strategies:
   - `replace`: target's existing styles are entirely replaced.
   - `merge`: source styles added; existing target styles remain.
   - `selective`: copy only the named styles list (with their dependency chain via `basedOn` and `next`).
   - Conflict resolution: skip / overwrite / rename (with suffix).
   - Returns: list of styles imported, list of dependents auto-pulled, conflict report.

2. **`word_apply_template`** — broader: copies any combination of `word/styles.xml`, `word/numbering.xml`, `word/theme/theme1.xml`, `w:sectPr`, and document defaults. Per-part flags:
   - `styles: true|false|"selective"`
   - `numbering: true|false`
   - `theme: true|false`
   - `pageSetup: true|false` (size, orientation, margins from first section)
   - `defaults: true|false` (`w:docDefaults`)
   - Updates `[Content_Types].xml` and relationships when adding parts.

3. **`word_get_styles` extension** — enrich existing tool output with per-style XML excerpt for inspection (so a user can preview a style's properties before importing).

4. **`word_set_default_font`** — set document-default font/size in `word/styles.xml` `w:docDefaults`. Useful for personal style standardization without editing every paragraph.

### Out of Scope

- Header/footer copying (covered by `references_navigation`)
- Custom XML parts (`customXml/`) — defer
- Image / media copying — defer

## Audience

User-agnostic. Top use cases by audience:
- **Academic**: apply a thesis template's styles to an existing draft.
- **Corporate**: standardize a proposal across multiple authors.
- **Personal**: maintain a personal "stylebook" DOCX and apply it to every new document.
- **Educational**: students bringing a teacher-provided template into their own paper.

## Acceptance Criteria

- All 4 tools registered.
- `word_import_styles` resolves `basedOn` / `next` chains automatically when in `selective` mode.
- `word_apply_template` produces a target file that opens in Word without warnings or recovery dialogs.
- `numId` collisions are remapped — pre-existing lists in target remain functional after import.
- Theme import correctly registers the new part in `[Content_Types].xml`.
- Round-trip test: apply template A to doc → apply template A again → no diff (idempotent).
- New fixtures: `tests/fixtures/template-academic.docx`, `tests/fixtures/template-corporate.docx` (rich style sets for testing).

## Risks

- ID remapping is the highest-risk area. Mitigation: thorough test coverage on numbering parts with overlapping IDs.
- Theme import touches `[Content_Types].xml` which must stay consistent. Mitigation: validate the produced file with the bundled `docs/reference/docx/scripts/office/validate.py` script.
- Performance on large `styles.xml`. Mitigation: `selective` mode for large templates.

## Dependencies

- Track `tracked_changes_v2_20260429` (correctness hardening) should land first to ensure writes from this track inherit normalization.
