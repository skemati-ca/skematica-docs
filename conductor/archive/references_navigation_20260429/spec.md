# Specification: References & Navigation

## Overview

Add a coherent set of cross-document navigation features: hyperlinks, bookmarks, cross-references, fields (TOC, PAGE, NUMPAGES, DATE, REF), and headers/footers. These are universal needs across academic theses, business proposals, technical reports, and personal documents.

A new shared utility (`src/rels-utils.ts`) handles `word/_rels/document.xml.rels` mutations and is reusable for future image-insertion tooling.

## Technical Approach

- **Relationships first:** Hyperlinks and header/footer references both require entries in `.rels` files. Build the helper once.
- **Field codes:** Word fields are either `w:fldSimple` (one element) or a triple of `w:fldChar` (begin/separate/end) with `w:instrText`. Use simple-form when possible.
- **Header/footer parts:** New XML parts under `word/header*.xml` and `word/footer*.xml`, referenced from section properties via `w:headerReference` / `w:footerReference`.
- **Reference:** `docs/reference/docx/SKILL.md` — sections on hyperlinks (`ExternalHyperlink`, `InternalHyperlink`), bookmarks, fields, page-number footers, and tab stops for footer alignment.

## Scope

### In Scope

1. **`src/rels-utils.ts`** — shared utility:
   - `addRelationship(doc, type, target, targetMode?) → relId`
   - `removeRelationship(doc, relId)`
   - `getRelationship(doc, relId)`
   - Relationship file: `word/_rels/document.xml.rels`. Header/footer rels handled separately.

2. **`word_insert_hyperlink`** — wrap text in `w:hyperlink`:
   - Inputs: filePath, searchText, url, optional matchIndex.
   - Adds `.rels` entry for external URLs (with `targetMode="External"`).
   - Internal anchors (to bookmarks) use `w:hyperlink/@w:anchor` (no rel).

3. **`word_remove_hyperlink`** — find `w:hyperlink` containing matching text, unwrap to plain run, drop the relationship.

4. **`word_create_bookmark`** — wrap text in `w:bookmarkStart` + `w:bookmarkEnd` with auto-generated id and unique name.

5. **`word_remove_bookmark`** — by name.

6. **`word_insert_cross_reference`** — `w:fldSimple` with `REF <bookmarkName>` or `PAGEREF <bookmarkName>` at the search location.

7. **`word_insert_field`** — generic field insertion:
   - Supported codes: `TOC`, `PAGE`, `NUMPAGES`, `DATE`, `AUTHOR`, `FILENAME`.
   - Sets `w:fldSimple` with proper instruction text.
   - For `TOC`, includes the `\o "1-N"` switch and `\h \z` for hyperlinks.
   - Sets dirty-flag (`w:dirty="true"`) so client updates on next render.

8. **`word_read_headers_footers`** — parse all `header*.xml` / `footer*.xml` parts, return text + section reference info (which sections use which header/footer types: `default`, `first`, `even`).

9. **`word_set_header_footer_text`** — write text into a header or footer:
   - Inputs: filePath, target (`header` | `footer`), type (`default` | `first` | `even`), sectionIndex, text.
   - Creates the part + relationship + section reference if missing.
   - Text-only (paragraph style optional). Page numbers via `word_insert_field` after this lands.

### Out of Scope

- Header/footer images (defer to later track with image support)
- Custom XML / SDT
- Equations, shapes

## Audience

User-agnostic. Examples per audience:
- **Academic**: TOC + cross-refs to figures/tables/sections; page numbers in footer.
- **Business**: external hyperlinks in proposals; document title + date in header.
- **Personal**: hyperlinks in CV; page count in long-form documents.
- **Technical**: bookmarks for clause references; PAGEREF for printable manuals.

## Acceptance Criteria

- All 8 tools registered.
- `.rels` mutations are reversible: insert + remove returns to original byte structure.
- Cross-references resolve when document is opened in Word (manual verification).
- TOC field with no existing entries is inserted as a "click to update" placeholder; when the user updates fields in Word it populates correctly.
- Headers/footers respect per-section reference types (`default`, `first`, `even`) and don't overwrite siblings.
- Round-trip on existing fixtures passes — new field/bookmark/hyperlink elements don't disturb unrelated content.

## Risks

- `.rels` ID collisions across different rels files (document.xml.rels vs header1.xml.rels). Mitigation: scope IDs per file in `rels-utils`.
- Field instruction syntax is brittle. Mitigation: tested template strings per supported code.
- Header/footer parts must coordinate with `[Content_Types].xml`. Mitigation: leverage helper from `style_portability` track if available; otherwise duplicate.

## Dependencies

- Track `tracked_changes_v2_20260429` (correctness hardening preferred first).
- Track `style_portability_20260429` is helpful for `[Content_Types].xml` helper but not blocking.
