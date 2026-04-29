# Implementation Plan: Template & Style Portability

## Phase 1: Cross-Document Infrastructure

Foundation: safe reading from a second DOCX, ID remapping helpers, content-types coordination.

- [ ] Task: Write tests for cross-document loading
    - [ ] Test loading two DocxDocument instances simultaneously
    - [ ] Test reading source xml part without affecting target
- [ ] Task: Extend DocxDocument
    - [ ] Add copyXmlPartFrom(sourceDoc, partPath) method
    - [ ] Add hasPart(path) method
    - [ ] Add registerPartInContentTypes(partPath, contentType) helper
- [ ] Task: Write tests for ID remapping helpers
    - [ ] Test numId collision detection
    - [ ] Test abstractNumId remap preserves cross-references
    - [ ] Test styleId conflict resolution (skip/overwrite/rename)
- [ ] Task: Implement src/style-merge-utils.ts
    - [ ] Style dependency resolver (follow basedOn / next)
    - [ ] numId / abstractNumId remap
    - [ ] styleId conflict handler
- [ ] Task: Add fixture: tests/fixtures/template-academic.docx
- [ ] Task: Add fixture: tests/fixtures/template-corporate.docx

## Phase 2: word_import_styles Tool

- [ ] Task: Write tests for word_import_styles
    - [ ] Test replace strategy: target's styles fully replaced
    - [ ] Test merge strategy: source added, existing kept
    - [ ] Test selective strategy: only listed styles + dependents
    - [ ] Test conflict resolution: skip / overwrite / rename
    - [ ] Test invalid source path returns error
    - [ ] Test response includes _suggestions (e.g., word_apply_style)
- [ ] Task: Implement word_import_styles tool
    - [ ] Schema with sourcePath, strategy, styles[], onConflict
    - [ ] Load source via DocxDocument
    - [ ] Apply chosen strategy via style-merge-utils
    - [ ] Atomic write to target
    - [ ] Register schema and handler

## Phase 3: word_apply_template Tool

- [ ] Task: Write tests for word_apply_template
    - [ ] Test applying full template (all parts)
    - [ ] Test applying only styles + numbering
    - [ ] Test applying only pageSetup from first section
    - [ ] Test theme import registers part in [Content_Types].xml
    - [ ] Test idempotency: apply twice = no diff
    - [ ] Test produced file opens cleanly (validate via XSDs if available)
    - [ ] Test response includes _suggestions
- [ ] Task: Implement word_apply_template tool
    - [ ] Schema with sourcePath + per-part flags
    - [ ] Orchestrate copy of styles, numbering, theme, pageSetup, defaults
    - [ ] Atomic write
    - [ ] Register schema and handler

## Phase 4: word_get_styles Extension

- [ ] Task: Write tests for extended word_get_styles
    - [ ] Test new includeXml flag returns per-style XML excerpt
    - [ ] Test default behavior unchanged
    - [ ] Test backward compatibility for existing callers
- [ ] Task: Extend word_get_styles tool
    - [ ] Add optional includeXml: boolean param
    - [ ] Surface w:style XML when requested
    - [ ] Update schema

## Phase 5: word_set_default_font Tool

- [ ] Task: Write tests for word_set_default_font
    - [ ] Test setting font name updates w:docDefaults
    - [ ] Test setting size in half-points (24 = 12pt)
    - [ ] Test setting both font and size
    - [ ] Test invalid font returns clear error if validation enabled
    - [ ] Test response includes _suggestions
- [ ] Task: Implement word_set_default_font tool
    - [ ] Schema: filePath, fontName?, sizeHalfPoints?
    - [ ] Modify w:docDefaults > w:rPrDefault > w:rPr
    - [ ] Atomic write
    - [ ] Register schema and handler

## Phase 6: Conductor - User Manual Verification

- [ ] Task: Conductor - User Manual Verification 'Track style_portability' (Protocol in workflow.md)
