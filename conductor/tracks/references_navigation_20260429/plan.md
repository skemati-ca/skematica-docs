# Implementation Plan: References & Navigation

## Phase 1: Relationships Utility

Foundational shared module — used by hyperlinks, headers/footers, future image tooling.

- [ ] Task: Write tests for rels-utils
    - [ ] Test addRelationship returns unique ID
    - [ ] Test addRelationship with targetMode="External"
    - [ ] Test removeRelationship cleans up entry
    - [ ] Test getRelationship returns target
    - [ ] Test multiple .rels files are scoped independently
- [ ] Task: Implement src/rels-utils.ts
    - [ ] addRelationship(doc, type, target, targetMode?)
    - [ ] removeRelationship(doc, relId)
    - [ ] getRelationship(doc, relId)
    - [ ] Per-file ID counter

## Phase 2: Hyperlinks

- [ ] Task: Write tests for word_insert_hyperlink
    - [ ] Test external URL adds .rels entry and wraps run
    - [ ] Test internal anchor uses w:anchor without .rels
    - [ ] Test searchText not found returns clear error
    - [ ] Test cross-run hyperlink wrapping
    - [ ] Test response includes _suggestions
- [ ] Task: Implement word_insert_hyperlink tool
- [ ] Task: Write tests for word_remove_hyperlink
    - [ ] Test removal unwraps w:hyperlink to plain run
    - [ ] Test associated .rels entry is dropped
    - [ ] Test response includes _suggestions
- [ ] Task: Implement word_remove_hyperlink tool

## Phase 3: Bookmarks

- [ ] Task: Write tests for word_create_bookmark
    - [ ] Test wraps text in bookmarkStart + bookmarkEnd
    - [ ] Test auto-generated unique id
    - [ ] Test name uniqueness validation
    - [ ] Test response includes _suggestions
- [ ] Task: Implement word_create_bookmark tool
- [ ] Task: Write tests for word_remove_bookmark
    - [ ] Test removal by name
    - [ ] Test removal of unknown name returns error
- [ ] Task: Implement word_remove_bookmark tool

## Phase 4: Cross-References and Fields

- [ ] Task: Write tests for word_insert_cross_reference
    - [ ] Test REF <bookmark> inserts w:fldSimple
    - [ ] Test PAGEREF variant
    - [ ] Test reference to non-existent bookmark returns error
    - [ ] Test response includes _suggestions
- [ ] Task: Implement word_insert_cross_reference tool
- [ ] Task: Write tests for word_insert_field
    - [ ] Test TOC field with default depth and switches
    - [ ] Test PAGE / NUMPAGES insertion
    - [ ] Test DATE / AUTHOR / FILENAME
    - [ ] Test dirty-flag set
    - [ ] Test response includes _suggestions
- [ ] Task: Implement word_insert_field tool

## Phase 5: Headers and Footers

- [ ] Task: Write tests for word_read_headers_footers
    - [ ] Test reads default / first / even types
    - [ ] Test returns per-section reference map
    - [ ] Test document with no headers returns empty
    - [ ] Test response includes _suggestions
- [ ] Task: Implement word_read_headers_footers tool
    - [ ] Parse word/header*.xml and word/footer*.xml parts
    - [ ] Build section -> {default,first,even} reference map
- [ ] Task: Write tests for word_set_header_footer_text
    - [ ] Test creating new header/footer (was missing)
    - [ ] Test updating existing header/footer
    - [ ] Test type=first / type=even
    - [ ] Test [Content_Types].xml registration when adding new part
    - [ ] Test section reference (w:headerReference) added
    - [ ] Test response includes _suggestions
- [ ] Task: Implement word_set_header_footer_text tool
    - [ ] Create part + .rels + section reference if needed
    - [ ] Write text-only paragraph

## Phase 6: Conductor - User Manual Verification

- [ ] Task: Conductor - User Manual Verification 'Track references_navigation' (Protocol in workflow.md)
