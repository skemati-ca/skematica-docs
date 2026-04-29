# Implementation Plan: Tables & Lists

## Phase 1: xml-utils preserveOrder Migration (Largest Risk)

Migrate fast-xml-parser config to preserveOrder:true. Required for safe table edits.

- [ ] Task: Spike + design doc on preserveOrder structure
    - [ ] Document new node shape (array of single-key records)
    - [ ] Identify every existing helper that reads body content
- [ ] Task: Update parseXml / buildXml in src/xml-utils.ts
    - [ ] preserveOrder: true
    - [ ] Update related options
- [ ] Task: Migrate read helpers
    - [ ] collectBodyParagraphs adapted to ordered structure
    - [ ] extractParagraphs / extractParagraphStyles / extractParagraphText
    - [ ] extractCommentRange / findAllTextInNode
    - [ ] parseComments / parseCommentsExtended
    - [ ] extractPageSetup / extractParagraphNumPr
- [ ] Task: Migrate write tools
    - [ ] word_search_replace
    - [ ] word_create_comment / word_reply_to_comment / word_resolve_comment
    - [ ] word_insert_tracked_change
    - [ ] word_apply_style
    - [ ] word_set_page_size / word_set_orientation / word_set_margins
- [ ] Task: Run full test suite — no regressions
- [ ] Task: Commit isolated migration commit

## Phase 2: Table Reading Infrastructure

- [ ] Task: Add fixture: tests/fixtures/with-tables.docx
- [ ] Task: Write tests for table accessors
    - [ ] Test getTables returns tables in document order
    - [ ] Test each table reports rows / cols / dimensions
    - [ ] Test mixed body/table interleaving preserved post-migration
- [ ] Task: Add DocxDocument.getTables() method
    - [ ] Walk ordered body, collect w:tbl entries
    - [ ] Return list with index, rows, cols, cell widths

## Phase 3: word_insert_table Tool

- [ ] Task: Write tests for word_insert_table
    - [ ] Test create simple 2x3 table at end
    - [ ] Test create with named tableStyle
    - [ ] Test create at specific paragraph index
    - [ ] Test custom column widths in DXA
    - [ ] Test header row marked correctly
    - [ ] Test response includes _suggestions
- [ ] Task: Implement word_insert_table tool
    - [ ] Build w:tbl with w:tblGrid, w:tr, w:tc
    - [ ] Position at requested location
    - [ ] Atomic write
    - [ ] Register schema and handler

## Phase 4: word_set_cell_text Tool

- [ ] Task: Write tests for word_set_cell_text
    - [ ] Test replace text in cell preserves w:tcPr
    - [ ] Test out-of-range row/col returns error
    - [ ] Test response includes _suggestions
- [ ] Task: Implement word_set_cell_text tool

## Phase 5: word_modify_table Tool

- [ ] Task: Write tests for word_modify_table
    - [ ] Test addRow at index
    - [ ] Test deleteRow
    - [ ] Test addColumn at index (updates w:tblGrid + every row)
    - [ ] Test deleteColumn
    - [ ] Test mergeCells horizontal sets w:gridSpan
    - [ ] Test mergeCells vertical sets w:vMerge
    - [ ] Test setColumnWidth updates grid
    - [ ] Test response includes _suggestions
- [ ] Task: Implement word_modify_table tool
    - [ ] Operation switch dispatch
    - [ ] Atomic write
    - [ ] Register schema and handler

## Phase 6: Lists - word_apply_list and word_set_list_level

- [ ] Task: Write tests for word_apply_list
    - [ ] Test bullet list applied to single paragraph
    - [ ] Test decimal list applied to range
    - [ ] Test reuse of existing compatible numId
    - [ ] Test creation of new abstractNum when needed
    - [ ] Test response includes _suggestions
- [ ] Task: Implement word_apply_list tool
- [ ] Task: Write tests for word_set_list_level
    - [ ] Test level change preserves numId
    - [ ] Test invalid paragraph (not in list) returns clear error
    - [ ] Test response includes _suggestions
- [ ] Task: Implement word_set_list_level tool

## Phase 7: Conductor - User Manual Verification

- [ ] Task: Conductor - User Manual Verification 'Track tables_lists' (Protocol in workflow.md)
