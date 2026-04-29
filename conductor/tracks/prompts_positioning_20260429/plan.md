# Implementation Plan: MCP Workflows + User-Agnostic Positioning

## Phase 1: Prompts Infrastructure

Foundation: expose MCP Prompts alongside existing Tools without changing tool behavior.

- [ ] Task: Write tests for MCP prompts capability
    - [ ] Test server handshake reports prompts capability
    - [ ] Test prompts/list returns prompt metadata
    - [ ] Test prompts/get validates required arguments
    - [ ] Test unknown prompt name returns a clear MCP error
- [ ] Task: Implement prompt registry
    - [ ] Add src/prompts/registry.ts mirroring the tool registry pattern
    - [ ] Define shared prompt metadata and getMessages contract
    - [ ] Export listPrompts and getPrompt helpers
- [ ] Task: Wire prompts into src/server.ts
    - [ ] Register ListPromptsRequestSchema handler
    - [ ] Register GetPromptRequestSchema handler
    - [ ] Advertise prompts capability in server options
    - [ ] Preserve current tools behavior

## Phase 2: User-Agnostic Workflow Prompts

- [ ] Task: Write tests for workflow prompt contents
    - [ ] Test /review-document includes read + comment + structure review workflow
    - [ ] Test /respond-to-comments includes comment thread reply + resolve workflow
    - [ ] Test /summarize-changes handles tracked changes and optional comparePath
    - [ ] Test /clean-formatting uses style inspection and style application
    - [ ] Test /extract-structure includes sections, comments, tables, and footnotes
    - [ ] Test /apply-template includes filePath and templatePath
- [ ] Task: Implement workflow prompt modules
    - [ ] Add src/prompts/review-document.ts
    - [ ] Add src/prompts/respond-to-comments.ts
    - [ ] Add src/prompts/summarize-changes.ts
    - [ ] Add src/prompts/clean-formatting.ts
    - [ ] Add src/prompts/extract-structure.ts
    - [ ] Add src/prompts/apply-template.ts
    - [ ] Register all prompts in registry

## Phase 3: README Positioning Refresh

- [ ] Task: Audit README for lawyer-only framing
    - [ ] Identify examples and claims that narrow the audience unnecessarily
    - [ ] Identify tool count and category references that need updates
- [ ] Task: Refresh README
    - [ ] Reframe audience around long collaborative documents
    - [ ] Add professional, academic, personal, and educational examples
    - [ ] Add MCP Prompts section
    - [ ] Update tool count and capability categories
    - [ ] Keep README language consistent with existing Spanish convention

## Phase 4: Public Docs Refresh

- [ ] Task: Update docs/architecture.md
    - [ ] Document exposed MCP capabilities: Tools and Prompts
    - [ ] Keep Resources explicitly deferred
    - [ ] Update component diagram or narrative if needed
- [ ] Task: Update docs/mcp-tools.md
    - [ ] Reorganize by capability category
    - [ ] Add Prompts section with all six workflows
    - [ ] Ensure examples cover professional, academic, personal, and educational use
- [ ] Task: Update docs/tool-ontology.md
    - [ ] Add prompt naming pattern
    - [ ] Clarify difference between tools and user-invoked prompts
    - [ ] Remove domain-specific assumptions from taxonomy narrative

## Phase 5: Integration Verification

- [ ] Task: Add integration tests for prompt workflows
    - [ ] Verify prompt list contents exactly matches six expected prompts
    - [ ] Verify each prompt returns valid MCP prompt messages
    - [ ] Verify prompt arguments are reflected in generated messages
    - [ ] Verify prompts mention available related tools without requiring unavailable resources
- [ ] Task: Run full test suite
    - [ ] npm test
    - [ ] npm run build
    - [ ] Confirm no regressions in existing tools

## Phase 6: Conductor - User Manual Verification

- [ ] Task: Conductor - User Manual Verification 'Track prompts_positioning' (Protocol in workflow.md)
