# Specification: MCP Workflows + User-Agnostic Positioning

## Overview

Add the MCP **Prompts** capability to the server and refresh README + public documentation to reflect the actual user-agnostic audience (professional + domestic + educational).

MCP Prompts let the server expose composable, user-invokable workflow templates (e.g., `/review-document`). They are different from Tools (which the LLM invokes) and Resources (which the client reads). Prompts are the right primitive for canned workflows users want to trigger directly.

This track lands last in the phase, after enough tools exist to compose meaningful workflows.

## Technical Approach

- **MCP Prompts capability:** Wire `prompts/list` and `prompts/get` handlers in `src/server.ts`. Each prompt is a TypeScript module under `src/prompts/`.
- **Prompt structure:** Each prompt declares `name`, `description`, optional `arguments[]`, and a `getMessages(args)` function that returns the prompt content.
- **No new dependencies.** The `@modelcontextprotocol/sdk` already supports Prompts.
- **Positioning refresh:** README and `docs/*.md` are rewritten to remove lawyer-centric framing. Examples cover all four audience segments.
- **Reference:** `conductor/product.md` (audience definition), `docs/reference/docx/SKILL.md` (capability framing).

## Scope

### In Scope

1. **MCP Prompts capability**:
   - Wire `ListPromptsRequestSchema` and `GetPromptRequestSchema` handlers in `src/server.ts`.
   - Add `src/prompts/registry.ts` mirroring the tool registry pattern.
   - Each prompt as its own file under `src/prompts/`.

2. **Generic prompt templates** (all user-agnostic):
   - **`/review-document`** — structured review pass with comments + structural issues. Args: `filePath`.
   - **`/respond-to-comments`** — process comment threads, draft replies, mark resolved. Args: `filePath`, optional `commentIds[]`.
   - **`/summarize-changes`** — diff between versions or tracked-changes summary. Args: `filePath`, optional `comparePath`.
   - **`/clean-formatting`** — normalize style usage, remove redundant manual format. Args: `filePath`.
   - **`/extract-structure`** — TOC + sections + comments + tables + footnotes overview. Args: `filePath`.
   - **`/apply-template`** — apply a user-defined style set. Args: `filePath`, `templatePath`. Pairs with `style_portability` track.

3. **Documentation refresh**:
   - `README.md`:
     - Replace audience examples that imply only legal use cases.
     - Add use-case examples for academic, business, and personal documents.
     - Update tool count and category breakdown.
     - Add a "MCP Prompts" section explaining the new workflow primitive.
   - `docs/architecture.md`:
     - Note MCP capabilities exposed: Tools (existing), Prompts (new). Resources still deferred.
     - Update component diagram if needed.
   - `docs/mcp-tools.md`:
     - Reorganize by capability category, not user type.
     - Add a Prompts section.
   - `docs/tool-ontology.md`:
     - Add prompt naming pattern.
     - Update narrative to be user-agnostic.

### Out of Scope

- MCP Resources capability — deferred until a "session/active document" model is needed.
- MCP Sampling — deferred.
- Spanish-only documentation — README stays in Spanish per existing convention; technical docs stay in English.

## Audience

User-agnostic. Examples used in the refreshed docs span the four segments:
- Professional: contract review, business proposal, sales report.
- Academic: thesis revision, paper collaboration, syllabus formatting.
- Personal: CV, formal letter, household document.
- Educational: student essay, teacher template application.

## Acceptance Criteria

- Server reports `prompts` capability in MCP handshake.
- `prompts/list` returns all 6 prompts with descriptions.
- `prompts/get` returns valid prompt messages for each.
- README is reviewed for "lawyer-only" language and refreshed.
- All public docs in `docs/` mention all four audience segments.
- Tool count is accurate (current + any new from preceding tracks).
- New integration test verifies prompt list contents and message content per prompt.

## Risks

- Prompt content quality matters: a poorly designed prompt makes the feature feel useless. Mitigation: dogfood each prompt against fixture files.
- Documentation drift between README and `docs/*` if updated separately. Mitigation: review them together in one task.

## Dependencies

- Best landed last so prompts can leverage all the new tools from prior tracks.
- `style_portability` track required for `/apply-template` prompt to be useful.
- `references_navigation` track useful for `/extract-structure` (to surface fields/headers).
