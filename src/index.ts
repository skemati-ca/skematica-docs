#!/usr/bin/env node

import { SkematicaDocsServer } from './server.js';
import { getEnabledTools, type ToolName } from './config.js';
import { getAllTools } from './tool-registry.js';

const enabledTools = getEnabledTools();

if (enabledTools.size === 0) {
  console.error('No tools enabled. Check your config or SKEMATICA_DOCS_TOOLS env var.');
  process.exit(1);
}

const server = new SkematicaDocsServer();

// Load and register all enabled tools
const allTools = await getAllTools();

for (const toolName of enabledTools) {
  const toolDef = allTools.get(toolName as ToolName);
  if (toolDef) {
    server.registerTool(toolDef);
  } else {
    console.error(`Warning: Tool "${toolName}" not found in registry.`);
  }
}

// Handle graceful shutdown
process.on('SIGINT', async () => {
  await server.close();
  process.exit(0);
});

process.on('SIGTERM', async () => {
  await server.close();
  process.exit(0);
});

await server.connect();
