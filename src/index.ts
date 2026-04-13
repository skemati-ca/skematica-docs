#!/usr/bin/env node

import { SkematicaDocsServer } from './server.js';
import { getEnabledTools, type ToolName } from './config.js';

const enabledTools = getEnabledTools();

if (enabledTools.size === 0) {
  console.error('No tools enabled. Check your config or SKEMATICA_DOCS_TOOLS env var.');
  process.exit(1);
}

const server = new SkematicaDocsServer();

// Register enabled tools
const toolRegistry: Partial<Record<ToolName, { description: string; inputSchema: Record<string, unknown>; handler: (args: Record<string, unknown>) => Promise<Record<string, unknown>> }>> = {
  // Tools will be registered here as they are implemented
};

for (const toolName of enabledTools) {
  const toolDef = toolRegistry[toolName];
  if (toolDef) {
    server.registerTool({
      name: toolName,
      description: toolDef.description,
      inputSchema: toolDef.inputSchema,
      handler: toolDef.handler,
    });
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
