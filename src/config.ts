import { existsSync, readFileSync } from 'node:fs';
import { join } from 'node:path';
import { homedir } from 'node:os';

const ALL_TOOLS = [
  'word_get_document_info',
  'word_get_content',
  'word_get_sections',
  'word_get_section_content',
  'word_find_text',
  'word_search_replace',
  'word_list_comments',
  'word_create_comment',
  'word_reply_to_comment',
  'word_resolve_comment',
  'word_get_page_setup',
  'word_set_page_size',
  'word_set_orientation',
  'word_set_margins',
  'word_compare_versions',
  'word_get_styles',
  'word_apply_style',
  'word_get_footnotes',
  'word_insert_tracked_change',
] as const;

export type ToolName = (typeof ALL_TOOLS)[number];

export interface ToolConfig {
  tools: Partial<Record<ToolName, boolean>>;
}

function loadFromFile(): ToolConfig | null {
  const candidates = [
    join(process.cwd(), 'skematica-docs.json'),
    join(homedir(), '.config', 'skematica-docs', 'config.json'),
  ];

  for (const path of candidates) {
    if (existsSync(path)) {
      try {
        const raw = readFileSync(path, 'utf-8');
        const parsed = JSON.parse(raw) as ToolConfig;
        if (parsed?.tools && typeof parsed.tools === 'object') {
          console.error(`Loaded tool config from ${path}`);
          return parsed;
        }
      } catch {
        console.error(`Failed to parse config file: ${path}`);
      }
    }
  }

  return null;
}

function loadFromEnv(): Set<ToolName> | null {
  const envValue = process.env.SKEMATICA_DOCS_TOOLS;
  if (!envValue || envValue.trim() === '') {
    return null;
  }

  const enabled = new Set<ToolName>();
  const parts = envValue.split(',').map((s) => s.trim());

  for (const part of parts) {
    if (ALL_TOOLS.includes(part as ToolName)) {
      enabled.add(part as ToolName);
    }
  }

  console.error(`Loaded ${enabled.size} tools from SKEMATICA_DOCS_TOOLS env var`);
  return enabled;
}

export function getEnabledTools(): Set<ToolName> {
  // Environment variable takes precedence
  const envTools = loadFromEnv();
  if (envTools) {
    return envTools;
  }

  // Then config file
  const fileConfig = loadFromFile();
  if (fileConfig) {
    const enabled = new Set<ToolName>();
    for (const [name, value] of Object.entries(fileConfig.tools)) {
      if (value !== false && ALL_TOOLS.includes(name as ToolName)) {
        enabled.add(name as ToolName);
      }
    }
    return enabled;
  }

  // Default: all tools enabled
  return new Set<ToolName>(ALL_TOOLS);
}
