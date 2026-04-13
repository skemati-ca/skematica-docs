import { describe, it, expect, beforeEach, afterEach } from "vitest";
import { mkdirSync, writeFileSync, rmSync, existsSync } from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { getEnabledTools, type ToolName } from "../src/config.js";

describe("Tool configuration", () => {
  let originalEnv: string | undefined;
  let originalCwd: string;

  beforeEach(() => {
    originalEnv = process.env.SKEMATICA_DOCS_TOOLS;
    originalCwd = process.cwd();
    delete process.env.SKEMATICA_DOCS_TOOLS;
  });

  afterEach(() => {
    if (originalEnv !== undefined) {
      process.env.SKEMATICA_DOCS_TOOLS = originalEnv;
    } else {
      delete process.env.SKEMATICA_DOCS_TOOLS;
    }
    process.chdir(originalCwd);
  });

  it("returns all tools when no config is present", () => {
    // With no config file and no env var, all tools should be enabled
    const enabled = getEnabledTools();
    expect(enabled.size).toBeGreaterThan(0);
  });

  it("parses environment variable for tool list", () => {
    process.env.SKEMATICA_DOCS_TOOLS = "word_get_content,word_find_text";
    const enabled = getEnabledTools();
    expect(enabled.has("word_get_content")).toBe(true);
    expect(enabled.has("word_find_text")).toBe(true);
    expect(enabled.has("word_search_replace")).toBe(false);
  });

  it("ignores invalid tool names in env var", () => {
    process.env.SKEMATICA_DOCS_TOOLS = "word_get_content,invalid_tool,word_find_text";
    const enabled = getEnabledTools();
    expect(enabled.has("word_get_content")).toBe(true);
    expect(enabled.has("word_find_text")).toBe(true);
    expect(enabled.size).toBe(2);
  });

  it("handles empty env var gracefully", () => {
    process.env.SKEMATICA_DOCS_TOOLS = "";
    const enabled = getEnabledTools();
    expect(enabled.size).toBeGreaterThan(0);
  });
});

describe("Config file loading", () => {
  let testDir: string;
  let originalCwd: string;

  beforeEach(() => {
    delete process.env.SKEMATICA_DOCS_TOOLS;
    testDir = join(tmpdir(), `skematica-config-test-${Date.now()}`);
    mkdirSync(testDir, { recursive: true });
    originalCwd = process.cwd();
    process.chdir(testDir);
  });

  afterEach(() => {
    process.chdir(originalCwd);
    if (existsSync(testDir)) {
      rmSync(testDir, { recursive: true, force: true });
    }
  });

  it("loads tools from skematica-docs.json", () => {
    const config = {
      tools: {
        word_get_content: true,
        word_find_text: true,
        word_search_replace: false,
      },
    };
    writeFileSync(
      join(testDir, "skematica-docs.json"),
      JSON.stringify(config)
    );

    const enabled = getEnabledTools();
    expect(enabled.has("word_get_content")).toBe(true);
    expect(enabled.has("word_find_text")).toBe(true);
    expect(enabled.has("word_search_replace")).toBe(false);
  });

  it("env var takes precedence over config file", () => {
    const config = {
      tools: {
        word_get_content: true,
        word_find_text: true,
      },
    };
    writeFileSync(
      join(testDir, "skematica-docs.json"),
      JSON.stringify(config)
    );

    process.env.SKEMATICA_DOCS_TOOLS = "word_search_replace";
    const enabled = getEnabledTools();
    expect(enabled.has("word_search_replace")).toBe(true);
    expect(enabled.has("word_get_content")).toBe(false);
  });
});
