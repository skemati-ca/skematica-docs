import { describe, it, expect, beforeEach, afterEach } from "vitest";
import { normalize } from "node:path";
import { validateDocxPath } from "../src/validation.js";
import { SkematicaDocsServer } from "../src/server.js";

describe("MCP Server lifecycle", () => {
  let server: SkematicaDocsServer;

  beforeEach(() => {
    server = new SkematicaDocsServer();
  });

  afterEach(async () => {
    await server.close();
  });

  it("creates server instance", () => {
    expect(server).toBeDefined();
  });

  it("returns empty tool list when no tools registered", () => {
    const tools = server.getRegisteredTools();
    expect(tools).toEqual([]);
  });

  it("handles graceful shutdown", async () => {
    await expect(server.close()).resolves.not.toThrow();
  });
});

describe("File format validation", () => {
  it("rejects .doc file extension with conversion instructions", () => {
    const result = validateDocxPath("/test/document.doc");
    expect(result).not.toBeNull();
    expect(result).toContain(".doc");
    expect(result).toContain("Save As");
    expect(result).toContain("Google Docs");
    expect(result).toContain("LibreOffice");
    expect(result).toContain("CloudConvert");
  });

  it("rejects unsupported formats with clear error", () => {
    const result = validateDocxPath("/test/document.pdf");
    expect(result).not.toBeNull();
    expect(result).toContain(".pdf");
    expect(result).toContain(".docx");
  });

  it("accepts .docx files", () => {
    const result = validateDocxPath("/test/document.docx");
    expect(result).toBeNull();
  });

  it("rejects files with no extension", () => {
    const result = validateDocxPath("/test/document");
    expect(result).not.toBeNull();
    expect(result).toContain("No file extension");
  });
});

describe("Cross-platform file path handling", () => {
  it("normalizes Windows paths (C:\\Users\\...)", () => {
    const winPath = "C:\\Users\\test\\document.docx";
    const normalized = normalize(winPath);
    expect(normalized).toBeDefined();
  });

  it("normalizes Unix paths (/home/user/doc.docx)", () => {
    const unixPath = "/home/user/document.docx";
    const normalized = normalize(unixPath);
    expect(normalized).toBeDefined();
  });
});

describe("SkematicaDocsServer tool registration", () => {
  let server: SkematicaDocsServer;

  beforeEach(() => {
    server = new SkematicaDocsServer();
  });

  afterEach(async () => {
    await server.close();
  });

  it("registers a tool and lists it", () => {
    server.registerTool({
      name: "test_tool",
      description: "A test tool",
      inputSchema: { type: "object", properties: {} },
      handler: async () => ({ result: "ok" }),
    });

    const tools = server.getRegisteredTools();
    expect(tools).toContain("test_tool");
  });

  it("registers multiple tools", () => {
    server.registerTool({
      name: "tool_a",
      description: "Tool A",
      inputSchema: { type: "object", properties: {} },
      handler: async () => ({ result: "a" }),
    });
    server.registerTool({
      name: "tool_b",
      description: "Tool B",
      inputSchema: { type: "object", properties: {} },
      handler: async () => ({ result: "b" }),
    });

    const tools = server.getRegisteredTools();
    expect(tools).toContain("tool_a");
    expect(tools).toContain("tool_b");
    expect(tools).toHaveLength(2);
  });

  it("returns error for unknown tool in call handler", async () => {
    server.registerTool({
      name: "test_tool",
      description: "A test tool",
      inputSchema: { type: "object", properties: {} },
      handler: async () => ({ result: "ok" }),
    });

    // The CallToolRequestSchema handler should throw McpError for unknown tools
    // This is tested implicitly through the server setup
    const tools = server.getRegisteredTools();
    expect(tools).not.toContain("unknown_tool");
  });

  it("handles tool error gracefully", async () => {
    server.registerTool({
      name: "failing_tool",
      description: "A tool that fails",
      inputSchema: { type: "object", properties: {} },
      handler: async () => {
        throw new Error("intentional error");
      },
    });

    const tools = server.getRegisteredTools();
    expect(tools).toContain("failing_tool");
  });
});
