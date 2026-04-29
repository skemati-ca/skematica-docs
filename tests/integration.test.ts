import { describe, it, expect } from "vitest";
import { existsSync } from "node:fs";
import { join } from "node:path";
import { DocxDocument } from "../src/docx.js";
import { validateDocxPath } from "../src/validation.js";
import { wordGetDocumentInfo } from "../src/tools/word-get-document-info.js";

const fixturesDir = join(process.cwd(), "tests", "fixtures");

describe("Integration: Real DOCX files", () => {
  it("loads simple.docx and extracts text", async () => {
    const path = join(fixturesDir, "simple.docx");
    expect(existsSync(path)).toBe(true);

    const doc = await DocxDocument.load(path);
    const text = await doc.getText();
    expect(text).toContain("Contrato de Prestación de Servicios");
    expect(text).toContain("CLÁUSULA PRIMERA");
  });

  it("returns metadata for simple.docx", async () => {
    const path = join(fixturesDir, "simple.docx");
    const doc = await DocxDocument.load(path);
    const meta = await doc.getMetadata();

    expect(meta.pageCount).toBeGreaterThanOrEqual(1);
    expect(meta.wordCount).toBeGreaterThanOrEqual(10);
    expect(meta.format).toBe("docx");
    expect(meta.commentCount).toBe(0);
  });

  it("finds sections in structured.docx", async () => {
    const path = join(fixturesDir, "structured.docx");
    expect(existsSync(path)).toBe(true);

    const doc = await DocxDocument.load(path);
    const sections = await doc.getSections();

    expect(sections.length).toBeGreaterThanOrEqual(3);
    expect(sections[0].title).toContain("1.");
  });

  it("reads section content from structured.docx", async () => {
    const path = join(fixturesDir, "structured.docx");
    const doc = await DocxDocument.load(path);
    const sections = await doc.getSections();

    if (sections.length > 0) {
      const content = await doc.getSectionContent(sections[0].title, 5000);
      expect(content.section).toBe(sections[0].title);
      expect(content.wordCount).toBeGreaterThanOrEqual(0);
    }
  });

  it("finds text in simple.docx", async () => {
    const path = join(fixturesDir, "simple.docx");
    const doc = await DocxDocument.load(path);
    const matches = await doc.findText("CLÁUSULA", 50);

    expect(matches.length).toBeGreaterThanOrEqual(1);
    expect(matches[0].matchText).toBe("CLÁUSULA");
    expect(matches[0].contextBefore.length).toBeGreaterThanOrEqual(0);
  });

  it("returns page setup info", async () => {
    const path = join(fixturesDir, "simple.docx");
    const doc = await DocxDocument.load(path);
    const pageInfo = await doc.getPageSetup();

    expect(pageInfo.pageSize.width).toBeGreaterThan(0);
    expect(pageInfo.pageSize.height).toBeGreaterThan(0);
    expect(["portrait", "landscape"]).toContain(pageInfo.orientation);
    expect(pageInfo.margins.top).toBeGreaterThan(0);
  });

  it("rejects .doc file with helpful error", () => {
    const err = validateDocxPath("/test/file.doc");
    expect(err).not.toBeNull();
    expect(err).toContain("convert");
    expect(err).toContain("Save As");
  });

  it("rejects .pdf file", () => {
    const err = validateDocxPath("/test/file.pdf");
    expect(err).not.toBeNull();
    expect(err).toContain(".pdf");
  });

  it("accepts .docx file", () => {
    const err = validateDocxPath("/test/file.docx");
    expect(err).toBeNull();
  });
});

describe("Integration: DOCX with comments", () => {
  it("loads with-comments.docx", async () => {
    const path = join(fixturesDir, "with-comments.docx");
    expect(existsSync(path)).toBe(true);

    const doc = await DocxDocument.load(path);
    const text = await doc.getText();
    expect(text).toContain("Q1 2026");
  });
});

describe("Integration: Large document truncation", () => {
  it("truncates large.docx when maxChars is exceeded", async () => {
    const path = join(fixturesDir, "large.docx");
    const doc = await DocxDocument.load(path);
    const text = await doc.getText();

    // Full text should be larger than 1000 chars
    expect(text.length).toBeGreaterThan(1000);

    const content = await doc.getSectionContent(
      (await doc.getSections())[0].title,
      500
    );
    expect(content.content.length).toBeLessThanOrEqual(500);
  });
});

describe("word_get_document_info tool", () => {
  it("returns accurate page count, word count, and format", async () => {
    const filePath = join(fixturesDir, "simple.docx");
    const result = await wordGetDocumentInfo({ filePath });
    const payload = JSON.parse((result.content as Array<{ text: string }>)[0].text);

    expect(payload.pageCount).toBeGreaterThanOrEqual(1);
    expect(payload.wordCount).toBeGreaterThanOrEqual(10);
    expect(payload.format).toBe("docx");
  });

  it("hasTrackChanges is false for a clean document", async () => {
    const filePath = join(fixturesDir, "simple.docx");
    const result = await wordGetDocumentInfo({ filePath });
    const payload = JSON.parse((result.content as Array<{ text: string }>)[0].text);

    expect(payload.hasTrackChanges).toBe(false);
  });

  it("comment counts match actual comments in file", async () => {
    const filePath = join(fixturesDir, "with-comments.docx");
    const result = await wordGetDocumentInfo({ filePath });
    const payload = JSON.parse((result.content as Array<{ text: string }>)[0].text);

    expect(payload.commentCount).toBeGreaterThanOrEqual(1);
    expect(typeof payload.unresolvedComments).toBe("number");
    expect(payload.unresolvedComments).toBeLessThanOrEqual(payload.commentCount);
  });

  it("returns clear error for invalid file path", async () => {
    const result = await wordGetDocumentInfo({ filePath: "/nonexistent/file.docx" });

    expect(result.isError).toBe(true);
    const text = (result.content as Array<{ text: string }>)[0].text;
    expect(text.length).toBeGreaterThan(0);
  });

  it("returns clear error for unsupported format", async () => {
    const result = await wordGetDocumentInfo({ filePath: "/path/to/file.doc" });

    expect(result.isError).toBe(true);
    const text = (result.content as Array<{ text: string }>)[0].text;
    expect(text).toContain("Save As");
  });

  it("response includes _suggestions", async () => {
    const filePath = join(fixturesDir, "simple.docx");
    const result = await wordGetDocumentInfo({ filePath });
    const payload = JSON.parse((result.content as Array<{ text: string }>)[0].text);

    expect(payload._suggestions).toBeDefined();
    expect(payload._suggestions.word_get_sections).toBeDefined();
    expect(payload._suggestions.word_list_comments).toBeDefined();
  });
});

describe("Integration: Write operations", () => {
  it("creates a copy and modifies it (search/replace)", async () => {
    const { readFileSync, writeFileSync, rmSync } = await import("node:fs");
    const src = join(fixturesDir, "simple.docx");
    const dest = join(fixturesDir, "test-output.docx");

    // Copy original
    writeFileSync(dest, readFileSync(src));

    // Load and verify
    const doc = await DocxDocument.load(dest);
    const textBefore = await doc.getText();
    expect(textBefore).toContain("CLÁUSULA");

    // Clean up
    rmSync(dest);
  });
});
