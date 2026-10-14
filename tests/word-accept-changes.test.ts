import { afterEach, describe, expect, it } from "vitest";
import { mkdtempSync, rmSync, writeFileSync, readFileSync } from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";
import JSZip from "jszip";
import { wordAcceptChanges } from "../src/tools/word-accept-changes.js";

describe("wordAcceptChanges", () => {
  const tempDirs: string[] = [];

  afterEach(() => {
    for (const dir of tempDirs.splice(0)) {
      rmSync(dir, { recursive: true, force: true });
    }
  });

  it("accepts w:ins by promoting children and removing the wrapper", async () => {
    const filePath = await createDocx([
      run("Before "),
      change("w:ins", "1", "Reviewer", "2026-04-20T10:00:00Z", run("Inserted")),
      run(" After"),
    ]);

    const result = await wordAcceptChanges({ filePath });
    const documentXml = await readDocumentXml(filePath);

    expect(result.isError).not.toBe(true);
    expect(documentXml).not.toContain("<w:ins");
    expect(documentXml).toContain("<w:t>Inserted</w:t>");
  });

  it("accepts w:del by dropping deleted content entirely", async () => {
    const filePath = await createDocx([
      run("Keep "),
      change("w:del", "2", "Reviewer", "2026-04-20T10:00:00Z", delRun("Removed")),
      run(" Text"),
    ]);

    await wordAcceptChanges({ filePath });
    const documentXml = await readDocumentXml(filePath);

    expect(documentXml).not.toContain("<w:del");
    expect(documentXml).not.toContain("Removed");
    expect(documentXml).toContain("<w:t>Keep </w:t>");
    expect(documentXml).toContain("<w:t> Text</w:t>");
  });

  it("removes accepted pPrChange, rPrChange, and numberingChange records", async () => {
    const filePath = await createDocx([
      '<w:pPr><w:pPrChange w:id="3" w:author="Reviewer" w:date="2026-04-20T10:00:00Z"><w:pPr><w:jc w:val="left"/></w:pPr></w:pPrChange></w:pPr>',
      '<w:r><w:rPr><w:rPrChange w:id="4" w:author="Reviewer" w:date="2026-04-20T10:00:00Z"><w:rPr><w:b/></w:rPr></w:rPrChange></w:rPr><w:t>Text</w:t></w:r>',
      '<w:numberingChange w:id="5" w:author="Reviewer" w:date="2026-04-20T10:00:00Z"/>',
    ]);

    await wordAcceptChanges({ filePath });
    const documentXml = await readDocumentXml(filePath);

    expect(documentXml).not.toContain("w:pPrChange");
    expect(documentXml).not.toContain("w:rPrChange");
    expect(documentXml).not.toContain("w:numberingChange");
    expect(documentXml).toContain("<w:t>Text</w:t>");
  });

  it("filters accepted changes by author", async () => {
    const filePath = await createDocx([
      change("w:ins", "1", "Alice", "2026-04-20T10:00:00Z", run("AliceText")),
      change("w:ins", "2", "Bob", "2026-04-20T10:00:00Z", run("BobText")),
    ]);

    await wordAcceptChanges({ filePath, author: "Alice" });
    const documentXml = await readDocumentXml(filePath);

    expect(documentXml).toContain("<w:t>AliceText</w:t>");
    expect(documentXml).toContain('w:ins w:id="2" w:author="Bob"');
  });

  it("filters accepted changes by date range", async () => {
    const filePath = await createDocx([
      change("w:ins", "1", "Reviewer", "2026-04-01T10:00:00Z", run("Old")),
      change("w:ins", "2", "Reviewer", "2026-04-20T10:00:00Z", run("Current")),
    ]);

    await wordAcceptChanges({
      filePath,
      dateFrom: "2026-04-10T00:00:00Z",
      dateTo: "2026-04-30T23:59:59Z",
    });
    const documentXml = await readDocumentXml(filePath);

    expect(documentXml).toContain('w:ins w:id="1"');
    expect(documentXml).toContain("<w:t>Current</w:t>");
    expect(documentXml).not.toContain('w:ins w:id="2"');
  });

  it("filters accepted changes by change ID", async () => {
    const filePath = await createDocx([
      change("w:ins", "1", "Reviewer", "2026-04-20T10:00:00Z", run("One")),
      change("w:ins", "2", "Reviewer", "2026-04-20T10:00:00Z", run("Two")),
    ]);

    await wordAcceptChanges({ filePath, changeId: "2" });
    const documentXml = await readDocumentXml(filePath);

    expect(documentXml).toContain('w:ins w:id="1"');
    expect(documentXml).toContain("<w:t>Two</w:t>");
    expect(documentXml).not.toContain('w:ins w:id="2"');
  });

  it("response includes _suggestions", async () => {
    const filePath = await createDocx([
      change("w:ins", "1", "Reviewer", "2026-04-20T10:00:00Z", run("Inserted")),
    ]);

    const result = await wordAcceptChanges({ filePath });
    const payload = JSON.parse((result.content as Array<{ text: string }>)[0].text);

    expect(payload.acceptedCount).toBe(1);
    expect(payload._suggestions.word_get_document_info).toBeDefined();
  });

  async function createDocx(paragraphChildren: string[]): Promise<string> {
    const tempDir = mkdtempSync(join(tmpdir(), "skematica-accept-changes-"));
    tempDirs.push(tempDir);
    const filePath = join(tempDir, "fixture.docx");
    const zip = new JSZip();
    zip.file("word/document.xml", documentXml(paragraphChildren.join("")));
    writeFileSync(filePath, await zip.generateAsync({ type: "nodebuffer", compression: "DEFLATE" }));
    return filePath;
  }
});

function documentXml(paragraphContent: string): string {
  return [
    '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">',
    "<w:body>",
    `<w:p>${paragraphContent}</w:p>`,
    "</w:body>",
    "</w:document>",
  ].join("");
}

function run(text: string): string {
  return `<w:r><w:t>${text}</w:t></w:r>`;
}

function delRun(text: string): string {
  return `<w:r><w:delText>${text}</w:delText></w:r>`;
}

function change(tag: "w:ins" | "w:del", id: string, author: string, date: string, content: string): string {
  return `<${tag} w:id="${id}" w:author="${author}" w:date="${date}">${content}</${tag}>`;
}

async function readDocumentXml(filePath: string): Promise<string> {
  const zip = await JSZip.loadAsync(readFileSync(filePath));
  return (await zip.file("word/document.xml")?.async("text")) ?? "";
}
