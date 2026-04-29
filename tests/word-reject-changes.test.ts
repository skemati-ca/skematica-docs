import { afterEach, describe, expect, it } from "vitest";
import { mkdtempSync, readFileSync, rmSync, writeFileSync } from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";
import JSZip from "jszip";
import { wordRejectChanges } from "../src/tools/word-reject-changes.js";

describe("wordRejectChanges", () => {
  const tempDirs: string[] = [];

  afterEach(() => {
    for (const dir of tempDirs.splice(0)) {
      rmSync(dir, { recursive: true, force: true });
    }
  });

  it("rejects w:ins by dropping inserted content entirely", async () => {
    const filePath = await createDocx([
      run("Before "),
      change("w:ins", "1", "Reviewer", "2026-04-20T10:00:00Z", run("Inserted")),
      run(" After"),
    ]);

    const result = await wordRejectChanges({ filePath });
    const documentXml = await readDocumentXml(filePath);

    expect(result.isError).not.toBe(true);
    expect(documentXml).not.toContain("<w:ins");
    expect(documentXml).not.toContain("Inserted");
    expect(documentXml).toContain("<w:t>Before </w:t>");
    expect(documentXml).toContain("<w:t> After</w:t>");
  });

  it("rejects w:del by restoring w:delText as w:t", async () => {
    const filePath = await createDocx([
      run("Keep "),
      change("w:del", "2", "Reviewer", "2026-04-20T10:00:00Z", delRun("Restored")),
      run(" Text"),
    ]);

    await wordRejectChanges({ filePath });
    const documentXml = await readDocumentXml(filePath);

    expect(documentXml).not.toContain("<w:del");
    expect(documentXml).not.toContain("w:delText");
    expect(documentXml).toContain("<w:t>Restored</w:t>");
  });

  it("filters rejected changes by author, date, and id", async () => {
    const filePath = await createDocx([
      change("w:ins", "1", "Alice", "2026-04-20T10:00:00Z", run("AliceText")),
      change("w:ins", "2", "Bob", "2026-04-20T10:00:00Z", run("BobText")),
      change("w:ins", "3", "Alice", "2026-04-01T10:00:00Z", run("OldAliceText")),
    ]);

    await wordRejectChanges({
      filePath,
      author: "Alice",
      dateFrom: "2026-04-10T00:00:00Z",
      dateTo: "2026-04-30T23:59:59Z",
      changeId: "1",
    });
    const documentXml = await readDocumentXml(filePath);

    expect(documentXml).not.toContain('w:ins w:id="1"');
    expect(documentXml).toContain('w:ins w:id="2" w:author="Bob"');
    expect(documentXml).toContain('w:ins w:id="3" w:author="Alice"');
  });

  it("response includes _suggestions", async () => {
    const filePath = await createDocx([
      change("w:del", "1", "Reviewer", "2026-04-20T10:00:00Z", delRun("Restored")),
    ]);

    const result = await wordRejectChanges({ filePath });
    const payload = JSON.parse((result.content as Array<{ text: string }>)[0].text);

    expect(payload.rejectedCount).toBe(1);
    expect(payload._suggestions.word_get_document_info).toBeDefined();
  });

  async function createDocx(paragraphChildren: string[]): Promise<string> {
    const tempDir = mkdtempSync(join(tmpdir(), "skematica-reject-changes-"));
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
