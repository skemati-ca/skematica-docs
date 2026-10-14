import { afterEach, describe, expect, it } from "vitest";
import { mkdtempSync, rmSync, writeFileSync } from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";
import JSZip from "jszip";
import { wordInsertTrackedChange } from "../src/tools/word-insert-tracked-change.js";

describe("wordInsertTrackedChange", () => {
  const tempDirs: string[] = [];

  afterEach(() => {
    for (const dir of tempDirs.splice(0)) {
      rmSync(dir, { recursive: true, force: true });
    }
  });

  it("wraps a match spanning two runs with the same rPr in tracked changes", async () => {
    const filePath = await createDocxWithParagraph([
      run("Con", boldRpr()),
      run("tract", boldRpr()),
    ]);

    const result = await wordInsertTrackedChange({
      filePath,
      searchText: "Contract",
      replacementText: "Agreement",
      author: "Tester",
    });
    const documentXml = await readDocumentXml(filePath);

    expect(result.isError).not.toBe(true);
    expect(documentXml).toContain("<w:del");
    expect(documentXml).toContain("<w:ins");
    expect(documentXml).toContain("<w:delText>Contract</w:delText>");
    expect(documentXml).toContain("<w:t>Agreement</w:t>");
    expect(documentXml).toContain("<w:rPr><w:b></w:b></w:rPr>");
  });

  it("preserves rPr on each deleted segment when a match spans runs with different rPr", async () => {
    const filePath = await createDocxWithParagraph([
      run("Con", boldRpr()),
      run("tr", italicRpr()),
      run("act", colorRpr("FF0000")),
    ]);

    const result = await wordInsertTrackedChange({
      filePath,
      searchText: "Contract",
      replacementText: "Agreement",
      author: "Tester",
    });
    const documentXml = await readDocumentXml(filePath);

    expect(result.isError).not.toBe(true);
    expect(documentXml).toContain("<w:delText>Con</w:delText>");
    expect(documentXml).toContain("<w:delText>tr</w:delText>");
    expect(documentXml).toContain("<w:delText>act</w:delText>");
    expect(documentXml).toContain("<w:rPr><w:b></w:b></w:rPr><w:delText>Con</w:delText>");
    expect(documentXml).toContain("<w:rPr><w:i></w:i></w:rPr><w:delText>tr</w:delText>");
    expect(documentXml).toContain('<w:rPr><w:color w:val="FF0000"></w:color></w:rPr><w:delText>act</w:delText>');
  });

  it("handles a cross-run match that starts and ends on run boundaries", async () => {
    const filePath = await createDocxWithParagraph([
      run("Before "),
      run("Cross"),
      run("Run"),
      run(" After"),
    ]);

    const result = await wordInsertTrackedChange({
      filePath,
      searchText: "CrossRun",
      replacementText: "Joined",
      author: "Tester",
    });
    const documentXml = await readDocumentXml(filePath);

    expect(result.isError).not.toBe(true);
    expect(documentXml).toMatch(/<w:t[^>]*>Before <\/w:t>/);
    expect(documentXml).toContain("<w:delText>CrossRun</w:delText>");
    expect(documentXml).toContain("<w:t>Joined</w:t>");
    expect(documentXml).toMatch(/<w:t[^>]*> After<\/w:t>/);
  });

  it("keeps single-run tracked changes working", async () => {
    const filePath = await createDocxWithParagraph([
      run("Before Contract After", boldRpr()),
    ]);

    const result = await wordInsertTrackedChange({
      filePath,
      searchText: "Contract",
      replacementText: "Agreement",
      author: "Tester",
    });
    const documentXml = await readDocumentXml(filePath);

    expect(result.isError).not.toBe(true);
    expect(documentXml).toMatch(/<w:t[^>]*>Before <\/w:t>/);
    expect(documentXml).toContain("<w:delText>Contract</w:delText>");
    expect(documentXml).toContain("<w:t>Agreement</w:t>");
    expect(documentXml).toMatch(/<w:t[^>]*> After<\/w:t>/);
  });

  async function createDocxWithParagraph(runs: string[]): Promise<string> {
    const tempDir = mkdtempSync(join(tmpdir(), "skematica-tracked-change-"));
    tempDirs.push(tempDir);
    const filePath = join(tempDir, "fixture.docx");
    const zip = new JSZip();
    zip.file("word/document.xml", documentXml(runs.join("")));
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

function run(text: string, rPr = ""): string {
  const space = text !== text.trim() ? ' xml:space="preserve"' : "";
  return `<w:r>${rPr}<w:t${space}>${text}</w:t></w:r>`;
}

function boldRpr(): string {
  return "<w:rPr><w:b/></w:rPr>";
}

function italicRpr(): string {
  return "<w:rPr><w:i/></w:rPr>";
}

function colorRpr(color: string): string {
  return `<w:rPr><w:color w:val="${color}"/></w:rPr>`;
}

async function readDocumentXml(filePath: string): Promise<string> {
  const content = await import("node:fs").then(({ readFileSync }) => readFileSync(filePath));
  const zip = await JSZip.loadAsync(content);
  return (await zip.file("word/document.xml")?.async("text")) ?? "";
}
