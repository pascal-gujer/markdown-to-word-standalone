import { describe, expect, it } from "vitest";
import JSZip from "jszip";
import { readFile } from "node:fs/promises";
import { generateDocxBlob } from "../src/docx/generateDocx";
import { parseMarkdown } from "../src/markdown/parseMarkdown";
import { readTemplatePackage, type TemplateInfo } from "../src/docx/templateReader";

describe("Markdown conversion", () => {
  it("builds a structured model for common Markdown blocks", () => {
    const parsed = parseMarkdown(`# Grüsse 🚀

Text with **bold**, *italic*, and \`code\`.

- Alpha
  - Beta

| A | B |
| - | - |
| 1 | 2 |
`);

    expect(parsed.model.blocks.map((block) => block.type)).toEqual(["heading", "paragraph", "list", "table"]);
    expect(parsed.html).toContain("Grüsse");
    expect(parsed.warnings).toEqual([]);
  });

  it("exports a DOCX package with document XML, styles, numbering, tables, and links", async () => {
    const parsed = parseMarkdown(`# Title

Hello **bold** [link](https://example.invalid).

> Quote

| Feature | Status |
| --- | --- |
| Tables | Yes |

\`\`\`
const value = "Grüsse";
\`\`\`
`);

    const blob = await generateDocxBlob(parsed.model, { title: "unit-test" });
    const zip = await JSZip.loadAsync(await blob.arrayBuffer());
    const documentXml = await zip.file("word/document.xml")?.async("string");
    const stylesXml = await zip.file("word/styles.xml")?.async("string");
    const relsXml = await zip.file("word/_rels/document.xml.rels")?.async("string");
    const contentTypesXml = await zip.file("[Content_Types].xml")?.async("string");

    expect(zip.file("[Content_Types].xml")).toBeTruthy();
    expect(zip.file("word/theme/theme1.xml")).toBeNull();
    expect(contentTypesXml).not.toContain("/word/theme/theme1.xml");
    expect(documentXml).toContain('<w:pStyle w:val="Heading1"');
    expect(documentXml).toContain("<w:tbl>");
    expect(documentXml).toContain("<w:hyperlink");
    expect(documentXml).toContain("Grüsse");
    expect(stylesXml).toContain('w:styleId="CodeBlock"');
    expect(stylesXml).toContain('w:styleId="TableGrid"');
    expect(relsXml).toContain('Target="https://example.invalid"');
    expect(relsXml).not.toContain("/theme");

    const normalStyle = extractStyle(stylesXml || "", "Normal");
    const quoteStyle = extractStyle(stylesXml || "", "Quote");
    const codeStyle = extractStyle(stylesXml || "", "CodeBlock");
    expect(normalStyle.indexOf("<w:pPr>")).toBeLessThan(normalStyle.indexOf("<w:rPr>"));
    expect(quoteStyle.indexOf("<w:spacing")).toBeLessThan(quoteStyle.indexOf("<w:ind"));
    expect(codeStyle.indexOf("<w:shd")).toBeLessThan(codeStyle.indexOf("<w:spacing"));
  });

  it("preserves a template theme only when a template provides one", async () => {
    const parsed = parseMarkdown("# Title");
    const themeXml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><a:theme xmlns:a="urn:test" name="Template Theme"/>';
    const template: TemplateInfo = {
      fileName: "template.dotx",
      styles: [],
      themeXml,
      hasDocumentPart: false,
      hasStylesPart: false,
      hasThemePart: true,
      hasNumberingPart: false,
    };

    const blob = await generateDocxBlob(parsed.model, { template });
    const zip = await JSZip.loadAsync(await blob.arrayBuffer());
    const relsXml = await zip.file("word/_rels/document.xml.rels")?.async("string");

    await expect(zip.file("word/theme/theme1.xml")?.async("string")).resolves.toBe(themeXml);
    expect(relsXml).toContain("theme/theme1.xml");
  });

  it("uses a DOCX template as the base package and keeps headers, footers, media, fields, and custom parts", async () => {
    const templateBytes = await readFile("issues/2/WordTemplate.docx");
    const template = await readTemplatePackage(templateBytes, "WordTemplate.docx");
    const parsed = parseMarkdown(`# Inserted Report

Body text with [new link](https://example.invalid/new).

- One
- Two
`);

    const blob = await generateDocxBlob(parsed.model, { template, title: "template-preservation-test" });
    const zip = await JSZip.loadAsync(await blob.arrayBuffer());
    const documentXml = await zip.file("word/document.xml")?.async("string");
    const relsXml = await zip.file("word/_rels/document.xml.rels")?.async("string");
    const numberingXml = await zip.file("word/numbering.xml")?.async("string");
    const contentTypesXml = await zip.file("[Content_Types].xml")?.async("string");

    expect(zip.file("word/header1.xml")).toBeTruthy();
    expect(zip.file("word/footer1.xml")).toBeTruthy();
    expect(zip.file("word/_rels/header1.xml.rels")).toBeTruthy();
    expect(zip.file("word/media/image1.png")).toBeTruthy();
    expect(zip.file("customXml/item1.xml")).toBeTruthy();
    expect(zip.file("word/settings.xml")).toBeTruthy();
    expect(zip.file("word/fontTable.xml")).toBeTruthy();
    expect(zip.file("word/footnotes.xml")).toBeTruthy();
    expect(zip.file("word/endnotes.xml")).toBeTruthy();

    expect(documentXml).toContain("Inserted Report");
    expect(documentXml).not.toContain("I hope you");
    expect(documentXml).toContain('<w:headerReference w:type="default" r:id="rId11"/>');
    expect(documentXml).toContain('<w:footerReference w:type="default" r:id="rId12"/>');
    expect(relsXml).toContain('Target="header1.xml"');
    expect(relsXml).toContain('Target="footer1.xml"');
    expect(relsXml).toContain('Target="https://example.invalid/new" TargetMode="External"');
    expect(numberingXml).toContain('w:numId="26"');
    expect(numberingXml).toContain('w:abstractNumId="9"');
    expect(contentTypesXml).toContain('PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"');
  });
});

function extractStyle(stylesXml: string, styleId: string): string {
  const pattern = new RegExp(`<w:style\\b[^>]*\\bw:styleId="${styleId}"[\\s\\S]*?<\\/w:style>`);
  const match = stylesXml.match(pattern);
  expect(match).toBeTruthy();
  return match ? match[0] : "";
}
