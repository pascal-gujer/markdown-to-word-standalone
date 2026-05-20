import { describe, expect, it } from "vitest";
import JSZip from "jszip";
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

  it("keeps Markdown image references in the model and warns when no ZIP image bundle is loaded", () => {
    const parsed = parseMarkdown("![Logo](images/logo.png) caption text");

    expect(parsed.model.blocks[0]).toMatchObject({
      type: "paragraph",
      children: [
        {
          image: {
            alt: "Logo",
            src: "images/logo.png",
          },
        },
        {
          text: " caption text",
        },
      ],
    });
    expect(parsed.warnings).toContain("warning.images");
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

  it("embeds supported ZIP image assets into the DOCX package", async () => {
    const parsed = parseMarkdown(`# Images

![Logo](images/logo.png)
![Photo](images/photo.jpeg)
![Badge](images/badge.gif)
`, { imageMode: "embedded" });

    const blob = await generateDocxBlob(parsed.model, {
      images: {
        sourceFileName: "bundle.zip",
        markdownPath: "document.md",
        markdownBasePath: "",
        assets: [{
          path: "images/logo.png",
          fileName: "logo.png",
          extension: "png",
          contentType: "image/png",
          data: tinyPngBytes(),
          widthPx: 1,
          heightPx: 1,
        }, {
          path: "images/photo.jpeg",
          fileName: "photo.jpeg",
          extension: "jpeg",
          contentType: "image/jpeg",
          data: minimalJpegBytes(3, 2),
          widthPx: 3,
          heightPx: 2,
        }, {
          path: "images/badge.gif",
          fileName: "badge.gif",
          extension: "gif",
          contentType: "image/gif",
          data: tinyGifBytes(),
          widthPx: 1,
          heightPx: 1,
        }],
      },
      title: "image-test",
    });
    const zip = await JSZip.loadAsync(await blob.arrayBuffer());
    const documentXml = await zip.file("word/document.xml")?.async("string");
    const relsXml = await zip.file("word/_rels/document.xml.rels")?.async("string");
    const contentTypesXml = await zip.file("[Content_Types].xml")?.async("string");

    expect(zip.file("word/media/markdown-image-1.png")).toBeTruthy();
    expect(zip.file("word/media/markdown-image-2.jpg")).toBeTruthy();
    expect(zip.file("word/media/markdown-image-3.gif")).toBeTruthy();
    expect(documentXml).toContain("<w:drawing>");
    expect(documentXml).toContain('descr="Logo"');
    expect(documentXml).toContain('descr="Photo"');
    expect(documentXml).toContain('descr="Badge"');
    expect(documentXml).not.toContain("[Logo]");
    expect(relsXml).toContain("/image");
    expect(relsXml).toContain('Target="media/markdown-image-1.png"');
    expect(relsXml).toContain('Target="media/markdown-image-2.jpg"');
    expect(relsXml).toContain('Target="media/markdown-image-3.gif"');
    expect(contentTypesXml).toContain('Extension="gif" ContentType="image/gif"');
    expect(contentTypesXml).toContain('Extension="jpg" ContentType="image/jpeg"');
    expect(contentTypesXml).toContain('Extension="jpeg" ContentType="image/jpeg"');
    expect(contentTypesXml).toContain('Extension="png" ContentType="image/png"');
  });

  it("caps a tall portrait image's height so it fits on a single page while preserving aspect ratio", async () => {
    const parsed = parseMarkdown(`![Portrait](images/portrait.png)`, { imageMode: "embedded" });

    const blob = await generateDocxBlob(parsed.model, {
      images: {
        sourceFileName: "bundle.zip",
        markdownPath: "report.md",
        markdownBasePath: "",
        assets: [{
          path: "images/portrait.png",
          fileName: "portrait.png",
          extension: "png",
          contentType: "image/png",
          data: tinyPngBytes(),
          widthPx: 1284,
          heightPx: 2778,
        }],
      },
      title: "portrait-cap-test",
    });
    const zip = await JSZip.loadAsync(await blob.arrayBuffer());
    const documentXml = (await zip.file("word/document.xml")?.async("string")) || "";

    const extentMatch = /<wp:extent\b[^/]*\bcx="(\d+)"\s+cy="(\d+)"/.exec(documentXml);
    expect(extentMatch).toBeTruthy();
    const cx = Number(extentMatch?.[1]);
    const cy = Number(extentMatch?.[2]);
    expect(cy).toBeLessThanOrEqual(6_400_800);
    expect(cx).toBeLessThanOrEqual(5_760_000);
    const aspect = 1284 / 2778;
    expect(cx / cy).toBeCloseTo(aspect, 2);
  });

  it("assigns unique wp:docPr ids when the same image is referenced multiple times", async () => {
    const parsed = parseMarkdown(`# Repeats

![Logo](images/logo.png)
![Logo again](images/logo.png)
![Logo once more](images/logo.png)
`, { imageMode: "embedded" });

    const blob = await generateDocxBlob(parsed.model, {
      images: {
        sourceFileName: "bundle.zip",
        markdownPath: "report.md",
        markdownBasePath: "",
        assets: [{
          path: "images/logo.png",
          fileName: "logo.png",
          extension: "png",
          contentType: "image/png",
          data: tinyPngBytes(),
          widthPx: 1,
          heightPx: 1,
        }],
      },
      title: "repeat-test",
    });
    const zip = await JSZip.loadAsync(await blob.arrayBuffer());
    const documentXml = await zip.file("word/document.xml")?.async("string");
    const relsXml = await zip.file("word/_rels/document.xml.rels")?.async("string");

    const docPrIds = Array.from((documentXml || "").matchAll(/<wp:docPr\b[^>]*\bid="(\d+)"/g)).map((match) => match[1]);
    expect(docPrIds).toHaveLength(3);
    expect(new Set(docPrIds).size).toBe(3);

    const imageRelTargets = Array.from((relsXml || "").matchAll(/Target="(media\/markdown-image-[^"]+)"/g)).map((match) => match[1]);
    expect(imageRelTargets).toHaveLength(1);
    expect(zip.file(`word/${imageRelTargets[0]}`)).toBeTruthy();
  });

  it("wraps an image inside a w:hyperlink when the Markdown image is itself a link", async () => {
    const parsed = parseMarkdown(`[![Logo](images/logo.png)](https://example.invalid/home)
`, { imageMode: "embedded" });

    const blob = await generateDocxBlob(parsed.model, {
      images: {
        sourceFileName: "bundle.zip",
        markdownPath: "report.md",
        markdownBasePath: "",
        assets: [{
          path: "images/logo.png",
          fileName: "logo.png",
          extension: "png",
          contentType: "image/png",
          data: tinyPngBytes(),
          widthPx: 1,
          heightPx: 1,
        }],
      },
      title: "image-link-test",
    });
    const zip = await JSZip.loadAsync(await blob.arrayBuffer());
    const documentXml = await zip.file("word/document.xml")?.async("string");
    const relsXml = await zip.file("word/_rels/document.xml.rels")?.async("string");

    expect(documentXml).toMatch(/<w:hyperlink[^>]*>\s*<w:r>[\s\S]*<w:drawing>[\s\S]*<\/w:drawing>[\s\S]*<\/w:r>\s*<\/w:hyperlink>/);
    expect(relsXml).toContain('Target="https://example.invalid/home"');
  });

  it("falls back to bracketed alt text wrapped in a hyperlink when an image-as-link has no bundle asset", async () => {
    const parsed = parseMarkdown(`[![Missing](images/missing.png)](https://example.invalid/away)
`);

    const blob = await generateDocxBlob(parsed.model, { title: "missing-link-test" });
    const zip = await JSZip.loadAsync(await blob.arrayBuffer());
    const documentXml = await zip.file("word/document.xml")?.async("string");

    expect(documentXml).toMatch(/<w:hyperlink[^>]*>\s*<w:r>[\s\S]*\[Missing\][\s\S]*<\/w:r>\s*<\/w:hyperlink>/);
    expect(documentXml).toContain('<w:rStyle w:val="Hyperlink"/>');
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
    const templateBytes = await createPreservationTemplateFixture();
    const template = await readTemplatePackage(templateBytes, "fixture-template.docx");
    const parsed = parseMarkdown(`# Inserted Report

Body text with [new link](https://example.invalid/new).

![Inline logo](images/inline-logo.png)

- One
- Two
`, { imageMode: "embedded" });

    const blob = await generateDocxBlob(parsed.model, {
      images: {
        sourceFileName: "bundle.zip",
        markdownPath: "report.md",
        markdownBasePath: "",
        assets: [{
          path: "images/inline-logo.png",
          fileName: "inline-logo.png",
          extension: "png",
          contentType: "image/png",
          data: tinyPngBytes(),
          widthPx: 1,
          heightPx: 1,
        }],
      },
      template,
      title: "template-preservation-test",
    });
    const zip = await JSZip.loadAsync(await blob.arrayBuffer());
    const documentXml = await zip.file("word/document.xml")?.async("string");
    const relsXml = await zip.file("word/_rels/document.xml.rels")?.async("string");
    const numberingXml = await zip.file("word/numbering.xml")?.async("string");
    const contentTypesXml = await zip.file("[Content_Types].xml")?.async("string");

    expect(zip.file("word/header1.xml")).toBeTruthy();
    expect(zip.file("word/footer1.xml")).toBeTruthy();
    expect(zip.file("word/_rels/header1.xml.rels")).toBeTruthy();
    expect(zip.file("word/media/image1.png")).toBeTruthy();
    expect(zip.file("word/media/markdown-image-1.png")).toBeTruthy();
    expect(zip.file("customXml/item1.xml")).toBeTruthy();
    expect(zip.file("word/settings.xml")).toBeTruthy();
    expect(zip.file("word/fontTable.xml")).toBeTruthy();
    expect(zip.file("word/footnotes.xml")).toBeTruthy();
    expect(zip.file("word/endnotes.xml")).toBeTruthy();

    expect(documentXml).toContain("Inserted Report");
    expect(documentXml).toContain("xmlns:wp=");
    expect(documentXml).toContain("<w:drawing>");
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

async function createPreservationTemplateFixture(): Promise<Uint8Array> {
  const zip = new JSZip();
  const pkgRelNs = "http://schemas.openxmlformats.org/package/2006/relationships";
  const docRelNs = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
  const contentTypesNs = "http://schemas.openxmlformats.org/package/2006/content-types";
  const wordNs = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

  zip.file("[Content_Types].xml", [
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
    `<Types xmlns="${contentTypesNs}">`,
    '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>',
    '<Default Extension="xml" ContentType="application/xml"/>',
    '<Default Extension="png" ContentType="image/png"/>',
    '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>',
    '<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>',
    '<Override PartName="/word/numbering.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/>',
    '<Override PartName="/word/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/>',
    '<Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/>',
    '<Override PartName="/word/fontTable.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.fontTable+xml"/>',
    '<Override PartName="/word/footnotes.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml"/>',
    '<Override PartName="/word/endnotes.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.endnotes+xml"/>',
    '<Override PartName="/word/header1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/>',
    '<Override PartName="/word/footer1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"/>',
    '<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>',
    '<Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>',
    '<Override PartName="/customXml/itemProps1.xml" ContentType="application/vnd.openxmlformats-officedocument.customXmlProperties+xml"/>',
    "</Types>",
  ].join(""));

  zip.folder("_rels")?.file(".rels", [
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
    `<Relationships xmlns="${pkgRelNs}">`,
    `<Relationship Id="rId1" Type="${docRelNs}/officeDocument" Target="word/document.xml"/>`,
    '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>',
    `<Relationship Id="rId3" Type="${docRelNs}/extended-properties" Target="docProps/app.xml"/>`,
    "</Relationships>",
  ].join(""));

  const word = zip.folder("word");
  word?.file("document.xml", [
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
    `<w:document xmlns:w="${wordNs}" xmlns:r="${docRelNs}">`,
    "<w:body>",
    "<w:p><w:r><w:t>I hope you keep this template shell</w:t></w:r></w:p>",
    '<w:sectPr><w:headerReference w:type="default" r:id="rId11"/><w:footerReference w:type="default" r:id="rId12"/><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="900" w:right="900" w:bottom="900" w:left="900" w:header="450" w:footer="450" w:gutter="0"/></w:sectPr>',
    "</w:body>",
    "</w:document>",
  ].join(""));

  word?.folder("_rels")?.file("document.xml.rels", [
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
    `<Relationships xmlns="${pkgRelNs}">`,
    `<Relationship Id="rId1" Type="${docRelNs}/customXml" Target="../customXml/item1.xml"/>`,
    `<Relationship Id="rId2" Type="${docRelNs}/numbering" Target="numbering.xml"/>`,
    `<Relationship Id="rId3" Type="${docRelNs}/styles" Target="styles.xml"/>`,
    `<Relationship Id="rId4" Type="${docRelNs}/settings" Target="settings.xml"/>`,
    `<Relationship Id="rId5" Type="${docRelNs}/webSettings" Target="webSettings.xml"/>`,
    `<Relationship Id="rId6" Type="${docRelNs}/footnotes" Target="footnotes.xml"/>`,
    `<Relationship Id="rId7" Type="${docRelNs}/endnotes" Target="endnotes.xml"/>`,
    `<Relationship Id="rId8" Type="${docRelNs}/hyperlink" Target="https://example.invalid/existing" TargetMode="External"/>`,
    `<Relationship Id="rId11" Type="${docRelNs}/header" Target="header1.xml"/>`,
    `<Relationship Id="rId12" Type="${docRelNs}/footer" Target="footer1.xml"/>`,
    `<Relationship Id="rId13" Type="${docRelNs}/fontTable" Target="fontTable.xml"/>`,
    `<Relationship Id="rId14" Type="${docRelNs}/theme" Target="theme/theme1.xml"/>`,
    "</Relationships>",
  ].join(""));

  word?.file("styles.xml", [
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
    `<w:styles xmlns:w="${wordNs}">`,
    '<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>',
    '<w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="Heading 1"/></w:style>',
    '<w:style w:type="character" w:styleId="Hyperlink"><w:name w:val="Hyperlink"/></w:style>',
    "</w:styles>",
  ].join(""));
  word?.file("numbering.xml", [
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
    `<w:numbering xmlns:w="${wordNs}">`,
    '<w:abstractNum w:abstractNumId="8"><w:lvl w:ilvl="0"><w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/></w:lvl></w:abstractNum>',
    '<w:num w:numId="25"><w:abstractNumId w:val="8"/></w:num>',
    "</w:numbering>",
  ].join(""));
  word?.folder("theme")?.file("theme1.xml", '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="Fixture Theme"/>');
  word?.file("settings.xml", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:settings xmlns:w="${wordNs}"><w:updateFields w:val="true"/></w:settings>`);
  word?.file("webSettings.xml", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:webSettings xmlns:w="${wordNs}"/>`);
  word?.file("fontTable.xml", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:fonts xmlns:w="${wordNs}"><w:font w:name="Aptos"/></w:fonts>`);
  word?.file("footnotes.xml", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:footnotes xmlns:w="${wordNs}"><w:footnote w:id="0"/></w:footnotes>`);
  word?.file("endnotes.xml", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:endnotes xmlns:w="${wordNs}"><w:endnote w:id="0"/></w:endnotes>`);
  word?.file("header1.xml", [
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
    `<w:hdr xmlns:w="${wordNs}" xmlns:r="${docRelNs}" xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office">`,
    '<w:p><w:r><w:pict><v:shape id="Watermark" style="position:absolute;width:100pt;height:100pt" type="#_x0000_t75"><v:imagedata r:id="rId1" o:title="watermark"/></v:shape></w:pict></w:r></w:p>',
    "</w:hdr>",
  ].join(""));
  word?.folder("_rels")?.file("header1.xml.rels", [
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
    `<Relationships xmlns="${pkgRelNs}">`,
    `<Relationship Id="rId1" Type="${docRelNs}/image" Target="media/image1.png"/>`,
    "</Relationships>",
  ].join(""));
  word?.file("footer1.xml", [
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
    `<w:ftr xmlns:w="${wordNs}">`,
    '<w:p><w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> PAGE </w:instrText></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r></w:p>',
    "</w:ftr>",
  ].join(""));
  word?.folder("media")?.file("image1.png", Uint8Array.from([
    ...tinyPngBytes(),
  ]));

  zip.folder("customXml")?.file("item1.xml", '<company><field name="retained">yes</field></company>');
  zip.folder("customXml")?.file("itemProps1.xml", '<ds:datastoreItem ds:itemID="{11111111-1111-1111-1111-111111111111}" xmlns:ds="http://schemas.openxmlformats.org/officeDocument/2006/customXml"/>');
  zip.folder("docProps")?.file("core.xml", '<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties"/>');
  zip.folder("docProps")?.file("app.xml", '<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"><Application>Fixture</Application></Properties>');

  return zip.generateAsync({ type: "uint8array" });
}

function tinyPngBytes(): Uint8Array {
  return Uint8Array.from([
    137, 80, 78, 71, 13, 10, 26, 10, 0, 0, 0, 13, 73, 72, 68, 82,
    0, 0, 0, 1, 0, 0, 0, 1, 8, 6, 0, 0, 0, 31, 21, 196,
    137, 0, 0, 0, 13, 73, 68, 65, 84, 120, 156, 99, 248, 255, 255,
    63, 0, 5, 254, 2, 254, 167, 53, 129, 132, 0, 0, 0, 0,
    73, 69, 78, 68, 174, 66, 96, 130,
  ]);
}

function tinyGifBytes(): Uint8Array {
  return Uint8Array.from([
    71, 73, 70, 56, 57, 97, 1, 0, 1, 0, 128, 0, 0, 0, 0, 0,
    255, 255, 255, 44, 0, 0, 0, 0, 1, 0, 1, 0, 0, 2, 1, 76,
    0, 59,
  ]);
}

function minimalJpegBytes(width: number, height: number): Uint8Array {
  return Uint8Array.from([
    0xff, 0xd8,
    0xff, 0xe0, 0x00, 0x10, 0x4a, 0x46, 0x49, 0x46, 0x00, 0x01, 0x01, 0x00, 0x00, 0x01, 0x00, 0x01, 0x00, 0x00,
    0xff, 0xc0, 0x00, 0x11, 0x08, (height >> 8) & 0xff, height & 0xff, (width >> 8) & 0xff, width & 0xff,
    0x03, 0x01, 0x11, 0x00, 0x02, 0x11, 0x00, 0x03, 0x11, 0x00,
    0xff, 0xd9,
  ]);
}
