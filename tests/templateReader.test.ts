import { describe, expect, it } from "vitest";
import JSZip from "jszip";
import { buildStyleMap } from "../src/docx/styleMapper";
import { readTemplatePackage } from "../src/docx/templateReader";

describe("template reader", () => {
  it("reads DOCX/DOTX style, theme, numbering, and section parts", async () => {
    const zip = new JSZip();
    zip.file("[Content_Types].xml", "<Types/>");
    zip.file("word/styles.xml", [
      '<w:styles xmlns:w="urn:test">',
      '<w:style w:type="paragraph" w:styleId="MyHeading1"><w:name w:val="Heading 1"/></w:style>',
      '<w:style w:type="paragraph" w:styleId="MyNormal"><w:name w:val="Normal"/></w:style>',
      '<w:style w:type="paragraph" w:styleId="MyQuote"><w:name w:val="Quote"/></w:style>',
      '<w:style w:type="table" w:styleId="MyTable"><w:name w:val="Table Grid"/></w:style>',
      "</w:styles>",
    ].join(""));
    zip.file("word/theme/theme1.xml", "<theme/>");
    zip.file("word/numbering.xml", "<numbering/>");
    zip.file("word/document.xml", '<w:document><w:body><w:p/><w:sectPr><w:pgSz w:w="1"/></w:sectPr></w:body></w:document>');

    const bytes = await zip.generateAsync({ type: "uint8array" });
    const template = await readTemplatePackage(bytes, "template.dotx");
    const styleMap = buildStyleMap(template);

    expect(template.styles).toHaveLength(4);
    expect(template.hasThemePart).toBe(true);
    expect(template.hasNumberingPart).toBe(true);
    expect(template.sectPrXml).toContain("w:sectPr");
    expect(styleMap.headings[1]).toBe("MyHeading1");
    expect(styleMap.paragraph).toBe("MyNormal");
    expect(styleMap.blockquote).toBe("MyQuote");
    expect(styleMap.table).toBe("MyTable");
  });

  it("throws a clear error for non-zip data", async () => {
    await expect(readTemplatePackage(new TextEncoder().encode("not a zip"), "bad.docx")).rejects.toThrow(/not a readable/i);
  });
});
