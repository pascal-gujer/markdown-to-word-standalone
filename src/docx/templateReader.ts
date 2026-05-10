import JSZip from "jszip";

export type TemplateStyleType = "paragraph" | "character" | "table" | "numbering" | string;

export type TemplateStyle = {
  styleId: string;
  name: string;
  type: TemplateStyleType;
};

export type TemplateInfo = {
  fileName: string;
  packageBytes?: Uint8Array;
  styles: TemplateStyle[];
  documentXml?: string;
  documentRelsXml?: string;
  contentTypesXml?: string;
  stylesXml?: string;
  themeXml?: string;
  numberingXml?: string;
  sectPrXml?: string;
  hasDocumentPart: boolean;
  hasStylesPart: boolean;
  hasThemePart: boolean;
  hasNumberingPart: boolean;
};

export async function readTemplateFile(file: File): Promise<TemplateInfo> {
  return readTemplatePackage(await file.arrayBuffer(), file.name);
}

export async function readTemplatePackage(data: ArrayBuffer | Uint8Array, fileName = "template.docx"): Promise<TemplateInfo> {
  const packageBytes = data instanceof Uint8Array ? new Uint8Array(data) : new Uint8Array(data);
  let zip: JSZip;
  try {
    zip = await JSZip.loadAsync(packageBytes);
  } catch (_error) {
    throw new Error("error.template_unreadable");
  }

  const contentTypes = await readZipText(zip, "[Content_Types].xml");
  const documentXml = await readZipText(zip, "word/document.xml");
  const documentRelsXml = await readZipText(zip, "word/_rels/document.xml.rels");
  const stylesXml = await readZipText(zip, "word/styles.xml");
  const themeXml = await readZipText(zip, "word/theme/theme1.xml");
  const numberingXml = await readZipText(zip, "word/numbering.xml");

  if (!contentTypes && !documentXml && !stylesXml) {
    throw new Error("error.template_invalid");
  }

  return {
    fileName,
    packageBytes,
    styles: stylesXml ? extractStyles(stylesXml) : [],
    documentXml,
    documentRelsXml,
    contentTypesXml: contentTypes,
    stylesXml,
    themeXml,
    numberingXml,
    sectPrXml: documentXml ? extractLastSectionProperties(documentXml) : undefined,
    hasDocumentPart: Boolean(documentXml),
    hasStylesPart: Boolean(stylesXml),
    hasThemePart: Boolean(themeXml),
    hasNumberingPart: Boolean(numberingXml),
  };
}

async function readZipText(zip: JSZip, path: string): Promise<string | undefined> {
  const entry = zip.file(path);
  return entry ? entry.async("string") : undefined;
}

function extractStyles(stylesXml: string): TemplateStyle[] {
  const styles: TemplateStyle[] = [];
  const stylePattern = /<w:style\b([^>]*)>([\s\S]*?)<\/w:style>/g;
  let match: RegExpExecArray | null;

  while ((match = stylePattern.exec(stylesXml))) {
    const attrs = match[1];
    const body = match[2];
    const styleId = attr(attrs, "w:styleId");
    const type = attr(attrs, "w:type");
    const nameMatch = /<w:name\b[^>]*\bw:val="([^"]*)"/.exec(body);
    if (styleId && type) {
      styles.push({
        styleId: decodeXml(styleId),
        name: nameMatch && nameMatch[1] ? decodeXml(nameMatch[1]) : decodeXml(styleId),
        type,
      });
    }
  }

  return styles;
}

function extractLastSectionProperties(documentXml: string): string | undefined {
  const matches = Array.from(documentXml.matchAll(/<w:sectPr\b[\s\S]*?<\/w:sectPr>/g));
  const last = matches.at(-1);
  if (!last) {
    return undefined;
  }

  const simpleParts = ["pgSz", "pgMar", "cols", "docGrid"]
    .map((tagName) => last[0].match(new RegExp(`<w:${tagName}\\b[^>]*/>`, "g")) || [])
    .flat();

  return simpleParts.length ? `<w:sectPr>${simpleParts.join("")}</w:sectPr>` : undefined;
}

function attr(source: string, name: string): string | undefined {
  const escapedName = name.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
  const match = new RegExp(`${escapedName}="([^"]*)"`).exec(source);
  return match ? match[1] : undefined;
}

function decodeXml(value: string): string {
  return value
    .replace(/&quot;/g, "\"")
    .replace(/&apos;/g, "'")
    .replace(/&lt;/g, "<")
    .replace(/&gt;/g, ">")
    .replace(/&amp;/g, "&");
}
