import JSZip from "jszip";
import { resolveBundleImage, type MarkdownImageAsset, type MarkdownImageBundle } from "../markdown/imageBundle";
import type { DocxBlock, DocxModel, ListBlock, RichTextSpan, TableBlock } from "../markdown/markdownToDocxModel";
import { buildStyleMap, type StyleMap } from "./styleMapper";
import type { TemplateInfo } from "./templateReader";

export type GenerateDocxOptions = {
  images?: MarkdownImageBundle | null;
  template?: TemplateInfo | null;
  title?: string;
};

export const DOCX_MIME = "application/vnd.openxmlformats-officedocument.wordprocessingml.document";

const SCHEME = "http:" + "//";
const PKG_REL_NS = `${SCHEME}schemas.openxmlformats.org/package/2006/relationships`;
const DOC_REL_NS = `${SCHEME}schemas.openxmlformats.org/officeDocument/2006/relationships`;
const CONTENT_TYPES_NS = `${SCHEME}schemas.openxmlformats.org/package/2006/content-types`;
const WORD_NS = `${SCHEME}schemas.openxmlformats.org/wordprocessingml/2006/main`;
const MARKUP_NS = `${SCHEME}schemas.openxmlformats.org/markup-compatibility/2006`;
const DRAWING_NS = `${SCHEME}schemas.openxmlformats.org/drawingml/2006/main`;
const PICTURE_NS = `${SCHEME}schemas.openxmlformats.org/drawingml/2006/picture`;
const WORD_DRAWING_NS = `${SCHEME}schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing`;
const CORE_NS = `${SCHEME}schemas.openxmlformats.org/package/2006/metadata/core-properties`;
const APP_NS = `${SCHEME}schemas.openxmlformats.org/officeDocument/2006/extended-properties`;
const DC_NS = `${SCHEME}purl.org/dc/elements/1.1/`;
const DCTERMS_NS = `${SCHEME}purl.org/dc/terms/`;
const DCMI_NS = `${SCHEME}purl.org/dc/dcmitype/`;
const XSI_NS = `${SCHEME}www.w3.org/2001/XMLSchema-instance`;

type Relationship = {
  id: string;
  type: string;
  target: string;
  targetMode?: "External";
};

type NumberingSetup = {
  xml: string;
  bulletNumId: number;
  orderedNumId: number;
};

type RenderContext = {
  embeddedImages: EmbeddedImage[];
  existingPackagePaths: Set<string>;
  styleMap: StyleMap;
  relationships: Relationship[];
  hyperlinkIds: Map<string, string>;
  imageBundle?: MarkdownImageBundle | null;
  imageIds: Map<string, EmbeddedImage>;
  listNumIds: {
    bullet: number;
    ordered: number;
  };
  nextDocPrId: number;
};

type EmbeddedImage = {
  asset: MarkdownImageAsset;
  docPrId: number;
  relationshipId: string;
  target: string;
};

type RenderOptions = {
  quote?: boolean;
};

export async function generateDocxBlob(model: DocxModel, options: GenerateDocxOptions = {}): Promise<Blob> {
  if (options.template?.packageBytes && options.template.documentXml) {
    return generateDocxFromTemplatePackage(model, options.template, options.title, options.images);
  }

  const styleMap = buildStyleMap(options.template);
  const numbering = createNumberingXml();
  const context: RenderContext = {
    embeddedImages: [],
    existingPackagePaths: new Set(),
    styleMap,
    relationships: [
      { id: "rId1", type: `${DOC_REL_NS}/styles`, target: "styles.xml" },
      { id: "rId2", type: `${DOC_REL_NS}/numbering`, target: "numbering.xml" },
    ],
    hyperlinkIds: new Map(),
    imageBundle: options.images,
    imageIds: new Map(),
    listNumIds: {
      bullet: numbering.bulletNumId,
      ordered: numbering.orderedNumId,
    },
    nextDocPrId: 1,
  };

  if (options.template?.themeXml) {
    context.relationships.push({ id: "rId3", type: `${DOC_REL_NS}/theme`, target: "theme/theme1.xml" });
  }

  const documentXml = createDocumentXml(renderBlocks(model.blocks, context), options.template?.sectPrXml);
  const stylesXml = ensureStylesXml(options.template?.stylesXml, styleMap.requiredFallbackStyleIds);

  const zip = new JSZip();
  zip.file("[Content_Types].xml", createContentTypesXml(Boolean(options.template?.themeXml), context.embeddedImages.length > 0));
  zip.folder("_rels")?.file(".rels", createPackageRelationshipsXml());
  zip.folder("docProps")?.file("core.xml", createCorePropertiesXml(options.title || "Converted Markdown"));
  zip.folder("docProps")?.file("app.xml", createAppPropertiesXml());
  const word = zip.folder("word");
  word?.file("document.xml", documentXml);
  word?.file("styles.xml", stylesXml);
  word?.file("numbering.xml", numbering.xml);
  writeImageFiles(word, context.embeddedImages);
  word?.folder("_rels")?.file("document.xml.rels", createDocumentRelationshipsXml(context.relationships));
  if (options.template?.themeXml) {
    word?.folder("theme")?.file("theme1.xml", options.template.themeXml);
  }

  return zip.generateAsync({
    type: "blob",
    mimeType: DOCX_MIME,
    compression: "DEFLATE",
    compressionOptions: { level: 6 },
  });
}

async function generateDocxFromTemplatePackage(
  model: DocxModel,
  template: TemplateInfo,
  title?: string,
  images?: MarkdownImageBundle | null,
): Promise<Blob> {
  const zip = await JSZip.loadAsync(template.packageBytes!);
  const styleMap = buildStyleMap(template);
  const relationships = parseRelationshipsXml(
    template.documentRelsXml || createDocumentRelationshipsXml([]),
  );
  ensureDocumentRelationship(relationships, "styles", "styles.xml");
  ensureDocumentRelationship(relationships, "numbering", "numbering.xml");
  if (template.themeXml) {
    ensureDocumentRelationship(relationships, "theme", "theme/theme1.xml");
  }

  const numbering = createNumberingXml(template.numberingXml);
  const context: RenderContext = {
    embeddedImages: [],
    existingPackagePaths: new Set(Object.keys(zip.files)),
    styleMap,
    relationships,
    hyperlinkIds: relationshipHyperlinkMap(relationships),
    imageBundle: images,
    imageIds: new Map(),
    listNumIds: {
      bullet: numbering.bulletNumId,
      ordered: numbering.orderedNumId,
    },
    nextDocPrId: nextDrawingDocPrId(template.documentXml || ""),
  };

  const bodyXml = renderBlocks(model.blocks, context);
  const documentXml = replaceTemplateDocumentBody(template.documentXml || "", bodyXml);
  const stylesXml = ensureStylesXml(template.stylesXml, styleMap.requiredFallbackStyleIds);
  const contentTypesXml = ensurePackageContentTypes(template.contentTypesXml || createContentTypesXml(Boolean(template.themeXml), context.embeddedImages.length > 0), {
    hasTheme: Boolean(template.themeXml),
    hasImages: context.embeddedImages.length > 0,
  });

  zip.file("[Content_Types].xml", contentTypesXml);
  zip.file("word/document.xml", documentXml);
  zip.file("word/styles.xml", stylesXml);
  zip.file("word/numbering.xml", numbering.xml);
  writeImageFiles(zip.folder("word"), context.embeddedImages);
  zip.folder("word")?.folder("_rels")?.file("document.xml.rels", createDocumentRelationshipsXml(context.relationships));
  if (template.themeXml) {
    zip.folder("word")?.folder("theme")?.file("theme1.xml", template.themeXml);
  }
  zip.folder("docProps")?.file("core.xml", createCorePropertiesXml(title || "Converted Markdown"));

  return zip.generateAsync({
    type: "blob",
    mimeType: DOCX_MIME,
    compression: "DEFLATE",
    compressionOptions: { level: 6 },
  });
}

function renderBlocks(blocks: DocxBlock[], context: RenderContext, options: RenderOptions = {}): string {
  return blocks.map((block) => renderBlock(block, context, options)).join("");
}

function renderBlock(block: DocxBlock, context: RenderContext, options: RenderOptions): string {
  switch (block.type) {
    case "heading":
      return paragraphXml(block.children, context, {
        styleId: options.quote ? context.styleMap.blockquote : context.styleMap.headings[block.depth],
      });
    case "paragraph":
      return paragraphXml(block.children, context, {
        styleId: options.quote ? context.styleMap.blockquote : context.styleMap.paragraph,
      });
    case "code":
      return codeBlockXml(block.text, context);
    case "hr":
      return horizontalRuleXml();
    case "blockquote":
      return block.blocks.length
        ? renderBlocks(block.blocks, context, { quote: true })
        : paragraphXml([{ text: "" }], context, { styleId: context.styleMap.blockquote });
    case "list":
      return listXml(block, context, 0, options);
    case "table":
      return tableXml(block, context);
  }
}

function listXml(list: ListBlock, context: RenderContext, level: number, options: RenderOptions): string {
  return list.items
    .map((item) => {
      let firstTextBlockNumbered = false;
      return item.blocks
        .map((block) => {
          if (block.type === "paragraph" || block.type === "heading") {
            const children = block.children;
            const numbered = !firstTextBlockNumbered;
            firstTextBlockNumbered = true;
            return paragraphXml(children, context, {
              styleId: options.quote ? context.styleMap.blockquote : context.styleMap.paragraph,
              numbering: numbered ? { level, ordered: list.ordered } : undefined,
              indentLevel: numbered ? undefined : level + 1,
            });
          }

          if (block.type === "list") {
            return listXml(block, context, level + 1, options);
          }

          if (!firstTextBlockNumbered) {
            firstTextBlockNumbered = true;
            return paragraphXml([{ text: "" }], context, {
              styleId: context.styleMap.paragraph,
              numbering: { level, ordered: list.ordered },
            }) + renderBlock(block, context, options);
          }

          return renderBlock(block, context, options);
        })
        .join("");
    })
    .join("");
}

function tableXml(table: TableBlock, context: RenderContext): string {
  const rows = table.header ? [table.header, ...table.rows] : table.rows;
  const columnCount = Math.max(1, ...rows.map((row) => row.cells.length));
  const grid = Array.from({ length: columnCount }, () => `<w:gridCol w:w="${Math.floor(9000 / columnCount)}"/>`).join("");
  const rowXml = rows.map((row, rowIndex) => {
    const cells = row.cells.map((cell) => tableCellXml(cell.children, context, rowIndex === 0 && Boolean(table.header))).join("");
    return `<w:tr>${cells}</w:tr>`;
  }).join("");

  return [
    "<w:tbl>",
    "<w:tblPr>",
    `<w:tblStyle w:val="${attr(context.styleMap.table)}"/>`,
    '<w:tblW w:w="5000" w:type="pct"/>',
    '<w:tblBorders><w:top w:val="single" w:sz="4" w:space="0" w:color="D8DEE8"/><w:left w:val="single" w:sz="4" w:space="0" w:color="D8DEE8"/><w:bottom w:val="single" w:sz="4" w:space="0" w:color="D8DEE8"/><w:right w:val="single" w:sz="4" w:space="0" w:color="D8DEE8"/><w:insideH w:val="single" w:sz="4" w:space="0" w:color="D8DEE8"/><w:insideV w:val="single" w:sz="4" w:space="0" w:color="D8DEE8"/></w:tblBorders>',
    "</w:tblPr>",
    `<w:tblGrid>${grid}</w:tblGrid>`,
    rowXml,
    "</w:tbl>",
  ].join("");
}

function tableCellXml(children: RichTextSpan[], context: RenderContext, header: boolean): string {
  const cellPr = header
    ? '<w:tcPr><w:shd w:val="clear" w:color="auto" w:fill="F0F4F3"/><w:tcMar><w:top w:w="90" w:type="dxa"/><w:left w:w="120" w:type="dxa"/><w:bottom w:w="90" w:type="dxa"/><w:right w:w="120" w:type="dxa"/></w:tcMar></w:tcPr>'
    : '<w:tcPr><w:tcMar><w:top w:w="90" w:type="dxa"/><w:left w:w="120" w:type="dxa"/><w:bottom w:w="90" w:type="dxa"/><w:right w:w="120" w:type="dxa"/></w:tcMar></w:tcPr>';
  const cellChildren = header ? children.map((span) => ({ ...span, bold: true })) : children;
  return `<w:tc>${cellPr}${paragraphXml(cellChildren, context, { styleId: context.styleMap.tableText })}</w:tc>`;
}

function paragraphXml(
  children: RichTextSpan[],
  context: RenderContext,
  options: {
    styleId?: string;
    numbering?: { level: number; ordered: boolean };
    indentLevel?: number;
  } = {},
): string {
  const pPr = paragraphPropertiesXml(options, context);
  const runs = children.length ? children.map((span) => spanXml(span, context)).join("") : "<w:r/>";
  return `<w:p>${pPr}${runs}</w:p>`;
}

function paragraphPropertiesXml(options: {
  styleId?: string;
  numbering?: { level: number; ordered: boolean };
  indentLevel?: number;
}, context: RenderContext): string {
  const parts: string[] = [];
  if (options.styleId) {
    parts.push(`<w:pStyle w:val="${attr(options.styleId)}"/>`);
  }
  if (options.numbering) {
    const level = Math.max(0, Math.min(8, options.numbering.level));
    const numId = options.numbering.ordered ? context.listNumIds.ordered : context.listNumIds.bullet;
    parts.push(`<w:numPr><w:ilvl w:val="${level}"/><w:numId w:val="${numId}"/></w:numPr>`);
  }
  if (typeof options.indentLevel === "number") {
    parts.push(`<w:ind w:left="${Math.max(0, options.indentLevel) * 720}"/>`);
  }
  return parts.length ? `<w:pPr>${parts.join("")}</w:pPr>` : "";
}

function spanXml(span: RichTextSpan, context: RenderContext): string {
  if (span.image) {
    const image = embeddedImageForSpan(span, context);
    return image ? imageRunXml(image, span.image.alt) : runXml({ ...span, image: undefined, text: `[${span.image.alt || "image"}]` }, context);
  }

  const run = runXml(span, context);
  if (!span.link) {
    return run;
  }

  const linkId = hyperlinkRelationshipId(span.link, context);
  return `<w:hyperlink r:id="${attr(linkId)}" w:history="1">${runXml({ ...span, link: undefined }, context, true)}</w:hyperlink>`;
}

function embeddedImageForSpan(span: RichTextSpan, context: RenderContext): EmbeddedImage | undefined {
  if (!span.image) {
    return undefined;
  }
  const asset = resolveBundleImage(context.imageBundle, span.image.src);
  if (!asset) {
    return undefined;
  }

  const existing = context.imageIds.get(asset.path);
  if (existing) {
    return existing;
  }

  const target = nextImageTarget(asset, context);
  const embedded: EmbeddedImage = {
    asset,
    docPrId: context.nextDocPrId,
    relationshipId: nextRelationshipId(context.relationships),
    target,
  };
  context.nextDocPrId += 1;
  context.imageIds.set(asset.path, embedded);
  context.embeddedImages.push(embedded);
  context.relationships.push({
    id: embedded.relationshipId,
    type: `${DOC_REL_NS}/image`,
    target,
  });
  return embedded;
}

function imageRunXml(image: EmbeddedImage, alt: string): string {
  const size = imageSizeEmu(image.asset);
  const descr = attr(alt || image.asset.fileName);
  const name = attr(image.asset.fileName);
  return [
    "<w:r><w:drawing>",
    '<wp:inline distT="0" distB="0" distL="0" distR="0">',
    `<wp:extent cx="${size.cx}" cy="${size.cy}"/>`,
    '<wp:effectExtent l="0" t="0" r="0" b="0"/>',
    `<wp:docPr id="${image.docPrId}" name="${name}" descr="${descr}"/>`,
    '<wp:cNvGraphicFramePr><a:graphicFrameLocks noChangeAspect="1"/></wp:cNvGraphicFramePr>',
    '<a:graphic>',
    `<a:graphicData uri="${PICTURE_NS}">`,
    "<pic:pic>",
    `<pic:nvPicPr><pic:cNvPr id="0" name="${name}" descr="${descr}"/><pic:cNvPicPr/></pic:nvPicPr>`,
    `<pic:blipFill><a:blip r:embed="${attr(image.relationshipId)}"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>`,
    `<pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${size.cx}" cy="${size.cy}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr>`,
    "</pic:pic>",
    "</a:graphicData>",
    "</a:graphic>",
    "</wp:inline>",
    "</w:drawing></w:r>",
  ].join("");
}

function imageSizeEmu(asset: MarkdownImageAsset): { cx: number; cy: number } {
  const emuPerPixel = 9525;
  const maxWidth = 5_760_000;
  const naturalWidth = Math.max(1, asset.widthPx) * emuPerPixel;
  const naturalHeight = Math.max(1, asset.heightPx) * emuPerPixel;
  if (naturalWidth <= maxWidth) {
    return { cx: naturalWidth, cy: naturalHeight };
  }
  return {
    cx: maxWidth,
    cy: Math.max(1, Math.round(naturalHeight * (maxWidth / naturalWidth))),
  };
}

function nextImageTarget(asset: MarkdownImageAsset, context: RenderContext): string {
  let index = context.embeddedImages.length + 1;
  let target = "";
  do {
    target = `media/markdown-image-${index}.${asset.extension === "jpeg" ? "jpg" : asset.extension}`;
    index += 1;
  } while (context.existingPackagePaths.has(`word/${target}`));
  context.existingPackagePaths.add(`word/${target}`);
  return target;
}

function writeImageFiles(word: JSZip | null, images: EmbeddedImage[]): void {
  for (const image of images) {
    word?.file(image.target, image.asset.data);
  }
}

function runXml(span: RichTextSpan, context: RenderContext, forceHyperlinkStyle = false): string {
  const props: string[] = [];
  if (span.bold) props.push("<w:b/>");
  if (span.italic) props.push("<w:i/>");
  if (span.code) {
    props.push('<w:rFonts w:ascii="Consolas" w:hAnsi="Consolas" w:cs="Consolas"/>');
    props.push('<w:shd w:val="clear" w:color="auto" w:fill="EEF2F7"/>');
  }
  if (forceHyperlinkStyle) {
    props.push(`<w:rStyle w:val="${attr(context.styleMap.hyperlink)}"/>`);
  }

  const rPr = props.length ? `<w:rPr>${props.join("")}</w:rPr>` : "";
  return `<w:r>${rPr}${textXml(span.text)}</w:r>`;
}

function textXml(text: string): string {
  const clean = stripIllegalXmlChars(text);
  const parts = clean.split("\n");
  return parts
    .map((part, index) => {
      const prefix = index === 0 ? "" : "<w:br/>";
      return `${prefix}<w:t xml:space="preserve">${escapeXml(part)}</w:t>`;
    })
    .join("");
}

function codeBlockXml(text: string, context: RenderContext): string {
  const normalized = text.endsWith("\n") ? text.slice(0, -1) : text;
  return paragraphXml([{ text: normalized || " ", code: true }], context, { styleId: context.styleMap.codeBlock });
}

function horizontalRuleXml(): string {
  return '<w:p><w:pPr><w:pBdr><w:bottom w:val="single" w:sz="8" w:space="4" w:color="D8DEE8"/></w:pBdr></w:pPr></w:p>';
}

function hyperlinkRelationshipId(url: string, context: RenderContext): string {
  const target = stripIllegalXmlChars(url.trim());
  const existing = context.hyperlinkIds.get(target);
  if (existing) {
    return existing;
  }
  const id = nextRelationshipId(context.relationships);
  context.relationships.push({
    id,
    type: `${DOC_REL_NS}/hyperlink`,
    target,
    targetMode: "External",
  });
  context.hyperlinkIds.set(target, id);
  return id;
}

function createDocumentXml(bodyXml: string, templateSectPr?: string): string {
  const sectPr = templateSectPr || '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="708" w:footer="708" w:gutter="0"/></w:sectPr>';
  return xmlDeclaration() +
    `<w:document xmlns:w="${WORD_NS}" xmlns:r="${DOC_REL_NS}" xmlns:mc="${MARKUP_NS}" xmlns:wp="${WORD_DRAWING_NS}" xmlns:a="${DRAWING_NS}" xmlns:pic="${PICTURE_NS}" mc:Ignorable="">` +
    `<w:body>${bodyXml || '<w:p/>'}${sectPr}</w:body>` +
    "</w:document>";
}

function createContentTypesXml(includeTheme: boolean, includeImages = false): string {
  const themeOverride = includeTheme
    ? '<Override PartName="/word/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/>'
    : "";
  const imageDefaults = includeImages ? imageContentTypeDefaults() : "";

  return xmlDeclaration() +
    `<Types xmlns="${CONTENT_TYPES_NS}">` +
    '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
    '<Default Extension="xml" ContentType="application/xml"/>' +
    imageDefaults +
    '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>' +
    '<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>' +
    '<Override PartName="/word/numbering.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/>' +
    themeOverride +
    '<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>' +
    '<Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>' +
    "</Types>";
}

function createPackageRelationshipsXml(): string {
  return xmlDeclaration() +
    `<Relationships xmlns="${PKG_REL_NS}">` +
    `<Relationship Id="rId1" Type="${DOC_REL_NS}/officeDocument" Target="word/document.xml"/>` +
    `<Relationship Id="rId2" Type="${PKG_REL_NS}/metadata/core-properties" Target="docProps/core.xml"/>` +
    `<Relationship Id="rId3" Type="${DOC_REL_NS}/extended-properties" Target="docProps/app.xml"/>` +
    "</Relationships>";
}

function createDocumentRelationshipsXml(relationships: Relationship[]): string {
  return xmlDeclaration() +
    `<Relationships xmlns="${PKG_REL_NS}">` +
    relationships
      .map((rel) => {
        const targetMode = rel.targetMode ? ` TargetMode="${rel.targetMode}"` : "";
        return `<Relationship Id="${attr(rel.id)}" Type="${attr(rel.type)}" Target="${attr(rel.target)}"${targetMode}/>`;
      })
      .join("") +
    "</Relationships>";
}

function parseRelationshipsXml(xml: string): Relationship[] {
  const relationships: Relationship[] = [];
  const pattern = /<Relationship\b([^>]*)\/?>/g;
  let match: RegExpExecArray | null;
  while ((match = pattern.exec(xml))) {
    const attrs = match[1] || "";
    const id = readAttr(attrs, "Id");
    const type = readAttr(attrs, "Type");
    const target = readAttr(attrs, "Target");
    const targetMode = readAttr(attrs, "TargetMode");
    if (id && type && target) {
      relationships.push({
        id: decodeXmlAttr(id),
        type: decodeXmlAttr(type),
        target: decodeXmlAttr(target),
        targetMode: targetMode === "External" ? "External" : undefined,
      });
    }
  }
  return relationships;
}

function ensureDocumentRelationship(relationships: Relationship[], relKind: string, target: string): void {
  const type = `${DOC_REL_NS}/${relKind}`;
  const found = relationships.some((relationship) => relationship.type === type && relationship.target === target);
  if (!found) {
    relationships.push({ id: nextRelationshipId(relationships), type, target });
  }
}

function relationshipHyperlinkMap(relationships: Relationship[]): Map<string, string> {
  const hyperlinks = new Map<string, string>();
  for (const relationship of relationships) {
    if (relationship.type === `${DOC_REL_NS}/hyperlink` && relationship.targetMode === "External") {
      hyperlinks.set(relationship.target, relationship.id);
    }
  }
  return hyperlinks;
}

function nextRelationshipId(relationships: Relationship[]): string {
  const used = new Set(relationships.map((relationship) => relationship.id));
  let next = 1;
  for (const relationship of relationships) {
    const match = /^rId(\d+)$/.exec(relationship.id);
    if (match) {
      next = Math.max(next, Number(match[1]) + 1);
    }
  }
  while (used.has(`rId${next}`)) {
    next += 1;
  }
  return `rId${next}`;
}

function ensurePackageContentTypes(xml: string, options: { hasImages?: boolean; hasTheme: boolean }): string {
  let output = xml && /<Types\b/.test(xml) ? xml : createContentTypesXml(options.hasTheme);
  output = upsertContentTypeOverride(
    output,
    "/word/document.xml",
    "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml",
  );
  output = upsertContentTypeOverride(
    output,
    "/word/styles.xml",
    "application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml",
  );
  output = upsertContentTypeOverride(
    output,
    "/word/numbering.xml",
    "application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml",
  );
  if (options.hasTheme) {
    output = upsertContentTypeOverride(
      output,
      "/word/theme/theme1.xml",
      "application/vnd.openxmlformats-officedocument.theme+xml",
    );
  }
  if (options.hasImages) {
    output = upsertDefaultContentType(output, "gif", "image/gif");
    output = upsertDefaultContentType(output, "jpeg", "image/jpeg");
    output = upsertDefaultContentType(output, "jpg", "image/jpeg");
    output = upsertDefaultContentType(output, "png", "image/png");
  }
  return output;
}

function imageContentTypeDefaults(): string {
  return '<Default Extension="gif" ContentType="image/gif"/>' +
    '<Default Extension="jpeg" ContentType="image/jpeg"/>' +
    '<Default Extension="jpg" ContentType="image/jpeg"/>' +
    '<Default Extension="png" ContentType="image/png"/>';
}

function upsertDefaultContentType(xml: string, extension: string, contentType: string): string {
  const defaultPattern = new RegExp(`<Default\\b[^>]*Extension="${escapeRegExp(extension)}"[^>]*/>`);
  const item = `<Default Extension="${extension}" ContentType="${contentType}"/>`;
  if (defaultPattern.test(xml)) {
    return xml.replace(defaultPattern, item);
  }
  const firstOverrideIndex = xml.search(/<Override\b/);
  if (firstOverrideIndex >= 0) {
    return `${xml.slice(0, firstOverrideIndex)}${item}${xml.slice(firstOverrideIndex)}`;
  }
  return xml.replace("</Types>", `${item}</Types>`);
}

function upsertContentTypeOverride(xml: string, partName: string, contentType: string): string {
  const overridePattern = new RegExp(`<Override\\b[^>]*PartName="${escapeRegExp(partName)}"[^>]*/>`);
  const override = `<Override PartName="${partName}" ContentType="${contentType}"/>`;
  if (overridePattern.test(xml)) {
    return xml.replace(overridePattern, override);
  }
  return xml.replace("</Types>", `${override}</Types>`);
}

function replaceTemplateDocumentBody(templateDocumentXml: string, bodyXml: string): string {
  const match = /(<w:body\b[^>]*>)([\s\S]*)(<\/w:body>)/.exec(templateDocumentXml);
  if (!match) {
    return createDocumentXml(bodyXml);
  }

  const bodyOpen = match[1];
  const bodyContent = match[2] || "";
  const bodyClose = match[3];
  const sectPr = extractLastSectionPropertiesXml(bodyContent) ||
    '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="708" w:footer="708" w:gutter="0"/></w:sectPr>';

  return ensureDocumentNamespaces(templateDocumentXml.slice(0, match.index) +
    bodyOpen +
    (bodyXml || "<w:p/>") +
    sectPr +
    bodyClose +
    templateDocumentXml.slice(match.index + match[0].length));
}

function extractLastSectionPropertiesXml(bodyContent: string): string | undefined {
  const matches = Array.from(bodyContent.matchAll(/<w:sectPr\b[\s\S]*?<\/w:sectPr>/g));
  return matches.at(-1)?.[0];
}

function ensureDocumentNamespaces(documentXml: string): string {
  const match = /<w:document\b([^>]*)>/.exec(documentXml);
  if (!match) {
    return documentXml;
  }

  const attrs = match[1];
  const additions = [
    attrs.includes("xmlns:wp=") ? "" : ` xmlns:wp="${WORD_DRAWING_NS}"`,
    attrs.includes("xmlns:a=") ? "" : ` xmlns:a="${DRAWING_NS}"`,
    attrs.includes("xmlns:pic=") ? "" : ` xmlns:pic="${PICTURE_NS}"`,
    attrs.includes("xmlns:r=") ? "" : ` xmlns:r="${DOC_REL_NS}"`,
  ].join("");

  return additions
    ? `${documentXml.slice(0, match.index)}<w:document${attrs}${additions}>${documentXml.slice(match.index + match[0].length)}`
    : documentXml;
}

function nextDrawingDocPrId(documentXml: string): number {
  let max = 0;
  let match: RegExpExecArray | null;
  const pattern = /<wp:docPr\b[^>]*\bid="(\d+)"/g;
  while ((match = pattern.exec(documentXml))) {
    max = Math.max(max, Number(match[1]) || 0);
  }
  return max + 1;
}

function createCorePropertiesXml(title: string): string {
  const now = new Date().toISOString();
  return xmlDeclaration() +
    `<cp:coreProperties xmlns:cp="${CORE_NS}" xmlns:dc="${DC_NS}" xmlns:dcterms="${DCTERMS_NS}" xmlns:dcmitype="${DCMI_NS}" xmlns:xsi="${XSI_NS}">` +
    `<dc:title>${escapeXml(title)}</dc:title>` +
    "<dc:creator>Markdown to Word Offline</dc:creator>" +
    "<cp:lastModifiedBy>Markdown to Word Offline</cp:lastModifiedBy>" +
    `<dcterms:created xsi:type="dcterms:W3CDTF">${now}</dcterms:created>` +
    `<dcterms:modified xsi:type="dcterms:W3CDTF">${now}</dcterms:modified>` +
    "</cp:coreProperties>";
}

function createAppPropertiesXml(): string {
  return xmlDeclaration() +
    `<Properties xmlns="${APP_NS}">` +
    "<Application>Markdown to Word Offline</Application>" +
    "<DocSecurity>0</DocSecurity>" +
    "<ScaleCrop>false</ScaleCrop>" +
    "</Properties>";
}

function ensureStylesXml(templateStylesXml: string | undefined, requiredFallbackStyleIds: string[]): string {
  const base = templateStylesXml && /<w:styles\b/.test(templateStylesXml)
    ? templateStylesXml
    : createFallbackStylesXml([]);
  const missing = requiredFallbackStyleIds.filter((styleId) => !new RegExp(`\\bw:styleId="${escapeRegExp(styleId)}"`).test(base));

  if (!missing.length) {
    return base;
  }

  const additions = missing.map(fallbackStyleSnippet).filter(Boolean).join("");
  return base.replace("</w:styles>", `${additions}</w:styles>`);
}

function createFallbackStylesXml(_requiredStyleIds: string[]): string {
  const all = [
    "Normal",
    "Heading1",
    "Heading2",
    "Heading3",
    "Heading4",
    "Heading5",
    "Heading6",
    "Quote",
    "CodeBlock",
    "Hyperlink",
    "TableGrid",
    "ListParagraph",
  ].map(fallbackStyleSnippet).join("");
  return xmlDeclaration() + `<w:styles xmlns:w="${WORD_NS}">${all}</w:styles>`;
}

function fallbackStyleSnippet(styleId: string): string {
  const snippets: Record<string, string> = {
    Normal: '<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/><w:qFormat/><w:pPr><w:spacing w:after="160" w:line="276" w:lineRule="auto"/></w:pPr><w:rPr><w:rFonts w:ascii="Aptos" w:hAnsi="Aptos" w:cs="Aptos"/><w:sz w:val="22"/></w:rPr></w:style>',
    Heading1: headingStyle("Heading1", "Heading 1", 36, 0),
    Heading2: headingStyle("Heading2", "Heading 2", 30, 1),
    Heading3: headingStyle("Heading3", "Heading 3", 26, 2),
    Heading4: headingStyle("Heading4", "Heading 4", 24, 3),
    Heading5: headingStyle("Heading5", "Heading 5", 22, 4),
    Heading6: headingStyle("Heading6", "Heading 6", 20, 5),
    Quote: '<w:style w:type="paragraph" w:styleId="Quote"><w:name w:val="Quote"/><w:basedOn w:val="Normal"/><w:qFormat/><w:pPr><w:pBdr><w:left w:val="single" w:sz="16" w:space="8" w:color="0F766E"/></w:pBdr><w:spacing w:before="120" w:after="120"/><w:ind w:left="360"/></w:pPr><w:rPr><w:i/><w:color w:val="374151"/></w:rPr></w:style>',
    CodeBlock: '<w:style w:type="paragraph" w:styleId="CodeBlock"><w:name w:val="Code Block"/><w:basedOn w:val="Normal"/><w:qFormat/><w:pPr><w:shd w:val="clear" w:color="auto" w:fill="F3F4F6"/><w:spacing w:before="120" w:after="120"/></w:pPr><w:rPr><w:rFonts w:ascii="Consolas" w:hAnsi="Consolas" w:cs="Consolas"/><w:sz w:val="20"/></w:rPr></w:style>',
    Hyperlink: '<w:style w:type="character" w:styleId="Hyperlink"><w:name w:val="Hyperlink"/><w:uiPriority w:val="99"/><w:rPr><w:color w:val="0563C1"/><w:u w:val="single"/></w:rPr></w:style>',
    TableGrid: '<w:style w:type="table" w:styleId="TableGrid"><w:name w:val="Table Grid"/><w:uiPriority w:val="59"/><w:qFormat/><w:tblPr><w:tblBorders><w:top w:val="single" w:sz="4" w:space="0" w:color="D8DEE8"/><w:left w:val="single" w:sz="4" w:space="0" w:color="D8DEE8"/><w:bottom w:val="single" w:sz="4" w:space="0" w:color="D8DEE8"/><w:right w:val="single" w:sz="4" w:space="0" w:color="D8DEE8"/><w:insideH w:val="single" w:sz="4" w:space="0" w:color="D8DEE8"/><w:insideV w:val="single" w:sz="4" w:space="0" w:color="D8DEE8"/></w:tblBorders></w:tblPr></w:style>',
    ListParagraph: '<w:style w:type="paragraph" w:styleId="ListParagraph"><w:name w:val="List Paragraph"/><w:basedOn w:val="Normal"/><w:uiPriority w:val="34"/><w:qFormat/><w:pPr><w:ind w:left="720"/></w:pPr></w:style>',
  };
  return snippets[styleId] || "";
}

function headingStyle(styleId: string, name: string, size: number, outlineLevel: number): string {
  return `<w:style w:type="paragraph" w:styleId="${styleId}"><w:name w:val="${name}"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:qFormat/><w:pPr><w:keepNext/><w:spacing w:before="360" w:after="160"/><w:outlineLvl w:val="${outlineLevel}"/></w:pPr><w:rPr><w:b/><w:color w:val="111827"/><w:sz w:val="${size}"/></w:rPr></w:style>`;
}

function createNumberingXml(templateNumberingXml?: string): NumberingSetup {
  if (templateNumberingXml && /<w:numbering\b/.test(templateNumberingXml)) {
    const abstractNumStart = nextNumberingId(templateNumberingXml, "abstractNumId");
    const numStart = nextNumberingId(templateNumberingXml, "numId");
    const bulletAbstractNumId = abstractNumStart;
    const orderedAbstractNumId = abstractNumStart + 1;
    const bulletNumId = numStart;
    const orderedNumId = numStart + 1;
    const additions = numberingDefinitionsXml(bulletAbstractNumId, orderedAbstractNumId, bulletNumId, orderedNumId);
    return {
      xml: insertNumberingDefinitions(templateNumberingXml, additions),
      bulletNumId,
      orderedNumId,
    };
  }

  const bulletAbstractNumId = 1;
  const orderedAbstractNumId = 2;
  const bulletNumId = 1;
  const orderedNumId = 2;
  const definitions = numberingDefinitionsXml(bulletAbstractNumId, orderedAbstractNumId, bulletNumId, orderedNumId);
  return {
    xml: xmlDeclaration() +
      `<w:numbering xmlns:w="${WORD_NS}">` +
      definitions.abstractNums +
      definitions.nums +
      "</w:numbering>",
    bulletNumId,
    orderedNumId,
  };
}

function numberingDefinitionsXml(
  bulletAbstractNumId: number,
  orderedAbstractNumId: number,
  bulletNumId: number,
  orderedNumId: number,
): { abstractNums: string; nums: string } {
  const bulletLevels = Array.from({ length: 9 }, (_, level) => {
    const left = (level + 1) * 720;
    const bullets = ["•", "◦", "▪"];
    return `<w:lvl w:ilvl="${level}"><w:start w:val="1"/><w:numFmt w:val="bullet"/><w:lvlText w:val="${bullets[level % bullets.length]}"/><w:lvlJc w:val="left"/><w:pPr><w:ind w:left="${left}" w:hanging="360"/></w:pPr></w:lvl>`;
  }).join("");
  const orderedLevels = Array.from({ length: 9 }, (_, level) => {
    const left = (level + 1) * 720;
    return `<w:lvl w:ilvl="${level}"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%${level + 1}."/><w:lvlJc w:val="left"/><w:pPr><w:ind w:left="${left}" w:hanging="360"/></w:pPr></w:lvl>`;
  }).join("");

  return {
    abstractNums: `<w:abstractNum w:abstractNumId="${bulletAbstractNumId}">${bulletLevels}</w:abstractNum>` +
      `<w:abstractNum w:abstractNumId="${orderedAbstractNumId}">${orderedLevels}</w:abstractNum>`,
    nums: `<w:num w:numId="${bulletNumId}"><w:abstractNumId w:val="${bulletAbstractNumId}"/></w:num>` +
      `<w:num w:numId="${orderedNumId}"><w:abstractNumId w:val="${orderedAbstractNumId}"/></w:num>`,
  };
}

function insertNumberingDefinitions(numberingXml: string, additions: { abstractNums: string; nums: string }): string {
  const firstNumIndex = numberingXml.search(/<w:num\b/);
  const withAbstractNums = firstNumIndex >= 0
    ? `${numberingXml.slice(0, firstNumIndex)}${additions.abstractNums}${numberingXml.slice(firstNumIndex)}`
    : numberingXml.replace("</w:numbering>", `${additions.abstractNums}</w:numbering>`);

  return withAbstractNums.replace("</w:numbering>", `${additions.nums}</w:numbering>`);
}

function nextNumberingId(numberingXml: string, attrName: "abstractNumId" | "numId"): number {
  const pattern = new RegExp(`\\bw:${attrName}="(\\d+)"`, "g");
  let max = 0;
  let match: RegExpExecArray | null;
  while ((match = pattern.exec(numberingXml))) {
    max = Math.max(max, Number(match[1]) || 0);
  }
  return max + 1;
}

function xmlDeclaration(): string {
  return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';
}

function escapeXml(value: string): string {
  return value
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;");
}

function attr(value: string): string {
  return escapeXml(value).replace(/"/g, "&quot;");
}

function readAttr(source: string, name: string): string | undefined {
  const match = new RegExp(`\\b${escapeRegExp(name)}="([^"]*)"`).exec(source);
  return match ? match[1] : undefined;
}

function decodeXmlAttr(value: string): string {
  return value
    .replace(/&quot;/g, "\"")
    .replace(/&apos;/g, "'")
    .replace(/&lt;/g, "<")
    .replace(/&gt;/g, ">")
    .replace(/&amp;/g, "&");
}

function stripIllegalXmlChars(value: string): string {
  return value.replace(/[\u0000-\u0008\u000B\u000C\u000E-\u001F]/g, "");
}

function escapeRegExp(value: string): string {
  return value.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
}
