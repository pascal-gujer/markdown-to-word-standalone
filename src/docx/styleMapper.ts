import type { TemplateInfo, TemplateStyle } from "./templateReader";

export type StyleSource = "template" | "fallback";

export type StyleDiagnostic = {
  element: string;
  styleId: string;
  styleName: string;
  source: StyleSource;
};

export type StyleMap = {
  paragraph: string;
  headings: Record<1 | 2 | 3 | 4 | 5 | 6, string>;
  blockquote: string;
  codeBlock: string;
  table: string;
  tableText: string;
  hyperlink: string;
  diagnostics: StyleDiagnostic[];
  requiredFallbackStyleIds: string[];
};

type ResolveResult = {
  styleId: string;
  styleName: string;
  source: StyleSource;
};

const headingAliases: Record<1 | 2 | 3 | 4 | 5 | 6, string[]> = {
  1: ["Heading 1", "Ueberschrift 1", "Überschrift 1", "Titre 1"],
  2: ["Heading 2", "Ueberschrift 2", "Überschrift 2", "Titre 2"],
  3: ["Heading 3", "Ueberschrift 3", "Überschrift 3", "Titre 3"],
  4: ["Heading 4", "Ueberschrift 4", "Überschrift 4", "Titre 4"],
  5: ["Heading 5", "Ueberschrift 5", "Überschrift 5", "Titre 5"],
  6: ["Heading 6", "Ueberschrift 6", "Überschrift 6", "Titre 6"],
};

export function buildStyleMap(template?: TemplateInfo | null): StyleMap {
  const styles = template?.styles || [];
  const diagnostics: StyleDiagnostic[] = [];
  const fallbackIds = new Set<string>();

  const resolve = (
    element: string,
    type: string,
    aliases: string[],
    fallbackStyleId: string,
    fallbackName: string,
    styleIdAliases: string[] = [],
  ): ResolveResult => {
    const found = findStyle(styles, type, aliases, [fallbackStyleId, ...styleIdAliases]);
    const result: ResolveResult = found
      ? { styleId: found.styleId, styleName: found.name, source: "template" }
      : { styleId: fallbackStyleId, styleName: fallbackName, source: "fallback" };

    if (!found) {
      fallbackIds.add(fallbackStyleId);
    }

    diagnostics.push({
      element,
      styleId: result.styleId,
      styleName: result.styleName,
      source: result.source,
    });
    return result;
  };

  const normal = resolve("Paragraphs", "paragraph", ["Normal", "Standard"], "Normal", "Normal");
  const headings = {
    1: resolve("#", "paragraph", headingAliases[1], "Heading1", "Heading 1").styleId,
    2: resolve("##", "paragraph", headingAliases[2], "Heading2", "Heading 2").styleId,
    3: resolve("###", "paragraph", headingAliases[3], "Heading3", "Heading 3").styleId,
    4: resolve("####", "paragraph", headingAliases[4], "Heading4", "Heading 4").styleId,
    5: resolve("#####", "paragraph", headingAliases[5], "Heading5", "Heading 5").styleId,
    6: resolve("######", "paragraph", headingAliases[6], "Heading6", "Heading 6").styleId,
  } as Record<1 | 2 | 3 | 4 | 5 | 6, string>;

  const quote = resolve("Blockquotes", "paragraph", ["Quote", "Zitat", "Intense Quote"], "Quote", "Quote");
  const code = resolve(
    "Code blocks",
    "paragraph",
    ["Code", "Code Block", "Source Code", "Quellcode", "Preformatted Text"],
    "CodeBlock",
    "Code Block",
    ["SourceCode"],
  );
  const table = resolve("Tables", "table", ["Table Grid", "Tabellengitternetz", "Tableau Grille"], "TableGrid", "Table Grid");
  const hyperlink = resolve("Links", "character", ["Hyperlink", "Internet Link"], "Hyperlink", "Hyperlink");

  return {
    paragraph: normal.styleId,
    headings,
    blockquote: quote.styleId,
    codeBlock: code.styleId,
    table: table.styleId,
    tableText: normal.styleId,
    hyperlink: hyperlink.styleId,
    diagnostics,
    requiredFallbackStyleIds: Array.from(fallbackIds.add("Hyperlink")),
  };
}

function findStyle(styles: TemplateStyle[], type: string, names: string[], styleIds: string[]): TemplateStyle | undefined {
  const normalizedNames = new Set(names.map(normalizeName));
  const normalizedStyleIds = new Set(styleIds.map(normalizeName));

  return styles.find((style) => {
    if (style.type !== type) {
      return false;
    }
    return normalizedNames.has(normalizeName(style.name)) || normalizedStyleIds.has(normalizeName(style.styleId));
  });
}

function normalizeName(value: string): string {
  return value
    .normalize("NFD")
    .replace(/\p{Diacritic}/gu, "")
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, "");
}
