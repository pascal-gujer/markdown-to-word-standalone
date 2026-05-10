import type Token from "markdown-it/lib/token.mjs";

export type RichTextSpan = {
  text: string;
  bold?: boolean;
  italic?: boolean;
  code?: boolean;
  link?: string;
};

export type ParagraphBlock = {
  type: "paragraph";
  children: RichTextSpan[];
};

export type HeadingBlock = {
  type: "heading";
  depth: 1 | 2 | 3 | 4 | 5 | 6;
  children: RichTextSpan[];
};

export type CodeBlock = {
  type: "code";
  text: string;
  language?: string;
};

export type HorizontalRuleBlock = {
  type: "hr";
};

export type BlockquoteBlock = {
  type: "blockquote";
  blocks: DocxBlock[];
};

export type ListItem = {
  blocks: DocxBlock[];
};

export type ListBlock = {
  type: "list";
  ordered: boolean;
  start: number;
  items: ListItem[];
};

export type TableCellModel = {
  children: RichTextSpan[];
};

export type TableRowModel = {
  cells: TableCellModel[];
};

export type TableBlock = {
  type: "table";
  header: TableRowModel | null;
  rows: TableRowModel[];
};

export type DocxBlock =
  | ParagraphBlock
  | HeadingBlock
  | CodeBlock
  | HorizontalRuleBlock
  | BlockquoteBlock
  | ListBlock
  | TableBlock;

export type DocxModel = {
  blocks: DocxBlock[];
};

type ParseResult<T> = {
  value: T;
  nextIndex: number;
};

export function markdownTokensToDocxModel(tokens: Token[]): DocxModel {
  return { blocks: parseBlocks(tokens, 0, tokens.length).value };
}

function parseBlocks(tokens: Token[], startIndex: number, endIndex: number): ParseResult<DocxBlock[]> {
  const blocks: DocxBlock[] = [];
  let index = startIndex;

  while (index < endIndex) {
    const token = tokens[index];

    if (token.type === "heading_open") {
      const inline = tokens[index + 1];
      const level = Number(token.tag.slice(1));
      blocks.push({
        type: "heading",
        depth: clampHeadingDepth(level),
        children: inline && inline.type === "inline" ? parseInline(inline.children || []) : [],
      });
      index += 3;
      continue;
    }

    if (token.type === "paragraph_open") {
      const inline = tokens[index + 1];
      blocks.push({
        type: "paragraph",
        children: inline && inline.type === "inline" ? parseInline(inline.children || []) : [],
      });
      index = skipUntilClose(tokens, index + 1, "paragraph_open", "paragraph_close") + 1;
      continue;
    }

    if (token.type === "fence" || token.type === "code_block") {
      blocks.push({
        type: "code",
        text: token.content,
        language: token.info ? token.info.trim().split(/\s+/)[0] : undefined,
      });
      index += 1;
      continue;
    }

    if (token.type === "hr") {
      blocks.push({ type: "hr" });
      index += 1;
      continue;
    }

    if (token.type === "blockquote_open") {
      const closeIndex = findMatchingClose(tokens, index, "blockquote_open", "blockquote_close");
      blocks.push({
        type: "blockquote",
        blocks: parseBlocks(tokens, index + 1, closeIndex).value,
      });
      index = closeIndex + 1;
      continue;
    }

    if (token.type === "bullet_list_open" || token.type === "ordered_list_open") {
      const parsed = parseList(tokens, index);
      blocks.push(parsed.value);
      index = parsed.nextIndex;
      continue;
    }

    if (token.type === "table_open") {
      const parsed = parseTable(tokens, index);
      blocks.push(parsed.value);
      index = parsed.nextIndex;
      continue;
    }

    index += 1;
  }

  return { value: blocks, nextIndex: index };
}

function parseInline(children: Token[]): RichTextSpan[] {
  const spans: RichTextSpan[] = [];
  const state = {
    bold: false,
    italic: false,
    link: undefined as string | undefined,
  };

  for (const child of children) {
    if (child.type === "text") {
      pushSpan(spans, child.content, state);
      continue;
    }

    if (child.type === "code_inline") {
      pushSpan(spans, child.content, { ...state, code: true });
      continue;
    }

    if (child.type === "softbreak" || child.type === "hardbreak") {
      pushSpan(spans, "\n", state);
      continue;
    }

    if (child.type === "strong_open") {
      state.bold = true;
      continue;
    }

    if (child.type === "strong_close") {
      state.bold = false;
      continue;
    }

    if (child.type === "em_open") {
      state.italic = true;
      continue;
    }

    if (child.type === "em_close") {
      state.italic = false;
      continue;
    }

    if (child.type === "link_open") {
      state.link = child.attrGet("href") || undefined;
      continue;
    }

    if (child.type === "link_close") {
      state.link = undefined;
      continue;
    }

    if (child.type === "image") {
      const alt = child.content || child.attrGet("alt") || "image";
      pushSpan(spans, `[${alt}]`, state);
    }
  }

  return spans;
}

function pushSpan(
  spans: RichTextSpan[],
  text: string,
  state: { bold?: boolean; italic?: boolean; code?: boolean; link?: string },
): void {
  if (!text) {
    return;
  }

  const previous = spans[spans.length - 1];
  if (
    previous &&
    previous.bold === Boolean(state.bold) &&
    previous.italic === Boolean(state.italic) &&
    previous.code === Boolean(state.code) &&
    previous.link === state.link
  ) {
    previous.text += text;
    return;
  }

  spans.push({
    text,
    bold: Boolean(state.bold) || undefined,
    italic: Boolean(state.italic) || undefined,
    code: Boolean(state.code) || undefined,
    link: state.link,
  });
}

function parseList(tokens: Token[], startIndex: number): ParseResult<ListBlock> {
  const opening = tokens[startIndex];
  const ordered = opening.type === "ordered_list_open";
  const start = ordered ? Number(opening.attrGet("start") || "1") || 1 : 1;
  const closeIndex = findMatchingClose(tokens, startIndex, opening.type, ordered ? "ordered_list_close" : "bullet_list_close");
  const items: ListItem[] = [];
  let index = startIndex + 1;

  while (index < closeIndex) {
    if (tokens[index].type !== "list_item_open") {
      index += 1;
      continue;
    }

    const itemCloseIndex = findMatchingClose(tokens, index, "list_item_open", "list_item_close");
    items.push({
      blocks: parseBlocks(tokens, index + 1, itemCloseIndex).value,
    });
    index = itemCloseIndex + 1;
  }

  return {
    value: { type: "list", ordered, start, items },
    nextIndex: closeIndex + 1,
  };
}

function parseTable(tokens: Token[], startIndex: number): ParseResult<TableBlock> {
  const closeIndex = findMatchingClose(tokens, startIndex, "table_open", "table_close");
  const rows: TableRowModel[] = [];
  let header: TableRowModel | null = null;
  let index = startIndex + 1;
  let inHead = false;

  while (index < closeIndex) {
    const token = tokens[index];
    if (token.type === "thead_open") {
      inHead = true;
      index += 1;
      continue;
    }
    if (token.type === "thead_close") {
      inHead = false;
      index += 1;
      continue;
    }
    if (token.type === "tr_open") {
      const parsed = parseTableRow(tokens, index);
      if (inHead && !header) {
        header = parsed.value;
      } else {
        rows.push(parsed.value);
      }
      index = parsed.nextIndex;
      continue;
    }
    index += 1;
  }

  return {
    value: { type: "table", header, rows },
    nextIndex: closeIndex + 1,
  };
}

function parseTableRow(tokens: Token[], startIndex: number): ParseResult<TableRowModel> {
  const closeIndex = findMatchingClose(tokens, startIndex, "tr_open", "tr_close");
  const cells: TableCellModel[] = [];
  let index = startIndex + 1;

  while (index < closeIndex) {
    const token = tokens[index];
    if (token.type === "th_open" || token.type === "td_open") {
      const closeType = token.type === "th_open" ? "th_close" : "td_close";
      const cellCloseIndex = findMatchingClose(tokens, index, token.type, closeType);
      const inline = tokens.slice(index + 1, cellCloseIndex).find((candidate) => candidate.type === "inline");
      cells.push({ children: inline ? parseInline(inline.children || []) : [] });
      index = cellCloseIndex + 1;
      continue;
    }
    index += 1;
  }

  return { value: { cells }, nextIndex: closeIndex + 1 };
}

function findMatchingClose(tokens: Token[], openIndex: number, openType: string, closeType: string): number {
  let depth = 0;
  for (let index = openIndex; index < tokens.length; index += 1) {
    if (tokens[index].type === openType) {
      depth += 1;
    }
    if (tokens[index].type === closeType) {
      depth -= 1;
      if (depth === 0) {
        return index;
      }
    }
  }
  return tokens.length - 1;
}

function skipUntilClose(tokens: Token[], openIndex: number, openType: string, closeType: string): number {
  return findMatchingClose(tokens, openIndex - 1, openType, closeType);
}

function clampHeadingDepth(value: number): HeadingBlock["depth"] {
  if (value <= 1) return 1;
  if (value === 2) return 2;
  if (value === 3) return 3;
  if (value === 4) return 4;
  if (value === 5) return 5;
  return 6;
}
