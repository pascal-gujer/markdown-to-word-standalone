import MarkdownIt from "markdown-it";
import type Token from "markdown-it/lib/token.mjs";
import { markdownTokensToDocxModel, type DocxModel } from "./markdownToDocxModel";

export type MarkdownParseResult = {
  html: string;
  tokens: Token[];
  model: DocxModel;
  warnings: string[];
};

const parser = new MarkdownIt({
  html: false,
  linkify: true,
  typographer: false,
  breaks: false,
});

export function parseMarkdown(markdown: string): MarkdownParseResult {
  const tokens = parser.parse(markdown, {});
  const html = parser.render(markdown);
  return {
    html,
    tokens,
    model: markdownTokensToDocxModel(tokens),
    warnings: detectWarnings(markdown, tokens),
  };
}

export function renderMarkdown(markdown: string): string {
  return parser.render(markdown);
}

function detectWarnings(markdown: string, tokens: Token[]): string[] {
  const warnings = new Set<string>();

  if (/<\s*(script|iframe|style|object|embed|link|meta)\b/i.test(markdown)) {
    warnings.add("Raw HTML is escaped in the preview and ignored as HTML during DOCX export.");
  } else if (/<[a-z][\s\S]*>/i.test(markdown)) {
    warnings.add("Raw HTML is treated as plain text; HTML tags are not converted into Word elements.");
  }

  walkTokens(tokens, (token) => {
    if (token.type === "image") {
      warnings.add("Images are not embedded in this version; image alt text is exported in brackets.");
    }

    if (token.type === "ordered_list_open") {
      const start = Number(token.attrGet("start") || "1");
      if (Number.isFinite(start) && start !== 1) {
        warnings.add("Ordered lists are exported starting at 1 even when Markdown uses another start number.");
      }
    }
  });

  if (/\[[ xX]\]\s+/.test(markdown)) {
    warnings.add("Task list checkboxes are exported as normal list text.");
  }

  return Array.from(warnings);
}

function walkTokens(tokens: Token[], visitor: (token: Token) => void): void {
  for (const token of tokens) {
    visitor(token);
    if (token.children) {
      walkTokens(token.children, visitor);
    }
  }
}
