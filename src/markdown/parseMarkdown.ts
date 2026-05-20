import MarkdownIt from "markdown-it";
import type Token from "markdown-it/lib/token.mjs";
import { markdownTokensToDocxModel, type DocxModel } from "./markdownToDocxModel";

export type MarkdownParseResult = {
  html: string;
  tokens: Token[];
  model: DocxModel;
  warnings: MarkdownWarningKey[];
};

export type ParseMarkdownOptions = {
  imageMode?: "embedded" | "text";
};

export type MarkdownWarningKey =
  | "warning.images"
  | "warning.ordered_start"
  | "warning.raw_html_blocked"
  | "warning.raw_html_plain"
  | "warning.task_list";

const parser = new MarkdownIt({
  html: false,
  linkify: true,
  typographer: false,
  breaks: false,
});

export function parseMarkdown(markdown: string, options: ParseMarkdownOptions = {}): MarkdownParseResult {
  const tokens = parser.parse(markdown, {});
  const html = parser.render(markdown);
  return {
    html,
    tokens,
    model: markdownTokensToDocxModel(tokens),
    warnings: detectWarnings(markdown, tokens, options),
  };
}

export function renderMarkdown(markdown: string): string {
  return parser.render(markdown);
}

function detectWarnings(markdown: string, tokens: Token[], options: ParseMarkdownOptions): MarkdownWarningKey[] {
  const warnings = new Set<MarkdownWarningKey>();

  if (/<\s*(script|iframe|style|object|embed|link|meta)\b/i.test(markdown)) {
    warnings.add("warning.raw_html_blocked");
  } else if (/<[a-z][\s\S]*>/i.test(markdown)) {
    warnings.add("warning.raw_html_plain");
  }

  walkTokens(tokens, (token) => {
    if (token.type === "image") {
      if (options.imageMode !== "embedded") {
        warnings.add("warning.images");
      }
    }

    if (token.type === "ordered_list_open") {
      const start = Number(token.attrGet("start") || "1");
      if (Number.isFinite(start) && start !== 1) {
        warnings.add("warning.ordered_start");
      }
    }
  });

  if (/\[[ xX]\]\s+/.test(markdown)) {
    warnings.add("warning.task_list");
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
