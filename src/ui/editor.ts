import type { Translate } from "../i18n";

export function sampleMarkdown(t: Translate): string {
  return `${t("sample.markdown").trimEnd()}\n`;
}

export function readTextFile(file: File): Promise<string> {
  return file.text();
}

export function formatInputStats(value: string, locale: string, t: Translate): string {
  const chars = value.length;
  const words = value.trim() ? value.trim().split(/\s+/).length : 0;
  const numberFormat = new Intl.NumberFormat(locale);
  return t("stats.input", {
    chars: numberFormat.format(chars),
    words: numberFormat.format(words),
  });
}

export function ensureDocxFileName(value: string, defaultFileName = "converted-document.docx"): string {
  const trimmed = value.trim() || defaultFileName;
  const withoutUnsafe = trimmed.replace(/[\\/:*?"<>|]+/g, "-");
  return /\.docx$/i.test(withoutUnsafe) ? withoutUnsafe : `${withoutUnsafe}.docx`;
}
