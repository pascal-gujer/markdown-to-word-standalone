export const sampleMarkdown = `# Markdown to Word Offline

This sample preserves Unicode: Grüsse aus Zürich, naïve café, and emoji 🚀.

## Formatting

Normal text with **bold**, *italic*, \`inline code\`, and a [local-friendly link](https://example.invalid/reference).

> A blockquote should use the Quote style when the template provides one.

### Lists

- First unordered item
  - Nested unordered item
- Second unordered item

1. First ordered item
2. Second ordered item
   1. Nested ordered item

### Table

| Feature | Status |
| --- | --- |
| Headings | Supported |
| Tables | Supported |
| Code | Preserved |

---

\`\`\`ts
const message = "Grüsse, Markdown!";
console.log(message);
\`\`\`
`;

export function readTextFile(file: File): Promise<string> {
  return file.text();
}

export function formatInputStats(value: string): string {
  const chars = value.length;
  const words = value.trim() ? value.trim().split(/\s+/).length : 0;
  return `${chars.toLocaleString()} characters · ${words.toLocaleString()} words`;
}

export function ensureDocxFileName(value: string): string {
  const trimmed = value.trim() || "converted-document.docx";
  const withoutUnsafe = trimmed.replace(/[\\/:*?"<>|]+/g, "-");
  return /\.docx$/i.test(withoutUnsafe) ? withoutUnsafe : `${withoutUnsafe}.docx`;
}
