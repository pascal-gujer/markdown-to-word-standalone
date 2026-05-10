# Template Support

Template support is intentionally honest and incremental. The app uses `.docx` or `.dotx` files as OpenXML ZIP packages. When a complete package is loaded, the exporter starts from that package and replaces the main document body with the converted Markdown.

## What Is Read

When a template is loaded, the app attempts to read:

- `word/styles.xml`
- `word/theme/theme1.xml`
- `word/numbering.xml`
- `word/document.xml` for the last `w:sectPr` section settings block

The diagnostics panel lists detected styles and the Markdown-to-Word style mapping.

## Style Mapping

The mapper searches paragraph, character, and table styles by common style names and style IDs.

Default mappings:

- Markdown headings map to Heading 1 through Heading 6.
- Paragraphs map to Normal.
- Blockquotes map to Quote.
- Code blocks map to Code, Source Code, or Code Block.
- Links map to Hyperlink.
- Tables map to Table Grid.

Localized Word aliases such as `Überschrift 1`, `Standard`, `Zitat`, and `Quellcode` are recognized for common templates.

## What Is Preserved

The generated document can preserve or reuse:

- Existing style definitions from the template.
- The template theme part when present.
- Header and footer parts referenced by section settings.
- Media used by preserved parts, such as header and footer graphics.
- Settings, web settings, font tables, custom XML, footnotes, and endnotes.
- Fields in preserved parts, such as page numbering in footers.
- Page and section settings from the template document part.

The main document body is replaced by converted Markdown. The section properties at the end of the body are preserved so headers, footers, margins, and page geometry continue to apply.

Generated fallback styles are added when expected styles are missing.

## What Is Not Preserved

The app does not attempt to reproduce or execute:

- Macros
- Embedded fonts
- Placeholder logic
- Mail merge fields
- Complex content controls
- Arbitrary body placeholders without an explicit future mapping feature
- Complex page layouts inside the body content being replaced
- Watermarks
- Document variables
- Template numbering behavior
- Tracked changes or comments

The app appends its own list numbering definitions for predictable Markdown list export while keeping existing numbering definitions in the template package.

## `.dotx` Handling

`.dotx` files are handled like `.docx` files because both are ZIP-based OpenXML packages. The app does not execute template logic; it only reads useful XML parts.
