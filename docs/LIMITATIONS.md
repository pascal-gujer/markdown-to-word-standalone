# Limitations

This first version prioritizes reliable offline Markdown to DOCX export over perfect conversion of every Markdown extension.

## Markdown

Supported:

- Headings `h1` through `h6`
- Paragraphs
- Bold and italic text
- Inline code
- Unordered and ordered lists
- Nested lists in common Markdown structures
- Blockquotes
- Fenced code blocks
- Horizontal rules
- Links
- Tables
- Images from a local ZIP bundle when referenced by Markdown and stored as `.png`, `.jpg`, `.jpeg`, or `.gif`
- Unicode text, umlauts, emoji, and code block content

Known limitations:

- Images pasted as remote URLs, data URLs, unsupported formats, or missing ZIP entries are not embedded; alt text is exported in brackets.
- ZIP import selects one Markdown file from the archive. If multiple Markdown files are present, the selected path is shown in the status message.
- GIF files are embedded as image assets. Animation behavior depends on the Word processor.
- Raw HTML is treated as text and is not converted to Word elements.
- Task list checkboxes are exported as normal list text.
- Ordered lists always use generated numbering from 1.
- Markdown extensions such as footnotes, definition lists, math, diagrams, and citations are not implemented.

## DOCX Output

- The output is editable WordprocessingML, not a screenshot or Word-flavored HTML.
- Layout can differ between Word processors.
- Table layout is simple and grid-based.
- Code highlighting is not applied; code uses a monospace style.
- Hyperlinks are written into the document as external relationships, but the app does not open or validate them.
- When a `.docx` or `.dotx` template is used, the main document body is replaced by the converted Markdown. Headers, footers, media, settings, section references, and other package parts are preserved where practical, but arbitrary body placeholders are not yet interpreted.

## Browser APIs

The app depends on browser support for:

- File input
- `File.text()` and `File.arrayBuffer()`
- `Blob`
- Object URLs for downloads
- Modern JavaScript

Most current desktop browsers support these APIs. Some highly locked-down browser policies may disable local file loading or downloads.
