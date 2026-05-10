# Markdown to Word Offline

Markdown to Word Offline is a browser-only tool that converts Markdown into a downloadable `.docx` file. It is built for restricted and offline machines: no server, no telemetry, no CDN, no external fonts, and no runtime network access.

## Build

```sh
npm install
npm run build
```

The build writes one self-contained file at `dist/md-to-docx.html`.

Development server:

```sh
npm run dev
```

Tests:

```sh
npm run test
```

`npm run build` validates that the final HTML has no raw `http://` or `https://` text, CDN references, external scripts, external stylesheets, or module imports.

## Offline Use

After building, open `dist/md-to-docx.html` on the offline machine. It is intended to work directly from `file://` in browsers that permit local file selection, Blob downloads, and modern JavaScript APIs.

## Workflow

1. Paste Markdown or load a local `.md`, `.markdown`, or `.txt` file.
2. Optionally load a `.docx` or `.dotx` style template.
3. Choose the output file name and export the `.docx`.

Files are processed locally in your browser and are not uploaded.

## Template Support

The app reads `.docx` and `.dotx` files as OpenXML ZIP packages. When a complete Word package is loaded, export uses that package as the base and replaces the main document body with the converted Markdown. This preserves practical template parts such as headers, footers, media, theme, settings, footnotes, endnotes, fields in preserved parts, custom XML, and section references.

It also inspects `word/styles.xml`, `word/theme/theme1.xml`, `word/numbering.xml`, and section settings from `word/document.xml`.

Markdown structures are mapped to Word style names where practical:

- Headings map to Heading 1 through Heading 6.
- Paragraphs map to Normal.
- Blockquotes map to Quote.
- Code blocks map to Code, Source Code, or a generated Code Block fallback.
- Tables map to Table Grid or a generated fallback table style.

This does not mean perfect Word template automation. The goal is to keep the original package scaffolding intact and write the Markdown content into it with mapped Word styles.

## Why `.doc` Is Not Supported

Legacy `.doc` files are a binary Word format. They require a different parser/writer and are not suitable for a safe browser-only offline converter. `.docx` and `.dotx` are ZIP-based OpenXML packages, so they can be generated and inspected without native binaries.

## DOCX Generation vs. Perfect Word Rendering

The app generates editable WordprocessingML. Word, LibreOffice, Pages, and other editors may render details differently. The goal is a reliable `.docx`, not pixel-perfect reproduction of the Markdown preview or a source template.

## Security and Privacy

- No remote APIs are called.
- No analytics or telemetry are included.
- No CDN, external CSS, external scripts, or external fonts are used in the final HTML.
- A runtime guard blocks `fetch` and `XMLHttpRequest` calls to remote URLs.
- Raw HTML rendering is disabled in the Markdown parser.
- Preview links are displayed but click navigation is prevented inside the app.

More detail is available in [docs/LIMITATIONS.md](docs/LIMITATIONS.md) and [docs/TEMPLATE_SUPPORT.md](docs/TEMPLATE_SUPPORT.md).
