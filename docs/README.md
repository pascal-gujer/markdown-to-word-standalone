# Markdown to Word Offline

Markdown to Word Offline is a browser-only tool that converts Markdown into a downloadable `.docx` file. It is designed for restricted and offline machines: no server, no telemetry, no CDN, no external fonts, and no runtime network access.

## Build

```sh
npm install
npm run build
```

The build writes one self-contained file:

```text
dist/md-to-docx.html
```

For development:

```sh
npm run dev
```

For tests:

```sh
npm run test
```

`npm run build` also runs the offline build validator. It fails if the final HTML contains raw `http://`, `https://`, CDN references, external script tags, external stylesheet tags, or module imports.

## Offline Use

Build the project, then move or open `dist/md-to-docx.html` on the offline machine. The app is intended to work directly from `file://` in browsers that allow local file selection, Blob downloads, and ZIP generation APIs.

## Basic Workflow

1. Paste Markdown into the editor, load a local `.md`, `.markdown`, or `.txt` file, or load a `.zip` bundle containing Markdown plus referenced images.
2. Optionally load a `.docx` or `.dotx` style template.
3. Set the output file name and export the `.docx`.

Files are processed locally in the browser and are not uploaded.

ZIP import is an additional workflow, not a replacement for the editor. The app chooses the best Markdown file in the archive and resolves image references relative to that file. Supported embedded image formats are `.png`, `.jpg`, `.jpeg`, and `.gif`.

## Template Support

The app reads `.docx` and `.dotx` files as OpenXML ZIP packages. When a complete package is loaded, export uses the template package as the base and replaces the main document body with the converted Markdown. This preserves practical package parts such as headers, footers, media, theme, settings, footnotes, endnotes, fields in preserved parts, custom XML, and section references.

It inspects:

- `word/styles.xml`
- `word/theme/theme1.xml`
- `word/numbering.xml`
- simple section/page settings from `word/document.xml`

Markdown elements are mapped to Word style IDs where practical:

- `#` through `######` use Heading 1 through Heading 6 style names when detected.
- Paragraphs use Normal.
- Blockquotes use Quote when present.
- Code blocks use Code, Source Code, or a generated Code Block fallback.
- Tables use Table Grid or a generated fallback table style.

This is intentionally incremental. It mirrors common style names and preserves useful package parts, but it does not attempt full Word automation or placeholder/mail-merge execution.

## Why `.doc` Is Not Supported

Legacy `.doc` files are a binary Word format. They require a substantially different parser/writer and are not a good fit for a safe, browser-only offline converter. `.docx` and `.dotx` are ZIP-based OpenXML packages, so they can be generated and inspected in the browser without native binaries.

## DOCX Generation vs. Word Rendering

The app generates standards-oriented WordprocessingML. Microsoft Word, LibreOffice, Apple Pages, and other editors may render the same `.docx` slightly differently. The goal is a reliable editable Word document, not pixel-perfect reproduction of Markdown preview HTML or a source template.

## Security and Privacy Model

- No remote APIs are called.
- No analytics or telemetry are included.
- No CDN, external CSS, external scripts, or external fonts are used in the final HTML.
- A runtime guard blocks `fetch` and `XMLHttpRequest` calls to remote `http` and `https` URLs.
- Markdown preview disables raw HTML rendering through the Markdown parser.
- Preview links are displayed but click navigation is prevented inside the app.

The generated `.docx` may contain external hyperlink relationships if the Markdown contains links. That is document content, not runtime app network activity.

Markdown images are only previewed and embedded when they come from a local ZIP bundle. Remote image URLs are replaced with alt text in the preview and are not loaded by the app.
