# Template Support

See [docs/TEMPLATE_SUPPORT.md](docs/TEMPLATE_SUPPORT.md) for the full template support notes.

In short: `.docx` and `.dotx` files are read as OpenXML ZIP packages. The app uses complete templates as the base package and replaces the main body with converted Markdown, preserving practical parts such as headers, footers, media, settings, fields in preserved parts, styles, theme, and section references. It still does not execute macros, mail merge logic, or arbitrary placeholder/content-control workflows.
