import type { StyleMap } from "../docx/styleMapper";
import type { TemplateInfo } from "../docx/templateReader";
import type { Translate } from "../i18n";

export function renderWarnings(container: HTMLElement, warnings: string[], t: Translate): void {
  container.hidden = warnings.length === 0;
  if (!warnings.length) {
    container.innerHTML = "";
    return;
  }

  container.innerHTML = `<strong>${escapeHtml(t("warnings.title"))}</strong><ul>${warnings.map((warning) => `<li>${escapeHtml(t(warning))}</li>`).join("")}</ul>`;
}

export function renderTemplateStatus(
  container: HTMLElement,
  template: TemplateInfo | null,
  t: Translate,
  formatList: (values: string[]) => string,
): void {
  if (!template) {
    container.className = "message";
    container.textContent = t("template.status.empty");
    return;
  }

  container.className = "message ok";
  const parts = [
    template.hasDocumentPart ? t("template.part.base_document") : null,
    template.hasStylesPart ? t("template.part.styles") : null,
    template.hasThemePart ? t("template.part.theme") : null,
    template.hasNumberingPart ? t("template.part.numbering") : null,
    template.sectPrXml ? t("template.part.section_refs") : null,
  ].filter((part): part is string => Boolean(part));
  container.textContent = parts.length
    ? t("template.status.loaded_with_parts", {
      fileName: template.fileName,
      count: template.styles.length,
      parts: formatList(parts),
    })
    : t("template.status.loaded_simple", {
      fileName: template.fileName,
      count: template.styles.length,
    });
}

export function renderStyleDiagnostics(container: HTMLElement, template: TemplateInfo | null, styleMap: StyleMap, t: Translate): void {
  const styleList = template?.styles.length
    ? `<div class="style-list"><strong>${escapeHtml(t("style.detected_title"))}</strong> ${template.styles
      .slice(0, 80)
      .map((style) => `${escapeHtml(style.name)} (${escapeHtml(style.styleId)}, ${escapeHtml(style.type)})`)
      .join("; ")}${template.styles.length > 80 ? `; ${escapeHtml(t("style.more"))}` : ""}</div>`
    : "";

  const rows = styleMap.diagnostics
    .map((item) => {
      const badgeClass = item.source === "template" ? "template" : "fallback";
      const badgeText = item.source === "template" ? t("style.source.template") : t("style.source.fallback");
      return [
        '<div class="diagnostic-row">',
        `<b>${escapeHtml(styleElementLabel(item.element, t))}</b>`,
        `<span>${escapeHtml(item.styleName)} <code>${escapeHtml(item.styleId)}</code></span>`,
        `<span class="badge ${badgeClass}">${escapeHtml(badgeText)}</span>`,
        "</div>",
      ].join("");
    })
    .join("");

  container.innerHTML = `${styleList}<div class="diagnostic-grid">${rows}</div>`;
}

function styleElementLabel(element: string, t: Translate): string {
  const keys: Record<string, string> = {
    Paragraphs: "style.element.paragraphs",
    Blockquotes: "style.element.blockquotes",
    "Code blocks": "style.element.code_blocks",
    Tables: "style.element.tables",
    Links: "style.element.links",
  };
  return keys[element] ? t(keys[element]) : element;
}

function escapeHtml(value: string): string {
  return value
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;");
}
