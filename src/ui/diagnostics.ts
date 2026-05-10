import type { StyleMap } from "../docx/styleMapper";
import type { TemplateInfo } from "../docx/templateReader";

export function renderWarnings(container: HTMLElement, warnings: string[]): void {
  container.hidden = warnings.length === 0;
  if (!warnings.length) {
    container.innerHTML = "";
    return;
  }

  container.innerHTML = `<strong>Export warnings</strong><ul>${warnings.map((warning) => `<li>${escapeHtml(warning)}</li>`).join("")}</ul>`;
}

export function renderTemplateStatus(container: HTMLElement, template: TemplateInfo | null): void {
  if (!template) {
    container.className = "message";
    container.textContent = "No template loaded. Sensible default Word styles will be generated.";
    return;
  }

  container.className = "message ok";
  const parts = [
    template.hasDocumentPart ? "base document package retained" : null,
    template.hasStylesPart ? "styles" : null,
    template.hasThemePart ? "theme" : null,
    template.hasNumberingPart ? "numbering detected" : null,
    template.sectPrXml ? "section, header, and footer references" : null,
  ].filter(Boolean);
  container.textContent = `${template.fileName}: ${template.styles.length} style(s) detected${parts.length ? `, ${parts.join(", ")}` : ""}.`;
}

export function renderStyleDiagnostics(container: HTMLElement, template: TemplateInfo | null, styleMap: StyleMap): void {
  const styleList = template?.styles.length
    ? `<div class="style-list"><strong>Detected styles:</strong> ${template.styles
      .slice(0, 80)
      .map((style) => `${escapeHtml(style.name)} (${escapeHtml(style.styleId)}, ${escapeHtml(style.type)})`)
      .join("; ")}${template.styles.length > 80 ? "; ..." : ""}</div>`
    : "";

  const rows = styleMap.diagnostics
    .map((item) => {
      const badgeClass = item.source === "template" ? "template" : "fallback";
      const badgeText = item.source === "template" ? "Template" : "Fallback";
      return [
        '<div class="diagnostic-row">',
        `<b>${escapeHtml(item.element)}</b>`,
        `<span>${escapeHtml(item.styleName)} <code>${escapeHtml(item.styleId)}</code></span>`,
        `<span class="badge ${badgeClass}">${badgeText}</span>`,
        "</div>",
      ].join("");
    })
    .join("");

  container.innerHTML = `${styleList}<div class="diagnostic-grid">${rows}</div>`;
}

function escapeHtml(value: string): string {
  return value
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;");
}
