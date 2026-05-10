import { buildStyleMap } from "./docx/styleMapper";
import { generateDocxBlob } from "./docx/generateDocx";
import { readTemplateFile, type TemplateInfo } from "./docx/templateReader";
import { applyI18nToDom, formatList, getCurrentLocale, hasTranslationKey, initializeI18n, setLocale, t, type TranslationVars } from "./i18n";
import { parseMarkdown, type MarkdownParseResult } from "./markdown/parseMarkdown";
import { installNetworkGuard, runOfflineSelfCheck } from "./security/offlineSelfCheck";
import { ensureDocxFileName, formatInputStats, readTextFile, sampleMarkdown } from "./ui/editor";
import { renderStyleDiagnostics, renderTemplateStatus, renderWarnings } from "./ui/diagnostics";
import { renderPreview } from "./ui/preview";

type AppState = {
  markdown: string;
  parsed: MarkdownParseResult;
  template: TemplateInfo | null;
  exportStatus: ExportStatus;
};

type ExportStatus = {
  kind: "" | "error" | "ok" | "warn";
  key?: string;
  message?: string;
  vars?: TranslationVars;
};

const elements = {
  markdownInput: byId<HTMLTextAreaElement>("markdown-input"),
  markdownFileInput: byId<HTMLInputElement>("markdown-file-input"),
  templateFileInput: byId<HTMLInputElement>("template-file-input"),
  languageSelect: byId<HTMLSelectElement>("language-select"),
  loadMarkdownButton: byId<HTMLButtonElement>("load-markdown-button"),
  sampleButton: byId<HTMLButtonElement>("sample-button"),
  clearButton: byId<HTMLButtonElement>("clear-button"),
  livePreviewToggle: byId<HTMLInputElement>("live-preview-toggle"),
  refreshPreviewButton: byId<HTMLButtonElement>("refresh-preview-button"),
  preview: byId<HTMLElement>("preview"),
  inputStats: byId<HTMLElement>("input-stats"),
  loadTemplateButton: byId<HTMLButtonElement>("load-template-button"),
  clearTemplateButton: byId<HTMLButtonElement>("clear-template-button"),
  templateStatus: byId<HTMLElement>("template-status"),
  templateDiagnostics: byId<HTMLElement>("template-diagnostics"),
  filenameInput: byId<HTMLInputElement>("filename-input"),
  warnings: byId<HTMLElement>("warnings"),
  offlineCheck: byId<HTMLElement>("offline-check"),
  exportButton: byId<HTMLButtonElement>("export-button"),
  exportStatus: byId<HTMLElement>("export-status"),
};

initializeI18n();
installNetworkGuard(() => t("error.network_disabled"));

const state: AppState = {
  markdown: "",
  parsed: parseMarkdown(""),
  template: null,
  exportStatus: { key: "status.ready", kind: "" },
};

bindEvents();
applyStaticI18n();
renderAll();

function bindEvents(): void {
  elements.languageSelect.value = getCurrentLocale();
  elements.languageSelect.addEventListener("change", () => {
    setLocale(elements.languageSelect.value, true);
    applyStaticI18n();
    renderAll();
  });

  elements.markdownInput.addEventListener("input", () => {
    state.markdown = elements.markdownInput.value;
    updateParsed();
    renderEditorState();
    if (elements.livePreviewToggle.checked) {
      renderPreviewState();
    }
  });

  elements.loadMarkdownButton.addEventListener("click", () => {
    elements.markdownFileInput.click();
  });

  elements.markdownFileInput.addEventListener("change", async () => {
    const file = elements.markdownFileInput.files?.[0];
    if (!file) return;
    try {
      state.markdown = await readTextFile(file);
      elements.markdownInput.value = state.markdown;
      updateParsed();
      renderAll();
      setExportStatusKey("status.file_loaded", "ok", { fileName: file.name });
    } catch (error) {
      setExportError(error);
    } finally {
      elements.markdownFileInput.value = "";
    }
  });

  elements.sampleButton.addEventListener("click", () => {
    state.markdown = sampleMarkdown(t);
    elements.markdownInput.value = state.markdown;
    updateParsed();
    renderAll();
    setExportStatusKey("status.sample_loaded", "ok");
  });

  elements.clearButton.addEventListener("click", () => {
    state.markdown = "";
    elements.markdownInput.value = "";
    updateParsed();
    renderAll();
    setExportStatusKey("status.ready", "");
  });

  elements.refreshPreviewButton.addEventListener("click", () => {
    updateParsed();
    renderPreviewState();
  });

  elements.livePreviewToggle.addEventListener("change", () => {
    if (elements.livePreviewToggle.checked) {
      updateParsed();
      renderPreviewState();
    }
  });

  elements.loadTemplateButton.addEventListener("click", () => {
    elements.templateFileInput.click();
  });

  elements.templateFileInput.addEventListener("change", async () => {
    const file = elements.templateFileInput.files?.[0];
    if (!file) return;

    try {
      state.template = await readTemplateFile(file);
      renderTemplateState();
      setExportStatusKey("status.template_loaded", "ok");
    } catch (error) {
      state.template = null;
      renderTemplateState();
      setExportError(error);
    } finally {
      elements.templateFileInput.value = "";
    }
  });

  elements.clearTemplateButton.addEventListener("click", () => {
    state.template = null;
    renderTemplateState();
    setExportStatusKey("status.template_cleared", "");
  });

  elements.exportButton.addEventListener("click", () => {
    void exportDocx();
  });
}

function updateParsed(): void {
  state.parsed = parseMarkdown(state.markdown);
}

function renderAll(): void {
  renderEditorState();
  renderPreviewState();
  renderTemplateState();
  renderOfflineCheck();
  renderExportStatus();
}

function renderEditorState(): void {
  elements.inputStats.textContent = formatInputStats(state.markdown, getCurrentLocale(), t);
  renderWarnings(elements.warnings, state.parsed.warnings, t);
}

function renderPreviewState(): void {
  renderPreview(elements.preview, state.parsed.html, state.markdown.trim().length === 0, t("preview.empty"));
}

function renderTemplateState(): void {
  const styleMap = buildStyleMap(state.template);
  renderTemplateStatus(elements.templateStatus, state.template, t, formatList);
  renderStyleDiagnostics(elements.templateDiagnostics, state.template, styleMap, t);
}

function renderOfflineCheck(): void {
  const result = runOfflineSelfCheck();
  elements.offlineCheck.className = result.ok ? "message ok" : "message warn";
  elements.offlineCheck.textContent = result.messages.map((message) => t(message.key, message.vars)).join(" ");
}

async function exportDocx(): Promise<void> {
  updateParsed();
  renderEditorState();

  if (!state.markdown.trim()) {
    setExportStatusKey("status.add_markdown", "warn");
    return;
  }

  elements.exportButton.disabled = true;
  setExportStatusKey("status.generating", "");

  try {
    const fileName = ensureDocxFileName(elements.filenameInput.value);
    elements.filenameInput.value = fileName;
    const blob = await generateDocxBlob(state.parsed.model, {
      template: state.template,
      title: fileName.replace(/\.docx$/i, ""),
    });
    downloadBlob(blob, fileName);
    setExportStatusKey("status.exported", "ok", { fileName });
  } catch (error) {
    setExportError(error);
  } finally {
    elements.exportButton.disabled = false;
  }
}

function downloadBlob(blob: Blob, fileName: string): void {
  const url = URL.createObjectURL(blob);
  const anchor = document.createElement("a");
  anchor.href = url;
  anchor.download = fileName;
  anchor.rel = "noreferrer";
  document.body.append(anchor);
  anchor.click();
  anchor.remove();
  window.setTimeout(() => URL.revokeObjectURL(url), 500);
}

function applyStaticI18n(): void {
  applyI18nToDom();
  elements.languageSelect.value = getCurrentLocale();
}

function setExportStatusKey(key: string, kind: "" | "ok" | "warn" | "error", vars?: TranslationVars): void {
  state.exportStatus = { key, kind, vars };
  renderExportStatus();
}

function setExportError(error: unknown): void {
  const message = error instanceof Error ? error.message : String(error);
  if (hasTranslationKey(message)) {
    setExportStatusKey(message, "error");
    return;
  }

  setExportStatusKey("error.unexpected", "error", { message });
}

function renderExportStatus(): void {
  const status = state.exportStatus;
  elements.exportStatus.className = status.kind ? `message ${status.kind}` : "message";
  elements.exportStatus.textContent = status.key ? t(status.key, status.vars) : status.message || "";
}

function byId<T extends HTMLElement>(id: string): T {
  const element = document.getElementById(id);
  if (!element) {
    throw new Error(`Missing element #${id}`);
  }
  return element as T;
}
