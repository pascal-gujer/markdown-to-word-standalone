import { buildStyleMap } from "./docx/styleMapper";
import { generateDocxBlob } from "./docx/generateDocx";
import { readTemplateFile, type TemplateInfo } from "./docx/templateReader";
import { parseMarkdown, type MarkdownParseResult } from "./markdown/parseMarkdown";
import { installNetworkGuard, runOfflineSelfCheck } from "./security/offlineSelfCheck";
import { ensureDocxFileName, formatInputStats, readTextFile, sampleMarkdown } from "./ui/editor";
import { renderStyleDiagnostics, renderTemplateStatus, renderWarnings } from "./ui/diagnostics";
import { renderPreview } from "./ui/preview";

type AppState = {
  markdown: string;
  parsed: MarkdownParseResult;
  template: TemplateInfo | null;
};

const elements = {
  markdownInput: byId<HTMLTextAreaElement>("markdown-input"),
  markdownFileInput: byId<HTMLInputElement>("markdown-file-input"),
  templateFileInput: byId<HTMLInputElement>("template-file-input"),
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

installNetworkGuard();

const state: AppState = {
  markdown: "",
  parsed: parseMarkdown(""),
  template: null,
};

bindEvents();
renderAll();

function bindEvents(): void {
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
      setExportStatus(`Loaded ${file.name}.`, "ok");
    } catch (error) {
      setExportStatus(errorMessage(error), "error");
    } finally {
      elements.markdownFileInput.value = "";
    }
  });

  elements.sampleButton.addEventListener("click", () => {
    state.markdown = sampleMarkdown;
    elements.markdownInput.value = state.markdown;
    updateParsed();
    renderAll();
    setExportStatus("Sample Markdown loaded.", "ok");
  });

  elements.clearButton.addEventListener("click", () => {
    state.markdown = "";
    elements.markdownInput.value = "";
    updateParsed();
    renderAll();
    setExportStatus("Ready.", "");
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
      setExportStatus("Template loaded. Style mappings updated.", "ok");
    } catch (error) {
      state.template = null;
      renderTemplateState();
      setExportStatus(errorMessage(error), "error");
    } finally {
      elements.templateFileInput.value = "";
    }
  });

  elements.clearTemplateButton.addEventListener("click", () => {
    state.template = null;
    renderTemplateState();
    setExportStatus("Template cleared. Fallback styles will be used.", "");
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
}

function renderEditorState(): void {
  elements.inputStats.textContent = formatInputStats(state.markdown);
  renderWarnings(elements.warnings, state.parsed.warnings);
}

function renderPreviewState(): void {
  renderPreview(elements.preview, state.parsed.html, state.markdown.trim().length === 0);
}

function renderTemplateState(): void {
  const styleMap = buildStyleMap(state.template);
  renderTemplateStatus(elements.templateStatus, state.template);
  renderStyleDiagnostics(elements.templateDiagnostics, state.template, styleMap);
}

function renderOfflineCheck(): void {
  const result = runOfflineSelfCheck();
  elements.offlineCheck.className = result.ok ? "message ok" : "message warn";
  elements.offlineCheck.textContent = result.messages.join(" ");
}

async function exportDocx(): Promise<void> {
  updateParsed();
  renderEditorState();

  if (!state.markdown.trim()) {
    setExportStatus("Add Markdown before exporting.", "warn");
    return;
  }

  elements.exportButton.disabled = true;
  setExportStatus("Generating DOCX locally...", "");

  try {
    const fileName = ensureDocxFileName(elements.filenameInput.value);
    elements.filenameInput.value = fileName;
    const blob = await generateDocxBlob(state.parsed.model, {
      template: state.template,
      title: fileName.replace(/\.docx$/i, ""),
    });
    downloadBlob(blob, fileName);
    setExportStatus(`Exported ${fileName}.`, "ok");
  } catch (error) {
    setExportStatus(errorMessage(error), "error");
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

function setExportStatus(message: string, kind: "" | "ok" | "warn" | "error"): void {
  elements.exportStatus.className = kind ? `message ${kind}` : "message";
  elements.exportStatus.textContent = message;
}

function byId<T extends HTMLElement>(id: string): T {
  const element = document.getElementById(id);
  if (!element) {
    throw new Error(`Missing element #${id}`);
  }
  return element as T;
}

function errorMessage(error: unknown): string {
  return error instanceof Error ? error.message : String(error);
}
