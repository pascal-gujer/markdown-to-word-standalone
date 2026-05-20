import { buildStyleMap } from "./docx/styleMapper";
import { generateDocxBlob } from "./docx/generateDocx";
import { readTemplateFile, type TemplateInfo } from "./docx/templateReader";
import { applyI18nToDom, formatList, getCurrentLocale, hasTranslationKey, initializeI18n, plural, setLocale, t, type TranslationVars } from "./i18n";
import {
  collectMarkdownImageReferences,
  disposeImagePreviewUrls,
  prepareImagePreviewUrls,
  resolveBundleImage,
  type MarkdownImageBundle,
} from "./markdown/imageBundle";
import { parseMarkdown, type MarkdownParseResult } from "./markdown/parseMarkdown";
import { installNetworkGuard, runOfflineSelfCheck } from "./security/offlineSelfCheck";
import { ensureDocxFileName, formatInputStats, readTextFile, sampleMarkdown } from "./ui/editor";
import { renderStyleDiagnostics, renderTemplateStatus, renderWarnings } from "./ui/diagnostics";
import { renderPreview } from "./ui/preview";
import { extractMarkdownFromZipSource, loadMarkdownZipSource, type MarkdownZipImport, type MarkdownZipSource } from "./ui/zipImport";

type AppState = {
  markdown: string;
  images: MarkdownImageBundle | null;
  parsed: MarkdownParseResult;
  template: TemplateInfo | null;
  exportStatus: ExportStatus;
  zipSource: MarkdownZipSource | null;
};

type ExportStatusKind = "" | "error" | "ok" | "warn";

type ExportStatus = {
  kind: ExportStatusKind;
  build: () => string;
};

const elements = {
  editorPanel: byId<HTMLElement>("editor-panel"),
  markdownInput: byId<HTMLTextAreaElement>("markdown-input"),
  markdownFileInput: byId<HTMLInputElement>("markdown-file-input"),
  zipFileInput: byId<HTMLInputElement>("zip-file-input"),
  zipSourceBar: byId<HTMLElement>("zip-source-bar"),
  zipMarkdownSelect: byId<HTMLSelectElement>("zip-markdown-select"),
  templateFileInput: byId<HTMLInputElement>("template-file-input"),
  languageSelect: byId<HTMLSelectElement>("language-select"),
  loadMarkdownButton: byId<HTMLButtonElement>("load-markdown-button"),
  loadZipButton: byId<HTMLButtonElement>("load-zip-button"),
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
  images: null,
  parsed: parseMarkdown(""),
  template: null,
  exportStatus: { kind: "", build: () => t("status.ready") },
  zipSource: null,
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

  elements.loadZipButton.addEventListener("click", () => {
    elements.zipFileInput.click();
  });

  elements.markdownFileInput.addEventListener("change", async () => {
    const file = elements.markdownFileInput.files?.[0];
    if (!file) return;
    try {
      await loadMarkdownFile(file);
    } catch (error) {
      setExportError(error);
    } finally {
      elements.markdownFileInput.value = "";
    }
  });

  elements.zipFileInput.addEventListener("change", async () => {
    const file = elements.zipFileInput.files?.[0];
    if (!file) return;
    try {
      await loadZipFile(file);
    } catch (error) {
      setZipSource(null);
      setImageBundle(null);
      setExportError(error);
    } finally {
      elements.zipFileInput.value = "";
    }
  });

  elements.zipMarkdownSelect.addEventListener("change", async () => {
    if (!state.zipSource) return;
    try {
      await applyZipExtraction(state.zipSource, elements.zipMarkdownSelect.value);
    } catch (error) {
      setExportError(error);
    }
  });

  elements.sampleButton.addEventListener("click", () => {
    setZipSource(null);
    setImageBundle(null);
    state.markdown = sampleMarkdown(t);
    elements.markdownInput.value = state.markdown;
    updateParsed();
    renderAll();
    setExportStatusKey("status.sample_loaded", "ok");
  });

  elements.clearButton.addEventListener("click", () => {
    setZipSource(null);
    setImageBundle(null);
    state.markdown = "";
    elements.markdownInput.value = "";
    updateParsed();
    renderAll();
    setExportStatusKey("status.ready", "");
  });

  bindDropTarget(elements.editorPanel);

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
  state.parsed = parseMarkdown(state.markdown, { imageMode: state.images ? "embedded" : "text" });
  if (state.images && collectMarkdownImageReferences(state.markdown).some((reference) => !resolveBundleImage(state.images, reference.src))) {
    if (!state.parsed.warnings.includes("warning.images")) {
      state.parsed.warnings.push("warning.images");
    }
  }
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
  renderPreview(elements.preview, state.parsed.html, state.markdown.trim().length === 0, t("preview.empty"), state.images);
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
      images: state.images,
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

function setImageBundle(bundle: MarkdownImageBundle | null): void {
  disposeImagePreviewUrls(state.images);
  state.images = bundle;
  if (state.images) {
    prepareImagePreviewUrls(state.images);
  }
}

function setZipSource(source: MarkdownZipSource | null): void {
  state.zipSource = source;
  renderZipSourceBar();
}

function renderZipSourceBar(): void {
  const source = state.zipSource;
  if (!source || source.markdownPaths.length <= 1) {
    elements.zipSourceBar.hidden = true;
    elements.zipMarkdownSelect.innerHTML = "";
    return;
  }

  elements.zipSourceBar.hidden = false;
  elements.zipMarkdownSelect.innerHTML = "";
  for (const path of source.markdownPaths) {
    const option = document.createElement("option");
    option.value = path;
    option.textContent = path;
    elements.zipMarkdownSelect.append(option);
  }
}

async function loadMarkdownFile(file: File): Promise<void> {
  setZipSource(null);
  setImageBundle(null);
  state.markdown = await readTextFile(file);
  elements.markdownInput.value = state.markdown;
  setFilenameFromSource(file.name);
  updateParsed();
  renderAll();
  setExportStatusKey("status.file_loaded", "ok", { fileName: file.name });
}

async function loadZipFile(file: File): Promise<void> {
  const source = await loadMarkdownZipSource(file);
  setZipSource(source);
  await applyZipExtraction(source, source.preferredPath);
}

async function applyZipExtraction(source: MarkdownZipSource, markdownPath: string): Promise<void> {
  const imported = await extractMarkdownFromZipSource(source, markdownPath);
  setImageBundle(imported.imageBundle);
  state.markdown = imported.markdown;
  elements.markdownInput.value = state.markdown;
  elements.zipMarkdownSelect.value = markdownPath;
  setFilenameFromSource(markdownPath);
  updateParsed();
  renderAll();
  setZipImportStatus(imported);
}

function setZipImportStatus(imported: MarkdownZipImport): void {
  setExportStatus("ok", () => {
    const fileName = imported.imageBundle.sourceFileName;
    const imageCount = imported.imageBundle.assets.length;
    const notes = collectZipImportNotes(imported);
    const baseKey = notes.length ? "status.zip_loaded_with_notes" : "status.zip_loaded";
    return plural(baseKey, imageCount, {
      fileName,
      markdownPath: imported.markdownPath,
      notes: notes.join(" "),
    });
  });
}

function collectZipImportNotes(imported: MarkdownZipImport): string[] {
  const notes: string[] = [];
  if (imported.markdownFileCount > 1) {
    notes.push(t("zip.note.multiple_markdown", {
      count: imported.markdownFileCount,
      markdownPath: imported.markdownPath,
    }));
  }
  if (imported.missingReferences.length) {
    notes.push(plural("zip.note.missing_images", imported.missingReferences.length, {
      examples: formatExamples(imported.missingReferences),
    }));
  }
  if (imported.unsupportedReferences.length) {
    notes.push(plural("zip.note.unsupported_images", imported.unsupportedReferences.length, {
      examples: formatExamples(imported.unsupportedReferences),
    }));
  }
  if (imported.unreadableReferences.length) {
    notes.push(plural("zip.note.unreadable_images", imported.unreadableReferences.length, {
      examples: formatExamples(imported.unreadableReferences),
    }));
  }
  return notes;
}

function formatExamples(paths: string[], maxExamples = 3): string {
  return paths.slice(0, maxExamples).join(", ");
}

function setFilenameFromSource(sourcePath: string): void {
  const base = sourcePath.replace(/^.*[\\/]/, "").replace(/\.[^.]+$/, "");
  if (!base) return;
  elements.filenameInput.value = ensureDocxFileName(base);
}

function bindDropTarget(target: HTMLElement): void {
  let dragDepth = 0;

  target.addEventListener("dragenter", (event) => {
    if (!hasFileTransfer(event.dataTransfer)) return;
    event.preventDefault();
    dragDepth += 1;
    target.classList.add("is-drop-target");
  });

  target.addEventListener("dragover", (event) => {
    if (!hasFileTransfer(event.dataTransfer)) return;
    event.preventDefault();
    if (event.dataTransfer) {
      event.dataTransfer.dropEffect = "copy";
    }
  });

  target.addEventListener("dragleave", () => {
    dragDepth = Math.max(0, dragDepth - 1);
    if (dragDepth === 0) {
      target.classList.remove("is-drop-target");
    }
  });

  target.addEventListener("drop", async (event) => {
    if (!hasFileTransfer(event.dataTransfer)) return;
    event.preventDefault();
    dragDepth = 0;
    target.classList.remove("is-drop-target");
    const file = event.dataTransfer?.files?.[0];
    if (!file) return;
    try {
      if (isZipFile(file)) {
        await loadZipFile(file);
      } else {
        await loadMarkdownFile(file);
      }
    } catch (error) {
      if (isZipFile(file)) {
        setZipSource(null);
      }
      setImageBundle(null);
      setExportError(error);
    }
  });
}

function hasFileTransfer(transfer: DataTransfer | null): boolean {
  if (!transfer) return false;
  return Array.from(transfer.types || []).includes("Files");
}

function isZipFile(file: File): boolean {
  const lower = file.name.toLowerCase();
  return lower.endsWith(".zip") || /zip/i.test(file.type);
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

function setExportStatus(kind: ExportStatusKind, build: () => string): void {
  state.exportStatus = { kind, build };
  renderExportStatus();
}

function setExportStatusKey(key: string, kind: ExportStatusKind, vars?: TranslationVars): void {
  setExportStatus(kind, () => t(key, vars));
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
  elements.exportStatus.textContent = status.build();
}

function byId<T extends HTMLElement>(id: string): T {
  const element = document.getElementById(id);
  if (!element) {
    throw new Error(`Missing element #${id}`);
  }
  return element as T;
}
