import JSZip from "jszip";
import {
  collectMarkdownImageReferences,
  directoryName,
  normalizeZipPath,
  readImageDimensions,
  resolveMarkdownAssetPath,
  supportedImageType,
  uniqueImageReferences,
  type MarkdownImageAsset,
  type MarkdownImageBundle,
} from "../markdown/imageBundle";

export type MarkdownZipImport = {
  markdown: string;
  markdownPath: string;
  markdownFileCount: number;
  imageBundle: MarkdownImageBundle;
  missingReferences: string[];
  unsupportedReferences: string[];
};

const markdownExtensions = new Set([".md", ".markdown", ".txt"]);

export async function readMarkdownZipFile(file: File): Promise<MarkdownZipImport> {
  let zip: JSZip;
  try {
    zip = await JSZip.loadAsync(await file.arrayBuffer());
  } catch (_error) {
    throw new Error("error.zip_unreadable");
  }

  const markdownPaths = Object.keys(zip.files)
    .filter((path) => !zip.files[path].dir && isMarkdownPath(path))
    .sort(compareMarkdownCandidates);
  const markdownPath = markdownPaths[0];
  if (!markdownPath) {
    throw new Error("error.zip_no_markdown");
  }

  const markdown = await zip.file(markdownPath)!.async("string");
  const markdownBasePath = directoryName(markdownPath);
  const references = uniqueImageReferences(collectMarkdownImageReferences(markdown));
  const missingReferences: string[] = [];
  const unsupportedReferences: string[] = [];
  const assets: MarkdownImageAsset[] = [];

  for (const reference of references) {
    const resolvedPath = resolveMarkdownAssetPath(reference.src, markdownBasePath);
    if (!resolvedPath) {
      unsupportedReferences.push(reference.src);
      continue;
    }

    const type = supportedImageType(resolvedPath);
    if (!type) {
      unsupportedReferences.push(reference.src);
      continue;
    }

    const entry = findZipFile(zip, resolvedPath);
    if (!entry) {
      missingReferences.push(reference.src);
      continue;
    }

    const data = await entry.async("uint8array");
    const dimensions = readImageDimensions(data, type.extension);
    if (!dimensions || dimensions.widthPx <= 0 || dimensions.heightPx <= 0) {
      unsupportedReferences.push(reference.src);
      continue;
    }

    assets.push({
      path: normalizeZipPath(entry.name) || resolvedPath,
      fileName: fileNameFromPath(entry.name),
      extension: type.extension,
      contentType: type.contentType,
      data,
      widthPx: dimensions.widthPx,
      heightPx: dimensions.heightPx,
    });
  }

  return {
    markdown,
    markdownPath,
    markdownFileCount: markdownPaths.length,
    imageBundle: {
      sourceFileName: file.name,
      markdownPath,
      markdownBasePath,
      assets,
    },
    missingReferences,
    unsupportedReferences,
  };
}

function isMarkdownPath(path: string): boolean {
  const normalized = normalizeZipPath(path);
  if (!normalized || /^__MACOSX\//i.test(normalized)) {
    return false;
  }
  const lower = normalized.toLowerCase();
  return Array.from(markdownExtensions).some((extension) => lower.endsWith(extension));
}

function compareMarkdownCandidates(left: string, right: string): number {
  return markdownScore(left) - markdownScore(right) || left.localeCompare(right);
}

function markdownScore(path: string): number {
  const normalized = normalizeZipPath(path) || path;
  const lower = normalized.toLowerCase();
  const depth = normalized.split("/").length - 1;
  const name = lower.slice(lower.lastIndexOf("/") + 1);
  const extensionPenalty = name.endsWith(".txt") ? 20 : 0;
  const nameBonus = name === "index.md" || name === "document.md" || name === "readme.md" ? -10 : 0;
  return depth * 10 + extensionPenalty + nameBonus;
}

function findZipFile(zip: JSZip, normalizedPath: string): JSZip.JSZipObject | null {
  const exact = zip.file(normalizedPath);
  if (exact) {
    return exact;
  }

  const lower = normalizedPath.toLowerCase();
  const match = Object.keys(zip.files).find((path) => (normalizeZipPath(path) || path).toLowerCase() === lower);
  return match ? zip.file(match) : null;
}

function fileNameFromPath(path: string): string {
  const normalized = normalizeZipPath(path) || path;
  return normalized.slice(normalized.lastIndexOf("/") + 1) || "image";
}
