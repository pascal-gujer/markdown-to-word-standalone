import MarkdownIt from "markdown-it";
import type Token from "markdown-it/lib/token.mjs";

export type SupportedImageExtension = "gif" | "jpeg" | "jpg" | "png";

export type MarkdownImageAsset = {
  path: string;
  fileName: string;
  extension: SupportedImageExtension;
  contentType: "image/gif" | "image/jpeg" | "image/png";
  data: Uint8Array;
  widthPx: number;
  heightPx: number;
  previewUrl?: string;
};

export type MarkdownImageBundle = {
  sourceFileName: string;
  markdownPath: string;
  markdownBasePath: string;
  assets: MarkdownImageAsset[];
};

export type MarkdownImageReference = {
  src: string;
  alt: string;
};

const imageReferenceParser = new MarkdownIt({
  html: false,
  linkify: false,
  typographer: false,
});

export function collectMarkdownImageReferences(markdown: string): MarkdownImageReference[] {
  const tokens = imageReferenceParser.parse(markdown, {});
  const references: MarkdownImageReference[] = [];
  walkTokens(tokens, (token) => {
    if (token.type !== "image") {
      return;
    }
    const src = token.attrGet("src") || "";
    if (src) {
      references.push({
        src,
        alt: token.content || token.attrGet("alt") || "image",
      });
    }
  });
  return references;
}

export function resolveBundleImage(bundle: MarkdownImageBundle | null | undefined, src: string): MarkdownImageAsset | undefined {
  if (!bundle) {
    return undefined;
  }
  const resolvedPath = resolveMarkdownAssetPath(src, bundle.markdownBasePath);
  if (!resolvedPath) {
    return undefined;
  }

  const exact = bundle.assets.find((asset) => asset.path === resolvedPath);
  if (exact) {
    return exact;
  }

  const lower = resolvedPath.toLowerCase();
  return bundle.assets.find((asset) => asset.path.toLowerCase() === lower);
}

export function resolveMarkdownAssetPath(src: string, markdownBasePath: string): string | undefined {
  const clean = stripUrlDecorations(src.trim());
  if (!clean || isRemoteOrDataUrl(clean)) {
    return undefined;
  }

  const decoded = decodePath(clean);
  const relativeToRoot = decoded.startsWith("/")
    ? decoded.slice(1)
    : [markdownBasePath, decoded].filter(Boolean).join("/");
  return normalizeZipPath(relativeToRoot);
}

export function normalizeZipPath(path: string): string | undefined {
  const normalized = path.replace(/\\/g, "/");
  const parts: string[] = [];
  for (const rawPart of normalized.split("/")) {
    const part = rawPart.trim();
    if (!part || part === ".") {
      continue;
    }
    if (part === "..") {
      if (!parts.length) {
        return undefined;
      }
      parts.pop();
      continue;
    }
    parts.push(part);
  }
  return parts.join("/");
}

export function directoryName(path: string): string {
  const normalized = normalizeZipPath(path) || "";
  const index = normalized.lastIndexOf("/");
  return index >= 0 ? normalized.slice(0, index) : "";
}

export function supportedImageType(path: string): Pick<MarkdownImageAsset, "contentType" | "extension"> | undefined {
  const extension = extensionOf(path);
  if (extension === "png") {
    return { extension, contentType: "image/png" };
  }
  if (extension === "jpg" || extension === "jpeg") {
    return { extension, contentType: "image/jpeg" };
  }
  if (extension === "gif") {
    return { extension, contentType: "image/gif" };
  }
  return undefined;
}

export function readImageDimensions(data: Uint8Array, extension: SupportedImageExtension): { widthPx: number; heightPx: number } | undefined {
  if (extension === "png") {
    return readPngDimensions(data);
  }
  if (extension === "gif") {
    return readGifDimensions(data);
  }
  return readJpegDimensions(data);
}

export function prepareImagePreviewUrls(bundle: MarkdownImageBundle): void {
  if (typeof URL === "undefined" || typeof Blob === "undefined") {
    return;
  }
  for (const asset of bundle.assets) {
    if (!asset.previewUrl) {
      asset.previewUrl = URL.createObjectURL(new Blob([arrayBufferCopy(asset.data)], { type: asset.contentType }));
    }
  }
}

export function disposeImagePreviewUrls(bundle: MarkdownImageBundle | null): void {
  if (!bundle || typeof URL === "undefined") {
    return;
  }
  for (const asset of bundle.assets) {
    if (asset.previewUrl) {
      URL.revokeObjectURL(asset.previewUrl);
      asset.previewUrl = undefined;
    }
  }
}

export function uniqueImageReferences(references: MarkdownImageReference[]): MarkdownImageReference[] {
  const seen = new Set<string>();
  const unique: MarkdownImageReference[] = [];
  for (const reference of references) {
    const key = stripUrlDecorations(reference.src.trim());
    if (!key || seen.has(key)) {
      continue;
    }
    seen.add(key);
    unique.push(reference);
  }
  return unique;
}

function walkTokens(tokens: Token[], visitor: (token: Token) => void): void {
  for (const token of tokens) {
    visitor(token);
    if (token.children) {
      walkTokens(token.children, visitor);
    }
  }
}

function readPngDimensions(data: Uint8Array): { widthPx: number; heightPx: number } | undefined {
  if (
    data.length < 24 ||
    data[0] !== 0x89 ||
    data[1] !== 0x50 ||
    data[2] !== 0x4e ||
    data[3] !== 0x47
  ) {
    return undefined;
  }
  return {
    widthPx: readUInt32BE(data, 16),
    heightPx: readUInt32BE(data, 20),
  };
}

function readGifDimensions(data: Uint8Array): { widthPx: number; heightPx: number } | undefined {
  if (
    data.length < 10 ||
    data[0] !== 0x47 ||
    data[1] !== 0x49 ||
    data[2] !== 0x46
  ) {
    return undefined;
  }
  return {
    widthPx: data[6] | (data[7] << 8),
    heightPx: data[8] | (data[9] << 8),
  };
}

function readJpegDimensions(data: Uint8Array): { widthPx: number; heightPx: number } | undefined {
  if (data.length < 4 || data[0] !== 0xff || data[1] !== 0xd8) {
    return undefined;
  }

  let offset = 2;
  while (offset + 9 < data.length) {
    while (offset < data.length && data[offset] === 0xff) {
      offset += 1;
    }
    const marker = data[offset];
    offset += 1;
    if (!marker || marker === 0xd9 || marker === 0xda) {
      break;
    }
    if (offset + 2 > data.length) {
      break;
    }
    const length = readUInt16BE(data, offset);
    if (length < 2 || offset + length > data.length) {
      break;
    }
    if (isJpegStartOfFrame(marker) && offset + 7 < data.length) {
      return {
        heightPx: readUInt16BE(data, offset + 3),
        widthPx: readUInt16BE(data, offset + 5),
      };
    }
    offset += length;
  }

  return undefined;
}

function isJpegStartOfFrame(marker: number): boolean {
  return (
    (marker >= 0xc0 && marker <= 0xc3) ||
    (marker >= 0xc5 && marker <= 0xc7) ||
    (marker >= 0xc9 && marker <= 0xcb) ||
    (marker >= 0xcd && marker <= 0xcf)
  );
}

function readUInt16BE(data: Uint8Array, offset: number): number {
  return (data[offset] << 8) + data[offset + 1];
}

function readUInt32BE(data: Uint8Array, offset: number): number {
  return ((data[offset] * 0x1000000) + ((data[offset + 1] << 16) | (data[offset + 2] << 8) | data[offset + 3])) >>> 0;
}

function arrayBufferCopy(data: Uint8Array): ArrayBuffer {
  const copy = new ArrayBuffer(data.byteLength);
  new Uint8Array(copy).set(data);
  return copy;
}

function extensionOf(path: string): SupportedImageExtension | undefined {
  const match = /\.([a-z0-9]+)$/i.exec(stripUrlDecorations(path));
  const extension = match?.[1].toLowerCase();
  return extension === "png" || extension === "jpg" || extension === "jpeg" || extension === "gif"
    ? extension
    : undefined;
}

function stripUrlDecorations(value: string): string {
  const hashIndex = value.indexOf("#");
  const withoutHash = hashIndex >= 0 ? value.slice(0, hashIndex) : value;
  const queryIndex = withoutHash.indexOf("?");
  return queryIndex >= 0 ? withoutHash.slice(0, queryIndex) : withoutHash;
}

function isRemoteOrDataUrl(value: string): boolean {
  return /^(?:[a-z][a-z0-9+.-]*:)?\/\//i.test(value) || /^(?:data|blob|javascript):/i.test(value);
}

function decodePath(path: string): string {
  try {
    return decodeURIComponent(path);
  } catch (_error) {
    return path;
  }
}
