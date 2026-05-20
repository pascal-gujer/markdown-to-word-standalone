import { describe, expect, it } from "vitest";
import JSZip from "jszip";
import { readImageDimensions } from "../src/markdown/imageBundle";
import { readMarkdownZipFile } from "../src/ui/zipImport";

describe("ZIP Markdown import", () => {
  it("loads Markdown and supported referenced images from a ZIP file", async () => {
    const zip = new JSZip();
    zip.file("report.md", [
      "# Report",
      "",
      "![Logo](images/logo.png)",
      "![Photo](images/photo.jpg)",
      "![Badge](images/badge.gif)",
      "![Missing](images/missing.jpg)",
      "![Unsupported](images/vector.svg)",
    ].join("\n"));
    zip.file("notes.txt", "This is not the preferred Markdown file.");
    zip.file("images/logo.png", tinyPngBytes());
    zip.file("images/photo.jpg", minimalJpegBytes(3, 2));
    zip.file("images/badge.gif", tinyGifBytes());
    zip.file("images/vector.svg", "<svg/>");

    const imported = await readMarkdownZipFile(await zipFile(zip, "bundle.zip"));

    expect(imported.markdownPath).toBe("report.md");
    expect(imported.markdownFileCount).toBe(2);
    expect(imported.markdown).toContain("# Report");
    expect(imported.imageBundle.assets).toHaveLength(3);
    expect(imported.imageBundle.assets[0]).toMatchObject({
      path: "images/logo.png",
      contentType: "image/png",
      widthPx: 1,
      heightPx: 1,
    });
    expect(imported.imageBundle.assets[1]).toMatchObject({
      path: "images/photo.jpg",
      contentType: "image/jpeg",
      widthPx: 3,
      heightPx: 2,
    });
    expect(imported.imageBundle.assets[2]).toMatchObject({
      path: "images/badge.gif",
      contentType: "image/gif",
      widthPx: 1,
      heightPx: 1,
    });
    expect(imported.missingReferences).toEqual(["images/missing.jpg"]);
    expect(imported.unsupportedReferences).toEqual(["images/vector.svg"]);
  });

  it("rejects ZIP files without Markdown content", async () => {
    const zip = new JSZip();
    zip.file("images/logo.png", tinyPngBytes());

    await expect(readMarkdownZipFile(await zipFile(zip, "images-only.zip"))).rejects.toThrow("error.zip_no_markdown");
  });

  it("reads PNG, GIF, and JPEG dimensions without browser image decoding", () => {
    expect(readImageDimensions(tinyPngBytes(), "png")).toEqual({ widthPx: 1, heightPx: 1 });
    expect(readImageDimensions(tinyGifBytes(), "gif")).toEqual({ widthPx: 1, heightPx: 1 });
    expect(readImageDimensions(minimalJpegBytes(3, 2), "jpg")).toEqual({ widthPx: 3, heightPx: 2 });
  });
});

async function zipFile(zip: JSZip, name: string): Promise<File> {
  const bytes = await zip.generateAsync({ type: "uint8array" });
  const copy = new ArrayBuffer(bytes.byteLength);
  new Uint8Array(copy).set(bytes);
  return new File([copy], name, { type: "application/zip" });
}

function tinyPngBytes(): Uint8Array {
  return Uint8Array.from([
    137, 80, 78, 71, 13, 10, 26, 10, 0, 0, 0, 13, 73, 72, 68, 82,
    0, 0, 0, 1, 0, 0, 0, 1, 8, 6, 0, 0, 0, 31, 21, 196,
    137, 0, 0, 0, 13, 73, 68, 65, 84, 120, 156, 99, 248, 255, 255,
    63, 0, 5, 254, 2, 254, 167, 53, 129, 132, 0, 0, 0, 0,
    73, 69, 78, 68, 174, 66, 96, 130,
  ]);
}

function tinyGifBytes(): Uint8Array {
  return Uint8Array.from([
    71, 73, 70, 56, 57, 97, 1, 0, 1, 0, 128, 0, 0, 0, 0, 0,
    255, 255, 255, 44, 0, 0, 0, 0, 1, 0, 1, 0, 0, 2, 1, 76,
    0, 59,
  ]);
}

function minimalJpegBytes(width: number, height: number): Uint8Array {
  return Uint8Array.from([
    0xff, 0xd8,
    0xff, 0xe0, 0x00, 0x10, 0x4a, 0x46, 0x49, 0x46, 0x00, 0x01, 0x01, 0x00, 0x00, 0x01, 0x00, 0x01, 0x00, 0x00,
    0xff, 0xc0, 0x00, 0x11, 0x08, (height >> 8) & 0xff, height & 0xff, (width >> 8) & 0xff, width & 0xff,
    0x03, 0x01, 0x11, 0x00, 0x02, 0x11, 0x00, 0x03, 0x11, 0x00,
    0xff, 0xd9,
  ]);
}
