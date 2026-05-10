import { existsSync, readFileSync } from "node:fs";
import path from "node:path";
import { describe, expect, it } from "vitest";

describe("offline build artifact", () => {
  it("contains no external runtime references when dist exists", () => {
    const htmlPath = path.resolve("dist", "md-to-docx.html");
    if (!existsSync(htmlPath)) {
      return;
    }

    const html = readFileSync(htmlPath, "utf8");
    expect(html).not.toMatch(/https?:\/\//i);
    expect(html).not.toMatch(/<script\b[^>]*\bsrc=/i);
    expect(html).not.toMatch(/<link\b[^>]*\brel=["']stylesheet["'][^>]*\bhref=/i);
    expect(html).not.toMatch(/<script\b[^>]*\btype=["']module["']/i);
    expect(html).not.toMatch(/\bimport\s*\(/i);
  });
});
