import { existsSync, readFileSync } from "node:fs";
import path from "node:path";
import { describe, expect, it } from "vitest";
import { findOfflineBuildIssues, neuterAnchorHrefs } from "../scripts/validate-offline-build.mjs";

describe("offline build artifact", () => {
  it("contains no external runtime references when dist exists", () => {
    const htmlPath = path.resolve("dist", "md-to-docx.html");
    if (!existsSync(htmlPath)) {
      return;
    }

    const html = readFileSync(htmlPath, "utf8");
    expect(findOfflineBuildIssues(html)).toEqual([]);
  });

  it("ignores anchor href values when scanning for remote URLs", () => {
    const issues = findOfflineBuildIssues(
      `<!doctype html><html><body><a href="https://github.com/sponsors/pascal-gujer">Sponsor</a></body></html>`,
    );
    expect(issues).toEqual([]);
  });

  it("still flags non-anchor remote references", () => {
    const issues = findOfflineBuildIssues(
      `<!doctype html><html><body><img src="https://example.invalid/logo.png"></body></html>`,
    );
    expect(issues.map((issue) => issue.name)).toContain("raw https:// URL");
  });

  it("blanks out anchor href values without affecting surrounding markup", () => {
    const cleaned = neuterAnchorHrefs(`<a href="https://example.com" class="x">link</a><img src="https://example.invalid/a.png">`);
    expect(cleaned).toContain('<a href="" class="x">link</a>');
    expect(cleaned).toContain('<img src="https://example.invalid/a.png">');
  });
});
