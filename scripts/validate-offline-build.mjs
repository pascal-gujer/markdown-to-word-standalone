import { readFile } from "node:fs/promises";
import path from "node:path";
import { fileURLToPath } from "node:url";

const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");
const htmlPath = path.join(root, "dist", "md-to-docx.html");
const allowMissing = process.argv.includes("--if-present");

let html;
try {
  html = await readFile(htmlPath, "utf8");
} catch (error) {
  if (allowMissing && error && error.code === "ENOENT") {
    console.log("Offline build validation skipped because dist/md-to-docx.html is not present.");
    process.exit(0);
  }
  throw error;
}

const checks = [
  { name: "raw http:// URL", pattern: /http:\/\//i },
  { name: "raw https:// URL", pattern: /https:\/\//i },
  { name: "protocol-relative URL", pattern: /(?:src|href)=["']\/\//i },
  { name: "external script tag", pattern: /<script\b[^>]*\bsrc=/i },
  { name: "external stylesheet tag", pattern: /<link\b[^>]*\brel=["']stylesheet["'][^>]*\bhref=/i },
  { name: "CSS import", pattern: /@import\s+/i },
  { name: "module script", pattern: /<script\b[^>]*\btype=["']module["']/i },
  { name: "dynamic import", pattern: /\bimport\s*\(/i },
  { name: "static import", pattern: /\bimport\s+[^("]/i },
  { name: "known CDN reference", pattern: /\b(?:cdn|unpkg|jsdelivr|cdnjs)\b/i },
];

const failures = checks
  .map((check) => ({ ...check, match: html.match(check.pattern) }))
  .filter((check) => check.match);

if (failures.length) {
  console.error("Offline build validation failed:");
  for (const failure of failures) {
    const index = failure.match?.index ?? 0;
    const snippet = html.slice(Math.max(0, index - 80), index + 120).replace(/\s+/g, " ");
    console.error(`- ${failure.name}: ${snippet}`);
  }
  process.exit(1);
}

console.log("Offline build validation passed: dist/md-to-docx.html has no external runtime references.");
