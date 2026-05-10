import { build } from "esbuild";
import { mkdir, readFile, rm, writeFile } from "node:fs/promises";
import path from "node:path";
import { fileURLToPath } from "node:url";

const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");
const outDir = path.join(root, "dist");
const tempDir = path.join(root, ".build");
const tempJs = path.join(tempDir, "app.js");
const sourceHtml = path.join(root, "src", "index.html");
const sourceCss = path.join(root, "src", "styles.css");
const outputHtml = path.join(outDir, "md-to-docx.html");

await rm(tempDir, { recursive: true, force: true });
await mkdir(tempDir, { recursive: true });
await mkdir(outDir, { recursive: true });

await build({
  entryPoints: [path.join(root, "src", "main.ts")],
  outfile: tempJs,
  bundle: true,
  minify: true,
  sourcemap: false,
  platform: "browser",
  target: ["es2022"],
  format: "iife",
  legalComments: "none",
});

const [htmlTemplate, css, bundledJs] = await Promise.all([
  readFile(sourceHtml, "utf8"),
  readFile(sourceCss, "utf8"),
  readFile(tempJs, "utf8"),
]);

const csp = [
  "default-src 'none'",
  "script-src 'unsafe-inline'",
  "style-src 'unsafe-inline'",
  "img-src data: blob:",
  "font-src data:",
  "connect-src 'none'",
  "media-src blob: data:",
  "worker-src 'none'",
  "child-src 'none'",
  "form-action 'none'",
  "frame-ancestors 'none'",
  "base-uri 'none'",
].join("; ");

let finalHtml = htmlTemplate
  .replace(/<link\s+rel="stylesheet"\s+href="\/src\/styles\.css"\s*\/?>/, () => `<style>\n${css}\n</style>`)
  .replace(/<script\s+type="module"\s+src="\/src\/main\.ts"><\/script>/, () => `<script>\n${escapeRawProtocols(bundledJs)}\n</script>`)
  .replace(
    /<meta name="viewport" content="width=device-width, initial-scale=1.0" \/>/,
    `<meta name="viewport" content="width=device-width, initial-scale=1.0" />\n    <meta name="referrer" content="no-referrer" />\n    <meta http-equiv="Content-Security-Policy" content="${csp}" />`,
  );

finalHtml = finalHtml
  .replace(/\n\s*\n\s*\n/g, "\n\n")
  .replace(/^\s+$/gm, "");

await writeFile(outputHtml, finalHtml, "utf8");
await rm(tempDir, { recursive: true, force: true });

console.log(`Built ${path.relative(root, outputHtml)}`);

function escapeRawProtocols(source) {
  return source
    .replace(/https:\/\//g, "https:\\x2f\\x2f")
    .replace(/http:\/\//g, "http:\\x2f\\x2f");
}
