/**
 * Copies pdf.js and ExcelJS build artifacts into web/vendor/ for static hosting.
 */
import { copyFile, mkdir, readdir, stat } from "node:fs/promises";
import { join, dirname } from "node:path";
import { fileURLToPath } from "node:url";

const __dirname = dirname(fileURLToPath(import.meta.url));
const webRoot = join(__dirname, "..");
const nm = join(webRoot, "node_modules");

async function copyPdfJs() {
  const srcDir = join(nm, "pdfjs-dist", "build");
  const destDir = join(webRoot, "vendor", "pdfjs");
  await mkdir(destDir, { recursive: true });
  const files = ["pdf.mjs", "pdf.worker.mjs", "pdf.min.mjs", "pdf.worker.min.mjs"];
  for (const f of files) {
    await copyFile(join(srcDir, f), join(destDir, f));
  }
}

async function copyExcelJs() {
  const destDir = join(webRoot, "vendor", "exceljs");
  await mkdir(destDir, { recursive: true });
  await copyFile(
    join(nm, "exceljs", "dist", "exceljs.min.js"),
    join(destDir, "exceljs.min.js")
  );
}

await copyPdfJs();
await copyExcelJs();
console.log("Vendored pdfjs + exceljs into web/vendor/");
