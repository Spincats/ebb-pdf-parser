/**
 * PDF text extraction using vendored pdf.js (ESM).
 * @param {typeof import("../vendor/pdfjs/pdf.mjs")} pdfjsModule
 */

const LINE_Y_THRESHOLD = 4;

/**
 * Merges pdf.js text items into reading-order lines (approximate).
 * @param {import("../vendor/pdfjs/pdf.mjs").TextContent} textContent
 * @returns {string}
 */
export function textContentToLines(textContent) {
  const items = textContent.items.filter(
    (it) => "str" in it && it.str && it.transform
  );
  /** @type {{ str: string, x: number, y: number, h: number }[]} */
  const placed = [];
  for (const it of items) {
    const tr = it.transform;
    const x = tr[4];
    const y = tr[5];
    const h = Math.abs(tr[3]) || 10;
    placed.push({ str: it.str, x, y, h });
  }
  placed.sort((a, b) => {
    if (Math.abs(a.y - b.y) > LINE_Y_THRESHOLD) return b.y - a.y;
    return a.x - b.x;
  });

  const lines = [];
  let cur = [];
  let lastY = null;
  for (const p of placed) {
    if (lastY !== null && Math.abs(p.y - lastY) > LINE_Y_THRESHOLD) {
      if (cur.length) lines.push(cur.map((c) => c.str).join(" "));
      cur = [];
    }
    cur.push(p);
    lastY = p.y;
  }
  if (cur.length) lines.push(cur.map((c) => c.str).join(" "));
  return lines.join("\n");
}

/**
 * @param {import("../vendor/pdfjs/pdf.mjs").PDFDocumentProxy} pdf
 * @returns {Promise<string>}
 */
export async function extractFullText(pdf) {
  let out = "";
  for (let i = 1; i <= pdf.numPages; i++) {
    const page = await pdf.getPage(i);
    const tc = await page.getTextContent();
    out += textContentToLines(tc);
    out += "\n";
  }
  return out;
}

/**
 * First page only (1-based page index 1).
 * @param {import("../vendor/pdfjs/pdf.mjs").PDFDocumentProxy} pdf
 * @returns {Promise<string>}
 */
export async function extractPage1Text(pdf) {
  const page = await pdf.getPage(1);
  const tc = await page.getTextContent();
  return textContentToLines(tc);
}

/**
 * Loads a PDF from ArrayBuffer and returns text for page 1 and full document.
 * @param {ArrayBuffer} data
 * @param {typeof import("../vendor/pdfjs/pdf.mjs")} pdfjs
 * @param {string} workerSrc URL for pdf.worker
 * @returns {Promise<{ page1Text: string, fullText: string }>}
 */
export async function extractExamPdfText(data, pdfjs, workerSrc) {
  pdfjs.GlobalWorkerOptions.workerSrc = workerSrc;
  const loadingTask = pdfjs.getDocument({ data });
  const pdf = await loadingTask.promise;
  try {
    const page1Text = await extractPage1Text(pdf);
    const fullText = await extractFullText(pdf);
    return { page1Text, fullText };
  } finally {
    if (typeof pdf.destroy === "function") {
      await pdf.destroy();
    }
  }
}
