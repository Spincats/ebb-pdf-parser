/**
 * Line-based extraction of course, anonymous ID, and multiple-choice answers from concatenated
 * exam PDF text. Each MC/TF response is taken as a single character from the exam line (format "n. X");
 * the Excel workbook may later define multi-letter correct keys in row 2 (e.g. AB) for grading.
 *
 * pdf.js merged text (see pdfText.js) often inserts runs of spaces between words; marker and
 * MC-line matching tolerates that while preserving the same line-state behavior as the original
 * Python extractor.
 */

/** Collapses runs of whitespace to a single space for canonical marker comparisons. */
function collapseSpaces(line) {
  return line.replace(/\s+/g, " ").trim();
}

const MC_START = "--> Multiple Choice Answers";
const MC_END = "--->End Multiple Choice";

/**
 * @param {string} line processLine output
 * @returns {{ q: string, ans: string } | null}
 */
function parseMcAnswerLine(line) {
  const dot = line.indexOf(".", 0, 3);
  if (dot === -1) return null;
  const q = line.slice(0, dot);
  if (!/^\d+$/.test(q)) return null;
  const afterDot = line.slice(dot + 1).trimStart();
  if (afterDot.length === 0) return null;
  let ans = afterDot.charAt(0);
  if (ans === "T" || ans === "t") ans = "A";
  else if (ans === "F" || ans === "f") ans = "B";
  return { q, ans };
}

/**
 * Detects end-of-question markers (tolerates extra spaces from pdf.js).
 * @param {string} line processLine output
 */
function matchesEndOfQuestion(line) {
  const n = collapseSpaces(line);
  return n.includes("->End of Question") || /->\s*End\s+of\s+Question/.test(n);
}

/**
 * Matches "-->Question" with optional spaces from pdf.js (e.g. "-->   Question -1-").
 * @param {string} line processLine output
 */
function matchesQuestionStart(line) {
  return /-->\s*Question/.test(collapseSpaces(line));
}

/**
 * @param {string} line processLine output
 * @returns {string | null}
 */
function tryParseAnonymousNumber(line) {
  const m = line.match(/^\s*Anonymous\s+Number\s*:\s*(.+)$/i);
  if (!m) return null;
  return m[1].trim();
}

/**
 * @param {string} examText
 * @returns {{ course: string, exam_no: string, mc: Record<string, string> }}
 */
export function parseExamText(examText) {
  let inMc = false;
  let examNo = "";
  let course = "";
  /** @type {Record<string, string>} */
  const mc = {};

  const lines = examText.split(/\r?\n/);

  function processLine(line) {
    const blocked = line.indexOf("BLOCKED");
    if (blocked > -1) {
      return line.slice(0, blocked).trim();
    }
    return line.trim();
  }

  for (const raw of lines) {
    const line = processLine(raw);
    if (inMc) {
      const normMc = collapseSpaces(line);
      if (normMc.startsWith(MC_END)) {
        inMc = false;
        break;
      }
      const row = parseMcAnswerLine(line);
      if (row) mc[row.q] = row.ans;
    }
    if (course === "") {
      const pound = line.indexOf("#");
      if (pound === -1) continue;
      const nextComma = line.indexOf(",", pound);
      if (nextComma === -1) continue;
      course = line.slice(pound + 1, nextComma).trimStart();
      continue;
    }
    if (examNo === "") {
      const anon = tryParseAnonymousNumber(line);
      if (anon !== null) {
        examNo = anon;
        continue;
      }
    }
    // Suffix after the label is allowed (e.g. "Answers PM" from some scan exports).
    if (collapseSpaces(line).startsWith(MC_START)) {
      inMc = true;
      continue;
    }
  }

  const mcValues = Object.values(mc);
  if (mcValues.length > 0 && new Set(mcValues).size === 1) {
    inMc = false;
    let inQ = false;
    for (const raw of lines) {
      const line = processLine(raw);
      if (inMc) {
        const dot = line.indexOf(".", 0, 3);
        if (dot === -1) {
          if (matchesEndOfQuestion(line)) {
            inQ = false;
            inMc = false;
            break;
          }
          continue;
        }
        const row = parseMcAnswerLine(line);
        if (row) mc[row.q] = row.ans;
      }
      if (inQ && (line.indexOf("ultiple") > 0 || line.indexOf("MC") > -1)) {
        inMc = true;
        continue;
      }
      if (inQ && matchesEndOfQuestion(line)) {
        inQ = false;
        if (inMc) {
          inMc = false;
          break;
        }
        continue;
      }
      if (!inQ && matchesQuestionStart(line)) {
        inQ = true;
        continue;
      }
    }
  }

  return { course, exam_no: examNo, mc };
}

/**
 * Reads "Number of Essay Questions:" from page-1 text (merged lines).
 * Returns null if missing or unparsable.
 * @param {string} page1Text
 * @returns {number | null}
 */
export function extractEssayQuestionCount(page1Text) {
  if (!page1Text) return null;
  for (const raw of page1Text.split(/\r?\n/)) {
    const line = raw.trim();
    const re = /Number\s+of\s+Essay\s+Questions:\s+(\d+)/i;
    const m = line.match(re);
    if (m) {
      const n = parseInt(m[1], 10);
      if (Number.isFinite(n) && n >= 0) return n;
    }
  }
  return null;
}
