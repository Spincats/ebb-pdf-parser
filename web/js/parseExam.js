/**
 * Ports mcparse() from mc_parse.py: line-based extraction of course, anonymous ID,
 * and multiple-choice answers from concatenated exam PDF text.
 * Each MC/TF response is taken as a single character from the exam line (format "n. X");
 * the Excel workbook may later define multi-letter correct keys in row 2 (e.g. AB) for grading.
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
      if (line === "--->End Multiple Choice") {
        inMc = false;
        break;
      }
      const dot = line.indexOf(".", 0, 3);
      if (dot === -1) continue;
      const q = line.slice(0, dot);
      if (!/^\d+$/.test(q)) continue;
      let ans = line.slice(dot + 2, dot + 3).trim();
      if (ans === "T" || ans === "t") ans = "A";
      if (ans === "F" || ans === "f") ans = "B";
      mc[q] = ans;
    }
    if (course === "") {
      const pound = line.indexOf("#");
      if (pound === -1) continue;
      const nextComma = line.indexOf(",", pound);
      if (nextComma === -1) continue;
      course = line.slice(pound + 1, nextComma).trimStart();
      continue;
    }
    if (examNo === "" && line.slice(0, 17) === "Anonymous Number:") {
      examNo = line.slice(18).trimStart();
      continue;
    }
    if (line === "--> Multiple Choice Answers") {
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
          if (line.indexOf("->End of Question") > -1) {
            inQ = false;
            inMc = false;
            break;
          }
          continue;
        }
        const q = line.slice(0, dot);
        if (!/^\d+$/.test(q)) continue;
        let ans = line.slice(dot + 2, dot + 3).trim();
        if (ans === "T" || ans === "t") ans = "A";
        if (ans === "F" || ans === "f") ans = "B";
        mc[q] = ans;
      }
      if (inQ && (line.indexOf("ultiple") > 0 || line.indexOf("MC") > -1)) {
        inMc = true;
        continue;
      }
      if (inQ && line.indexOf("->End of Question") > -1) {
        inQ = false;
        if (inMc) {
          inMc = false;
          break;
        }
        continue;
      }
      if (!inQ && line.indexOf("-->Question") > -1) {
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
