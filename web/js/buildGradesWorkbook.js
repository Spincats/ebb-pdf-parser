import { colIndexToLetters, quoteSheetNameForFormula } from "./colUtils.js";
import {
  applyStudentCountToExams,
  validateMcTallyQuestionRanges,
} from "./workbookExport.js";
import {
  CUTOFFS_DATA_FIRST_ROW,
  CUTOFFS_DATA_LAST_ROW,
  CUTOFFS_HEADER_ROW,
  CUTOFFS_GPA_WITH_PART_COL,
  CUTOFFS_PART_NEG_LABEL_ROW,
  CUTOFFS_PART_NEG_VALUE_ROW,
  CUTOFFS_PART_POS_LABEL_ROW,
  CUTOFFS_PART_POS_VALUE_ROW,
  CUTOFFS_TOTAL_N_LABEL_ROW,
  CUTOFFS_TOTAL_N_VALUE_ROW,
  F25_GRADE_CUTOFF_ROWS,
  FIRST_STUDENT_ROW,
  HEADER_ROW,
  MC_ANSWER_KEY_ROW,
  MC_SUMMARY_GAP_ROWS,
  MC_CONFIG_INPUT_ROW,
  MC_CONFIG_MC_FIRST_Q_COL,
  MC_CONFIG_MC_LAST_Q_COL,
  MC_CONFIG_MC_PTS_COL,
  MC_CONFIG_TF_FIRST_Q_COL,
  MC_CONFIG_TF_LAST_Q_COL,
  MC_CONFIG_TF_PTS_COL,
  MC_TRAILER_HEADERS,
  SA_TRAILER_HEADERS,
  SHEET_CUTOFFS,
  SHEET_MC,
  SHEET_MC_TALLY,
  SHEET_MC_TF_CONFIG,
  SHEET_SA,
  SHEET_TOTAL,
  TITLE_ROW,
  TOTAL_LETTER_COL,
} from "./gradesSpec.js";

/** Title band fill aligned with F25 MC sheet header styling. */
const F25_TITLE_FILL = {
  type: "pattern",
  pattern: "solid",
  fgColor: { argb: "FFEFEFEF" },
  bgColor: { argb: "FFF3F3F3" },
};

const DATA_GRID_BORDER = {
  top: { style: "thin", color: { argb: "FFB4B4B4" } },
  left: { style: "thin", color: { argb: "FFB4B4B4" } },
  bottom: { style: "thin", color: { argb: "FFB4B4B4" } },
  right: { style: "thin", color: { argb: "FFB4B4B4" } },
};

const HEADER_BAND_FILL = {
  type: "pattern",
  pattern: "solid",
  fgColor: { argb: "FFD9E1F2" },
  bgColor: { argb: "FFD9E1F2" },
};

/**
 * Borders and optional header-row styling for a rectangular data block (no Excel ListObject).
 *
 * @param {import("exceljs").Worksheet} ws
 * @param {number} r1
 * @param {number} c1
 * @param {number} r2
 * @param {number} c2
 * @param {{ headerRow?: number }} [opts] 1-based row number treated as column headers
 */
function styleDataGrid(ws, r1, c1, r2, c2, opts = {}) {
  const headerRow = opts.headerRow;
  for (let r = r1; r <= r2; r++) {
    for (let c = c1; c <= c2; c++) {
      const cell = ws.getCell(r, c);
      cell.border = DATA_GRID_BORDER;
      if (headerRow !== undefined && r === headerRow) {
        cell.font = { ...(cell.font ?? {}), bold: true };
        cell.fill = HEADER_BAND_FILL;
      }
    }
  }
}

/**
 * Collects non-empty student answers for one MC question (trimmed strings).
 * @param {string} qid
 * @param {string[]} studentIds
 * @param {Map<string, Record<string, string>>} exams
 * @returns {string[]}
 */
export function collectAnswersForQuestion(qid, studentIds, exams) {
  const out = [];
  for (const id of studentIds) {
    const mc = exams.get(id);
    const raw = mc?.[qid];
    if (raw === undefined || raw === null) continue;
    const v = String(raw).trim();
    if (v !== "") out.push(v);
  }
  return out;
}

/**
 * Distinct answer choices for one question, sorted for display (single letters A-Z first, then other tokens).
 * @param {string} qid
 * @param {string[]} studentIds
 * @param {Map<string, Record<string, string>>} exams
 * @returns {string[]}
 */
export function distinctSortedAnswersForQuestion(qid, studentIds, exams) {
  const vals = collectAnswersForQuestion(qid, studentIds, exams);
  return sortMcAnswerChoices([...new Set(vals)]);
}

/**
 * @param {string[]} keys
 * @returns {string[]}
 */
export function sortMcAnswerChoices(keys) {
  return [...keys].filter(Boolean).sort((a, b) => {
    const letterA = /^[A-Za-z]$/.test(a);
    const letterB = /^[A-Za-z]$/.test(b);
    if (letterA !== letterB) return letterA ? -1 : 1;
    if (letterA && letterB) {
      return a.toUpperCase().charCodeAt(0) - b.toUpperCase().charCodeAt(0);
    }
    return a.localeCompare(b, undefined, { numeric: true, sensitivity: "base" });
  });
}

/**
 * True when every non-empty submitted answer for the question is a single letter A or B only
 * (omit C-E summary proportions for that column, matching F25 TF-style columns).
 *
 * @param {string} qid
 * @param {string[]} studentIds
 * @param {Map<string, Record<string, string>>} exams
 * @returns {boolean}
 */
export function isTfStyleMcColumn(qid, studentIds, exams) {
  const vals = collectAnswersForQuestion(qid, studentIds, exams);
  if (vals.length === 0) return false;
  for (const raw of vals) {
    const ch = String(raw).trim().charAt(0).toUpperCase();
    if (ch !== "A" && ch !== "B") return false;
  }
  return true;
}

/**
 * Excel formula: TRUE if the student's answer matches the key in row {@link MC_ANSWER_KEY_ROW}
 * of the same column. Key may contain several letters (e.g. AB); each letter is one accepted
 * answer (case-insensitive). Student answers are treated as a single leading letter via LEFT(TRIM(...),1).
 * Empty key or blank student cell yields FALSE.
 *
 * @param {number} mcCol1Based Excel column index (1-based) for both key and student answer
 * @param {number} studentRow 1-based row index for the student answer cell
 * @returns {string}
 */
export function buildMcKeyMatchFormula(mcCol1Based, studentRow) {
  return buildMcKeyMatchFormulaForSheet("", mcCol1Based, studentRow);
}

/**
 * Same as buildMcKeyMatchFormula but with optional `sheetPrefix` like `'MC+TF'!` for cross-sheet refs.
 * @param {string} sheetPrefix empty or `'MC+TF'!`
 */
export function buildMcKeyMatchFormulaForSheet(sheetPrefix, mcCol1Based, studentRow) {
  const colL = colIndexToLetters(mcCol1Based);
  const keyRef = `${sheetPrefix}${colL}$${MC_ANSWER_KEY_ROW}`;
  const ansRef = `${sheetPrefix}${colL}${studentRow}`;
  return (
    `IF(OR(LEN(TRIM(${keyRef}))=0,LEN(TRIM(${ansRef}))=0),FALSE,` +
    `ISNUMBER(SEARCH(LEFT(TRIM(${ansRef}),1),${keyRef})))`
  );
}

/**
 * @typedef {object} BuildGradesWorkbookParams
 * @property {string} reconciledCourse
 * @property {number} nSa short-answer / essay column count
 * @property {Map<string, Record<string, string>>} exams anonymous ID to MC answers
 * @property {number | null | undefined} [studentCount] when set, trim or pad rows (see {@link applyStudentCountToExams})
 * @property {number | null | undefined} [tallyMcFirst] MCtally MC Q first (1-based); default 1 or 0 if no questions
 * @property {number | null | undefined} [tallyMcLast] MCtally MC Q last (1-based); default nMc
 * @property {number | null | undefined} [tallyTfFirst] MCtally TF Q first; default 0 (TF off)
 * @property {number | null | undefined} [tallyTfLast] MCtally TF Q last; default 0
 */

/**
 * @param {typeof import("exceljs").default | import("exceljs").default} ExcelJSMod
 * @param {BuildGradesWorkbookParams} params
 * @returns {Promise<Blob>}
 */
export async function buildGradesWorkbook(params, ExcelJSMod = null) {
  const ExcelJS = ExcelJSMod ?? globalThis.ExcelJS;
  if (!ExcelJS) {
    throw new Error("ExcelJS is not loaded (expected global ExcelJS from vendor script).");
  }

  const {
    reconciledCourse,
    nSa,
    exams: examsInput,
    studentCount,
    tallyMcFirst: inMcFirst,
    tallyMcLast: inMcLast,
    tallyTfFirst: inTfFirst,
    tallyTfLast: inTfLast,
  } = params;

  const { exams, studentIds } = applyStudentCountToExams(examsInput, studentCount);
  const numStudents = studentIds.length;

  /** @type {Set<string>} */
  const mcKeySet = new Set();
  for (const id of studentIds) {
    const mc = exams.get(id);
    if (mc) for (const k of Object.keys(mc)) mcKeySet.add(k);
  }
  const mcKeys = [...mcKeySet].sort((a, b) => parseInt(a, 10) - parseInt(b, 10));
  const nMc = mcKeys.length;

  const tallyMcFirst =
    inMcFirst != null && Number.isFinite(inMcFirst)
      ? Math.floor(inMcFirst)
      : nMc > 0
        ? 1
        : 0;
  const tallyMcLast =
    inMcLast != null && Number.isFinite(inMcLast)
      ? Math.floor(inMcLast)
      : nMc > 0
        ? nMc
        : 0;
  let tallyTfFirst =
    inTfFirst != null && Number.isFinite(inTfFirst) ? Math.floor(inTfFirst) : 0;
  let tallyTfLast =
    inTfLast != null && Number.isFinite(inTfLast) ? Math.floor(inTfLast) : 0;
  if (tallyTfFirst <= 0 || tallyTfLast <= 0) {
    tallyTfFirst = 0;
    tallyTfLast = 0;
  }
  const rangeErr = validateMcTallyQuestionRanges(
    nMc,
    tallyMcFirst,
    tallyMcLast,
    tallyTfFirst,
    tallyTfLast
  );
  if (rangeErr) {
    throw new Error(rangeErr);
  }

  const wb = new ExcelJS.Workbook();
  wb.creator = "Exam parser web";
  wb.created = new Date();

  const wsSa = wb.addWorksheet(SHEET_SA, { views: [{ state: "frozen", ySplit: HEADER_ROW }] });
  const wsMc = wb.addWorksheet(SHEET_MC, { views: [{ state: "frozen", ySplit: HEADER_ROW }] });

  const lastStudentRow = FIRST_STUDENT_ROW + numStudents - 1;
  const lastSaCol = 1 + nSa + SA_TRAILER_HEADERS.length;
  const lastMcCol = 1 + nMc + MC_TRAILER_HEADERS.length;

  buildSaSheet(wsSa, {
    reconciledCourse,
    nSa,
    studentIds,
    numStudents,
    lastStudentRow,
    lastSaCol,
  });

  buildMcSheet(wsMc, {
    reconciledCourse,
    mcKeys,
    nMc,
    studentIds,
    exams,
    numStudents,
    lastStudentRow,
    lastMcCol,
  });

  const mcQuoted = quoteSheetNameForFormula(SHEET_MC);
  const tallyQuoted = quoteSheetNameForFormula(SHEET_MC_TALLY);
  const configQuoted = quoteSheetNameForFormula(SHEET_MC_TF_CONFIG);

  const tallyWs = wb.addWorksheet(SHEET_MC_TALLY, {
    views: [{ state: "frozen", ySplit: HEADER_ROW }],
  });
  buildMcTallySheet(tallyWs, {
    mcKeys,
    nMc,
    studentIds,
    numStudents,
    lastStudentRow,
    mcQuoted,
    configQuoted,
    tallyMcFirst,
    tallyMcLast,
    tallyTfFirst,
    tallyTfLast,
  });

  const configWs = wb.addWorksheet(SHEET_MC_TF_CONFIG);
  buildMcTfConfigSheet(configWs, {
    nMc,
    tallyMcFirst,
    tallyMcLast,
    tallyTfFirst,
    tallyTfLast,
  });

  const cutWs = wb.addWorksheet(SHEET_CUTOFFS);
  const totalWs = wb.addWorksheet(SHEET_TOTAL);
  const saQuoted = quoteSheetNameForFormula(SHEET_SA);
  buildTotalSheet(totalWs, {
    numStudents,
    nMc,
    nSa,
    mcQuoted,
    tallyQuoted,
    saQuoted,
    configQuoted,
  });
  buildCutoffsSheet(cutWs);

  wb.definedNames.add(
    `${SHEET_CUTOFFS}!$A$${CUTOFFS_DATA_FIRST_ROW}:$B$${CUTOFFS_DATA_LAST_ROW}`,
    "cutoffs"
  );
  const letterColL = colIndexToLetters(TOTAL_LETTER_COL);
  wb.definedNames.add(
    `${SHEET_TOTAL}!$${letterColL}$${FIRST_STUDENT_ROW}:$${letterColL}$${lastStudentRow}`,
    "letters"
  );

  const buf = await wb.xlsx.writeBuffer();
  return new Blob([buf], {
    type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
  });
}

/**
 * @param {import("exceljs").Worksheet} ws
 */
function buildSaSheet(ws, p) {
  const { reconciledCourse, nSa, studentIds, numStudents, lastStudentRow, lastSaCol } = p;

  ws.mergeCells(TITLE_ROW, 1, TITLE_ROW, lastSaCol);
  const t = ws.getCell(TITLE_ROW, 1);
  t.value = `Short Answer + Essay — ${reconciledCourse}`;
  t.font = { bold: true, size: 14 };
  t.fill = F25_TITLE_FILL;
  t.alignment = { vertical: "middle", horizontal: "center" };

  ws.getCell(HEADER_ROW, 1).value = "Anonymous ID";
  ws.getCell(HEADER_ROW, 1).font = { bold: true };

  const lastSaScoreCol = nSa > 0 ? 1 + nSa : 1;
  const lastSaLetter = colIndexToLetters(lastSaScoreCol);

  if (nSa > 0) {
    for (let i = 1; i <= nSa; i++) {
      const c = 1 + i;
      ws.getCell(HEADER_ROW, c).value = String(i);
      ws.getCell(HEADER_ROW, c).font = { bold: true };
    }
    for (let r = 0; r < numStudents; r++) {
      const row = FIRST_STUDENT_ROW + r;
      for (let c = 2; c <= lastSaScoreCol; c++) {
        ws.getCell(row, c).value = null;
      }
    }
    styleDataGrid(ws, HEADER_ROW, 1, lastStudentRow, lastSaScoreCol, {
      headerRow: HEADER_ROW,
    });
  } else {
    ws.getCell(HEADER_ROW, 2).value = "(no SA columns)";
  }

  let c = 2 + nSa;
  for (const h of SA_TRAILER_HEADERS) {
    ws.getCell(HEADER_ROW, c).value = h;
    ws.getCell(HEADER_ROW, c).font = { bold: true };
    c++;
  }

  const totalTrailerIndex = SA_TRAILER_HEADERS.indexOf("TOTAL");
  const totalCol =
    totalTrailerIndex >= 0 ? 2 + nSa + totalTrailerIndex : null;

  for (let r = 0; r < numStudents; r++) {
    const row = FIRST_STUDENT_ROW + r;
    ws.getCell(row, 1).value = studentIds[r];
    for (let tc = 0; tc < SA_TRAILER_HEADERS.length; tc++) {
      const colIndex = 2 + nSa + tc;
      if (totalCol !== null && colIndex === totalCol && nSa > 0) {
        const totalF =
          nSa === 1
            ? `=B${row}`
            : `=SUM(B${row}:$${lastSaLetter}$${row})`;
        ws.getCell(row, colIndex).value = { formula: totalF };
      } else {
        ws.getCell(row, colIndex).value = null;
      }
    }
  }

  if (nSa > 0) {
    styleDataGrid(ws, HEADER_ROW, lastSaScoreCol + 1, lastStudentRow, lastSaCol, {
      headerRow: HEADER_ROW,
    });
  }

  ws.getColumn(1).width = 18;
  for (let i = 2; i <= lastSaCol; i++) {
    ws.getColumn(i).width = 10;
  }
}

/**
 * @param {import("exceljs").Worksheet} ws
 */
function buildMcSheet(ws, p) {
  const {
    reconciledCourse,
    mcKeys,
    nMc,
    studentIds,
    exams,
    numStudents,
    lastStudentRow,
    lastMcCol,
  } = p;

  ws.mergeCells(TITLE_ROW, 1, TITLE_ROW, lastMcCol);
  const t = ws.getCell(TITLE_ROW, 1);
  t.value = `Multiple choice / True-False — ${reconciledCourse}`;
  t.font = { bold: true, size: 14 };
  t.fill = F25_TITLE_FILL;
  t.alignment = { vertical: "middle", horizontal: "center" };

  ws.getCell(MC_ANSWER_KEY_ROW, 1).value =
    "Answer key (row 2, same column as answers): e.g. A or AB = accept A or B (case-insensitive).";
  ws.getCell(MC_ANSWER_KEY_ROW, 1).font = { italic: true, size: 9 };
  for (let qi = 0; qi < nMc; qi++) {
    const col = 2 + qi;
    ws.getCell(MC_ANSWER_KEY_ROW, col).value = null;
  }

  ws.getCell(HEADER_ROW, 1).value = "Anonymous ID";
  ws.getCell(HEADER_ROW, 1).font = { bold: true };

  if (nMc > 0) {
    for (let qi = 0; qi < mcKeys.length; qi++) {
      const col = 2 + qi;
      ws.getCell(HEADER_ROW, col).value = String(mcKeys[qi]);
      ws.getCell(HEADER_ROW, col).font = { bold: true };
    }
    for (let r = 0; r < numStudents; r++) {
      const row = FIRST_STUDENT_ROW + r;
      const id = studentIds[r];
      const mc = exams.get(id) ?? {};
      for (let qi = 0; qi < mcKeys.length; qi++) {
        const col = 2 + qi;
        ws.getCell(row, col).value = mc[mcKeys[qi]] ?? "";
      }
    }
  } else {
    ws.getCell(HEADER_ROW, 2).value = "(no MC columns)";
  }

  let c = 2 + nMc;
  for (const h of MC_TRAILER_HEADERS) {
    ws.getCell(HEADER_ROW, c).value = h;
    ws.getCell(HEADER_ROW, c).font = { bold: true };
    c++;
  }

  for (let r = 0; r < numStudents; r++) {
    const row = FIRST_STUDENT_ROW + r;
    ws.getCell(row, 1).value = studentIds[r];
    for (let tc = 0; tc < MC_TRAILER_HEADERS.length; tc++) {
      ws.getCell(row, 2 + nMc + tc).value = null;
    }
  }

  styleDataGrid(ws, HEADER_ROW, 1, lastStudentRow, lastMcCol, {
    headerRow: HEADER_ROW,
  });

  ws.getColumn(1).width = 18;
  for (let i = 2; i <= lastMcCol; i++) {
    ws.getColumn(i).width = i <= 1 + nMc ? 6 : 10;
  }

  if (nMc > 0) {
    writeMcSummaryGrid(ws, {
      mcKeys,
      studentIds,
      exams,
      lastStudentRow,
    });
    applyMcConditionalFormatting(ws, {
      mcKeys,
      studentIds,
      exams,
      lastStudentRow,
      nMc,
    });
  }
}

/**
 * F25-style horizontal summary: question row, CORRECT (= row-2 key), then A-E as share of
 * non-blank responses (COUNTIF/COUNTA). TF-style columns omit C-E formulas.
 *
 * @param {import("exceljs").Worksheet} ws
 */
function writeMcSummaryGrid(ws, ctx) {
  const { mcKeys, studentIds, exams, lastStudentRow } = ctx;
  const baseRow = lastStudentRow + MC_SUMMARY_GAP_ROWS;
  const letters = ["A", "B", "C", "D", "E"];

  for (let qi = 0; qi < mcKeys.length; qi++) {
    const qid = mcKeys[qi];
    const col = 2 + qi;
    ws.getCell(baseRow, col).value = qid;
    ws.getCell(baseRow, col).font = { bold: true };
  }

  const correctRow = baseRow + 1;
  ws.getCell(correctRow, 1).value = "CORRECT";
  ws.getCell(correctRow, 1).font = { bold: true };
  for (let qi = 0; qi < mcKeys.length; qi++) {
    const col = 2 + qi;
    const colL = colIndexToLetters(col);
    ws.getCell(correctRow, col).value = {
      formula: `=${colL}$${MC_ANSWER_KEY_ROW}`,
    };
  }

  for (let li = 0; li < letters.length; li++) {
    const letter = letters[li];
    const row = baseRow + 2 + li;
    ws.getCell(row, 1).value = letter;
    ws.getCell(row, 1).font = { bold: true };

    for (let qi = 0; qi < mcKeys.length; qi++) {
      const qid = mcKeys[qi];
      const col = 2 + qi;
      const colL = colIndexToLetters(col);
      const tf = isTfStyleMcColumn(qid, studentIds, exams);
      if (tf && (letter === "C" || letter === "D" || letter === "E")) {
        continue;
      }
      const aCell = `$A$${row}`;
      const dataTop = FIRST_STUDENT_ROW;
      const f = `IFERROR(COUNTIF($${colL}$${dataTop}:$${colL}$${lastStudentRow},${aCell})/COUNTA($${colL}$${dataTop}:$${colL}$${lastStudentRow}),"")`;
      const cell = ws.getCell(row, col);
      cell.value = { formula: f };
      cell.numFmt = "0.00%";
    }
  }
}

/**
 * MC+TF: answer-area highlight (correct vs incorrect vs blank) and color scale on A-E proportion rows.
 *
 * @param {import("exceljs").Worksheet} ws
 */
function applyMcConditionalFormatting(ws, ctx) {
  const { mcKeys, studentIds, exams, lastStudentRow, nMc } = ctx;
  if (nMc <= 0) return;

  const kr = MC_ANSWER_KEY_ROW;
  const fr = FIRST_STUDENT_ROW;
  const colB = colIndexToLetters(2);
  const lastAnsL = colIndexToLetters(1 + nMc);
  const refAns = `${colB}${fr}:${lastAnsL}${lastStudentRow}`;
  const correctF = `AND(LEN(TRIM(${colB}${fr}))>0,ISNUMBER(SEARCH(LEFT(TRIM(${colB}${fr}),1),${colB}$${kr})))`;
  const wrongF = `AND(LEN(TRIM(${colB}${fr}))>0,ISERROR(SEARCH(LEFT(TRIM(${colB}${fr}),1),${colB}$${kr})))`;

  ws.addConditionalFormatting({
    ref: refAns,
    rules: [
      {
        type: "expression",
        priority: 2,
        formulae: [correctF],
        style: {
          fill: {
            type: "pattern",
            pattern: "solid",
            bgColor: { argb: "FFB7E1CD" },
          },
        },
      },
      {
        type: "expression",
        priority: 3,
        formulae: [wrongF],
        style: {
          fill: {
            type: "pattern",
            pattern: "solid",
            bgColor: { argb: "FFF4C7C3" },
          },
        },
      },
    ],
  });

  const baseRow = lastStudentRow + MC_SUMMARY_GAP_ROWS;
  const firstLetRow = baseRow + 2;
  const summaryParts = [];
  for (let qi = 0; qi < mcKeys.length; qi++) {
    const colL = colIndexToLetters(2 + qi);
    const tf = isTfStyleMcColumn(mcKeys[qi], studentIds, exams);
    const lastR = tf ? baseRow + 3 : baseRow + 6;
    summaryParts.push(`${colL}${firstLetRow}:${colL}${lastR}`);
  }

  ws.addConditionalFormatting({
    ref: summaryParts.join(" "),
    rules: [
      {
        type: "colorScale",
        priority: 4,
        cfvo: [
          { type: "min", value: 0 },
          { type: "percentile", value: 50 },
          { type: "max", value: 0 },
        ],
        color: [
          { argb: "FFFFFFFF" },
          { argb: "FFFFFFFF" },
          { argb: "FF6FA8DC" },
        ],
      },
    ],
  });
}

/**
 * MCtally: per-question Ok formulas point at {@link SHEET_MC} keys and answers; column A is total correct;
 * columns M–R hold points-per-question and 1-based MC/TF question index ranges (inclusive) into the MC sheet
 * columns B..; MC pts / TF pts columns use SUMPRODUCT over the local Ok cells in those index ranges.
 * @param {import("exceljs").Worksheet} ws
 */
function buildMcTallySheet(ws, p) {
  const {
    mcKeys,
    nMc,
    numStudents,
    lastStudentRow,
    mcQuoted,
    tallyMcFirst,
    tallyMcLast,
    tallyTfFirst,
    tallyTfLast,
    configQuoted,
  } = p;

  ws.mergeCells(TITLE_ROW, 1, TITLE_ROW, 2 + Math.max(nMc, 2));
  const title = ws.getCell(TITLE_ROW, 1);
  title.value = "Multiple Choice and True-False scoring";
  title.font = { bold: true, size: 14 };
  title.fill = F25_TITLE_FILL;
  title.alignment = { vertical: "middle", horizontal: "center" };

  const ir = MC_CONFIG_INPUT_ROW;
  const mL = `${configQuoted}!$${colIndexToLetters(MC_CONFIG_MC_PTS_COL)}$${ir}`;
  const nL = `${configQuoted}!$${colIndexToLetters(MC_CONFIG_TF_PTS_COL)}$${ir}`;
  const oL = `${configQuoted}!$${colIndexToLetters(MC_CONFIG_MC_FIRST_Q_COL)}$${ir}`;
  const pL = `${configQuoted}!$${colIndexToLetters(MC_CONFIG_MC_LAST_Q_COL)}$${ir}`;
  const qL = `${configQuoted}!$${colIndexToLetters(MC_CONFIG_TF_FIRST_Q_COL)}$${ir}`;
  const rL = `${configQuoted}!$${colIndexToLetters(MC_CONFIG_TF_LAST_Q_COL)}$${ir}`;

  if (nMc > 0) {
    const mcPtsCol = 3 + nMc;
    const tfPtsCol = 4 + nMc;

    const firstOkL = colIndexToLetters(3);
    const lastOkL = colIndexToLetters(2 + nMc);

    ws.getCell(HEADER_ROW, 1).value = "Total correct";
    ws.getCell(HEADER_ROW, 1).font = { bold: true };
    ws.getCell(HEADER_ROW, 2).value = "Anonymous ID";
    ws.getCell(HEADER_ROW, 2).font = { bold: true };
    for (let qi = 0; qi < nMc; qi++) {
      const col = 3 + qi;
      ws.getCell(HEADER_ROW, col).value = String(mcKeys[qi]);
      ws.getCell(HEADER_ROW, col).font = { bold: true };
    }
    ws.getCell(HEADER_ROW, mcPtsCol).value = "MC pts";
    ws.getCell(HEADER_ROW, mcPtsCol).font = { bold: true };
    ws.getCell(HEADER_ROW, tfPtsCol).value = "TF pts";
    ws.getCell(HEADER_ROW, tfPtsCol).font = { bold: true };

    for (let r = 0; r < numStudents; r++) {
      const row = FIRST_STUDENT_ROW + r;
      const totalF = `=SUMPRODUCT(0+(${firstOkL}${row}:${lastOkL}${row}))`;
      const idF = `=${mcQuoted}!A${row}`;
      ws.getCell(row, 1).value = { formula: totalF };
      ws.getCell(row, 2).value = { formula: idF };
      for (let qi = 0; qi < nMc; qi++) {
        const mcCol = 2 + qi;
        ws.getCell(row, 3 + qi).value = {
          formula: buildMcKeyMatchFormulaForSheet(`${mcQuoted}!`, mcCol, row),
        };
      }
      const mcPtsF =
        `=${mL}*IF(OR(${oL}<1,${pL}<${oL},${oL}>${nMc},${pL}>${nMc}),0,` +
        `SUMPRODUCT(0+(OFFSET($A${row},0,(1+${oL}),1,${pL}-${oL}+1))))`;
      const tfPtsF =
        `=${nL}*IF(OR(${qL}=0,${rL}<${qL},${qL}>${nMc},${rL}>${nMc}),0,` +
        `SUMPRODUCT(0+(OFFSET($A${row},0,(1+${qL}),1,${rL}-${qL}+1))))`;
      ws.getCell(row, mcPtsCol).value = { formula: mcPtsF };
      ws.getCell(row, tfPtsCol).value = { formula: tfPtsF };
    }

    const lastColL = colIndexToLetters(tfPtsCol);
    styleDataGrid(ws, HEADER_ROW, 1, lastStudentRow, tfPtsCol, {
      headerRow: HEADER_ROW,
    });
    ws.getRow(1).height = 36;
  } else {
    ws.getCell(HEADER_ROW, 3).value = "(no MC questions)";
    ws.getCell(FIRST_STUDENT_ROW, 1).value =
      "Enter MC answers on MC+TF; this sheet populates when at least one question exists.";
  }

  ws.getColumn(1).width = 14;
  ws.getColumn(2).width = 16;
  for (let c = 3; c <= Math.max(3, 2 + Math.max(nMc, 1)); c++) {
    ws.getColumn(c).width = 8;
  }
  if (nMc > 0) {
    ws.getColumn(3 + nMc).width = 10;
    ws.getColumn(4 + nMc).width = 10;
  }
}

/**
 * @param {import("exceljs").Worksheet} ws
 */
function buildMcTfConfigSheet(ws, p) {
  const { nMc, tallyMcFirst, tallyMcLast, tallyTfFirst, tallyTfLast } = p;

  const paramLabelStyle = { size: 9, italic: true };
  const paramLabels = [
    [MC_CONFIG_MC_PTS_COL, "Points per MC question"],
    [MC_CONFIG_TF_PTS_COL, "Points per TF question"],
    [MC_CONFIG_MC_FIRST_Q_COL, "MC block first Q (1-based)"],
    [MC_CONFIG_MC_LAST_Q_COL, "MC block last Q (1-based)"],
    [MC_CONFIG_TF_FIRST_Q_COL, "TF block first Q (1-based)"],
    [MC_CONFIG_TF_LAST_Q_COL, "TF block last Q (1-based)"],
  ];

  for (const [col, text] of paramLabels) {
    const c = ws.getCell(1, col);
    c.value = text;
    c.font = paramLabelStyle;
    c.alignment = { wrapText: true, vertical: "top", horizontal: "center" };
  }

  const ir = MC_CONFIG_INPUT_ROW;
  ws.getCell(ir, MC_CONFIG_MC_PTS_COL).value = 1;
  ws.getCell(ir, MC_CONFIG_TF_PTS_COL).value = 1;
  ws.getCell(ir, MC_CONFIG_MC_FIRST_Q_COL).value = tallyMcFirst;
  ws.getCell(ir, MC_CONFIG_MC_LAST_Q_COL).value = tallyMcLast;
  ws.getCell(ir, MC_CONFIG_TF_FIRST_Q_COL).value = tallyTfFirst;
  ws.getCell(ir, MC_CONFIG_TF_LAST_Q_COL).value = tallyTfLast;

  ws.getRow(1).height = 36;
  for (let i = 1; i <= 6; i++) {
    ws.getColumn(i).width = 12;
  }
}

/**
 * Cutoffs: F25-style ladder (Value, Grade, GPA, Count, Pct), Total N, class GPA, and curve guidance (not enforced).
 * Count uses COUNTIF(letters, Grade); Pct uses Count / Total N. Depends on workbook name `letters` (set after Total).
 * @param {import("exceljs").Worksheet} ws
 */
function buildCutoffsSheet(ws) {
  const lastRow = CUTOFFS_DATA_LAST_ROW;
  const totalNRow = CUTOFFS_TOTAL_N_VALUE_ROW;
  const headerR = CUTOFFS_HEADER_ROW;
  ws.getCell(headerR, 1).value = "Cutoff";
  ws.getCell(headerR, 2).value = "Grade";
  ws.getCell(headerR, 3).value = "GPA";
  ws.getCell(headerR, 4).value = "Count";
  ws.getCell(headerR, 5).value = "Pct";
  for (let c = 1; c <= 5; c++) {
    ws.getCell(headerR, c).font = { bold: true };
  }

  for (let i = 0; i < F25_GRADE_CUTOFF_ROWS.length; i++) {
    const [minV, letter, gpa] = F25_GRADE_CUTOFF_ROWS[i];
    const row = CUTOFFS_DATA_FIRST_ROW + i;
    ws.getCell(row, 1).value = minV;
    ws.getCell(row, 1).numFmt = "0.00%";
    ws.getCell(row, 2).value = letter;
    ws.getCell(row, 3).value = gpa;
    ws.getCell(row, 4).value = { formula: `=COUNTIF(letters,B${row})` };
    ws.getCell(row, 5).value = { formula: `=D${row}/$B$${totalNRow}` };
  }

  styleDataGrid(ws, headerR, 1, lastRow, 5, { headerRow: headerR });

  ws.getCell(CUTOFFS_TOTAL_N_LABEL_ROW, 2).value = "Total N";
  ws.getCell(CUTOFFS_TOTAL_N_LABEL_ROW, 2).font = { bold: true };
  ws.getCell(CUTOFFS_TOTAL_N_VALUE_ROW, 2).value = { formula: "=COUNTA(letters)" };

  const gpaLabel = ws.getCell(5, 7);
  gpaLabel.value = "Class GPA (pre-participation)";
  gpaLabel.font = { bold: true, size: 10 };
  ws.getCell(6, 7).value = {
    formula: `=SUMPRODUCT($C$${CUTOFFS_DATA_FIRST_ROW}:$C$${lastRow},$D$${CUTOFFS_DATA_FIRST_ROW}:$D$${lastRow})/$B$${totalNRow}`,
  };

  const bMinusHint = ws.getCell(6, 6);
  bMinusHint.value = "Curve target: 5-20% at B-";
  bMinusHint.font = { italic: true, size: 9 };
  bMinusHint.alignment = { wrapText: true, vertical: "top" };

  const aHint = ws.getCell(10, 6);
  aHint.value = "Curve target: 5-30% at A";
  aHint.font = { italic: true, size: 9 };
  aHint.alignment = { wrapText: true, vertical: "top" };

  const gpaBandHint = ws.getCell(6, 8);
  gpaBandHint.value = "Curve target: 3.2-3.4 class GPA (pre-participation)";
  gpaBandHint.font = { italic: true, size: 9 };
  gpaBandHint.alignment = { wrapText: true, vertical: "top" };

  ws.mergeCells(CUTOFFS_PART_POS_LABEL_ROW, 2, CUTOFFS_PART_POS_LABEL_ROW, 5);
  const posLab = ws.getCell(CUTOFFS_PART_POS_LABEL_ROW, 2);
  posLab.value = "Participation Adj. Positive";
  posLab.font = { bold: true };
  posLab.alignment = { vertical: "middle", horizontal: "center" };
  const gpaPartHead = ws.getCell(CUTOFFS_PART_POS_LABEL_ROW, CUTOFFS_GPA_WITH_PART_COL);
  gpaPartHead.value = "GPA w/ Part";
  gpaPartHead.font = { bold: true };

  ws.getCell(CUTOFFS_PART_POS_VALUE_ROW, 2).value = 0;

  const sumGpaCountPairs = [];
  for (let r = CUTOFFS_DATA_FIRST_ROW; r <= lastRow; r++) {
    sumGpaCountPairs.push(`C${r}*D${r}`);
  }
  const gpaWithPartF =
    `=(SUM(${sumGpaCountPairs.join(",")},B${CUTOFFS_PART_POS_VALUE_ROW}/3,(-1*B${CUTOFFS_PART_NEG_VALUE_ROW})/3))/$B$${totalNRow}`;
  ws.getCell(CUTOFFS_PART_POS_VALUE_ROW, CUTOFFS_GPA_WITH_PART_COL).value = { formula: gpaWithPartF };

  ws.mergeCells(CUTOFFS_PART_NEG_LABEL_ROW, 2, CUTOFFS_PART_NEG_LABEL_ROW, 5);
  const negLab = ws.getCell(CUTOFFS_PART_NEG_LABEL_ROW, 2);
  negLab.value = "Participation Adj. Negative";
  negLab.font = { bold: true };
  negLab.alignment = { vertical: "middle", horizontal: "center" };
  ws.getCell(CUTOFFS_PART_NEG_VALUE_ROW, 2).value = 0;

  for (let r = CUTOFFS_DATA_FIRST_ROW; r <= lastRow; r++) {
    ws.getCell(r, 5).numFmt = "0.00%";
  }

  ws.getColumn(1).width = 10;
  ws.getColumn(2).width = 7;
  ws.getColumn(3).width = 7;
  ws.getColumn(4).width = 9;
  ws.getColumn(5).width = 10;
  ws.getColumn(6).width = 36;
  ws.getColumn(7).width = 22;
  ws.getColumn(8).width = 44;
}

/**
 * Total: one row per student with IDs from MC+TF, tallies from MCtally, SA total from SA sheet,
 * combined fraction of max (MC+TF + SA), and letter from Cutoffs.
 * @param {import("exceljs").Worksheet} ws
 */
function buildTotalSheet(ws, p) {
  const { numStudents, nMc, nSa, mcQuoted, tallyQuoted, saQuoted, configQuoted } = p;
  const lastCol = 6;
  ws.mergeCells(TITLE_ROW, 1, TITLE_ROW, lastCol);
  const title = ws.getCell(TITLE_ROW, 1);
  title.value = "Totals (MC, TF, SA, combined percent, letter grade)";
  title.font = { bold: true, size: 14 };
  title.fill = F25_TITLE_FILL;
  title.alignment = { vertical: "middle", horizontal: "center" };

  ws.getCell(2, 1).value = "Max MC+TF pts (from MCtally)";
  ws.getCell(2, 1).font = { italic: true, size: 10 };
  ws.getCell(2, 3).value = "Max SA pts";
  ws.getCell(2, 3).font = { italic: true, size: 10 };

  const cq = configQuoted;
  const ir = MC_CONFIG_INPUT_ROW;
  const mLc = colIndexToLetters(MC_CONFIG_MC_PTS_COL);
  const nLc = colIndexToLetters(MC_CONFIG_TF_PTS_COL);
  const oLc = colIndexToLetters(MC_CONFIG_MC_FIRST_Q_COL);
  const pLc = colIndexToLetters(MC_CONFIG_MC_LAST_Q_COL);
  const qLc = colIndexToLetters(MC_CONFIG_TF_FIRST_Q_COL);
  const rLc = colIndexToLetters(MC_CONFIG_TF_LAST_Q_COL);

  const maxPtsF =
    `=${cq}!$${mLc}$${ir}*IF(OR(${cq}!$${oLc}$${ir}<1,${cq}!$${pLc}$${ir}<${cq}!$${oLc}$${ir},${cq}!$${oLc}$${ir}>${nMc},${cq}!$${pLc}$${ir}>${nMc}),0,` +
    `${cq}!$${pLc}$${ir}-${cq}!$${oLc}$${ir}+1)+` +
    `${cq}!$${nLc}$${ir}*IF(OR(${cq}!$${qLc}$${ir}=0,${cq}!$${rLc}$${ir}<${cq}!$${qLc}$${ir},${cq}!$${qLc}$${ir}>${nMc},${cq}!$${rLc}$${ir}>${nMc}),0,` +
    `${cq}!$${rLc}$${ir}-${cq}!$${qLc}$${ir}+1)`;
  ws.getCell(2, 2).value = { formula: maxPtsF };
  ws.getCell(2, 4).value = nSa > 0 ? 150 : 0;

  ws.getCell(HEADER_ROW, 1).value = "Anonymous ID";
  ws.getCell(HEADER_ROW, 1).font = { bold: true };
  ws.getCell(HEADER_ROW, 2).value = "MC pts";
  ws.getCell(HEADER_ROW, 2).font = { bold: true };
  ws.getCell(HEADER_ROW, 3).value = "TF pts";
  ws.getCell(HEADER_ROW, 3).font = { bold: true };
  ws.getCell(HEADER_ROW, 4).value = "SA total";
  ws.getCell(HEADER_ROW, 4).font = { bold: true };
  ws.getCell(HEADER_ROW, 5).value = "Overall pct";
  ws.getCell(HEADER_ROW, 5).font = { bold: true };
  ws.getCell(HEADER_ROW, 6).value = "Letter";
  ws.getCell(HEADER_ROW, 6).font = { bold: true };

  const mcPtsColL = nMc > 0 ? colIndexToLetters(3 + nMc) : "C";
  const tfPtsColL = nMc > 0 ? colIndexToLetters(4 + nMc) : "D";
  const totalIdx = SA_TRAILER_HEADERS.indexOf("TOTAL");
  const saTotalColL =
    nSa > 0 && totalIdx >= 0
      ? colIndexToLetters(2 + nSa + totalIdx)
      : null;

  for (let r = 0; r < numStudents; r++) {
    const row = FIRST_STUDENT_ROW + r;
    ws.getCell(row, 1).value = { formula: `=${mcQuoted}!A${row}` };
    if (nMc > 0) {
      ws.getCell(row, 2).value = { formula: `=${tallyQuoted}!${mcPtsColL}${row}` };
      ws.getCell(row, 3).value = { formula: `=${tallyQuoted}!${tfPtsColL}${row}` };
    } else {
      ws.getCell(row, 2).value = 0;
      ws.getCell(row, 3).value = 0;
    }
    if (nSa > 0 && saTotalColL) {
      ws.getCell(row, 4).value = {
        formula: `=${saQuoted}!${saTotalColL}${row}`,
      };
    } else {
      ws.getCell(row, 4).value = "";
    }
    // Overall percent: combined points (MC pts in B, TF pts in C, SA total in D — blank SA treated as 0)
    // divided by max combined points (Max MC+TF in $B$2 + Max SA in $D$2). Returns blank when max is zero.
    const pctCell = ws.getCell(row, 5);
    pctCell.value = {
      formula: `=IF($B$2+$D$2=0,"",(B${row}+C${row}+IF(D${row}="",0,D${row}))/($B$2+$D$2))`,
    };
    pctCell.numFmt = "0.00%";
    ws.getCell(row, 6).value = {
      formula: `=IF(E${row}="","",VLOOKUP(E${row},cutoffs,2,TRUE))`,
    };
  }

  ws.getColumn(1).width = 16;
  for (let c = 2; c <= lastCol; c++) {
    ws.getColumn(c).width = 12;
  }
}
