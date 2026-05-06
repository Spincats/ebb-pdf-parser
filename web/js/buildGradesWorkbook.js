import { colIndexToLetters, quoteSheetNameForFormula } from "./colUtils.js";
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
  MC_SUMMARY_BLOCK_GAP_ROWS,
  MC_SUMMARY_GAP_ROWS,
  MC_TALLY_INPUT_ROW,
  MC_TALLY_MC_FIRST_Q_COL,
  MC_TALLY_MC_LAST_Q_COL,
  MC_TALLY_MC_PTS_COL,
  MC_TALLY_TF_FIRST_Q_COL,
  MC_TALLY_TF_LAST_Q_COL,
  MC_TALLY_TF_PTS_COL,
  MC_TRAILER_HEADERS,
  SA_TRAILER_HEADERS,
  SHEET_CUTOFFS,
  SHEET_MC,
  SHEET_MC_TALLY,
  SHEET_SA,
  SHEET_TOTAL,
  TITLE_ROW,
  TOTAL_LETTER_COL,
} from "./gradesSpec.js";

/**
 * @param {string} qid
 * @returns {string}
 */
function mcTableColumnName(qid) {
  return `Q_${String(qid).replace(/[^0-9A-Za-z_]/g, "_")}`;
}

/**
 * @param {number} i 1-based SA index
 */
function saTableColumnName(i) {
  return `SA_${i}`;
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
 * Escapes a literal for use inside an Excel double-quoted string in a formula.
 * @param {string} s
 */
function escapeExcelQuotedLiteral(s) {
  return s.replace(/"/g, '""');
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
    `SUMPRODUCT(--((MID(UPPER(TRIM(${keyRef})),ROW(INDIRECT("1:"&LEN(TRIM(${keyRef})))),1)=` +
    `UPPER(LEFT(TRIM(${ansRef}),1))))>0)`
  );
}

/**
 * @param {typeof import("exceljs").default | import("exceljs").default} ExcelJSMod
 * @param {{ reconciledCourse: string, nSa: number, exams: Map<string, Record<string, string>> }} params
 * @returns {Promise<Blob>}
 */
export async function buildGradesWorkbook(params, ExcelJSMod = null) {
  const ExcelJS = ExcelJSMod ?? globalThis.ExcelJS;
  if (!ExcelJS) {
    throw new Error("ExcelJS is not loaded (expected global ExcelJS from vendor script).");
  }

  const { reconciledCourse, nSa, exams } = params;
  const studentIds = [...exams.keys()].sort((a, b) => {
    const na = parseInt(a, 10);
    const nb = parseInt(b, 10);
    if (Number.isFinite(na) && Number.isFinite(nb)) return na - nb;
    return a.localeCompare(b);
  });
  const numStudents = studentIds.length;

  /** @type {Set<string>} */
  const mcKeySet = new Set();
  for (const id of studentIds) {
    const mc = exams.get(id);
    if (mc) for (const k of Object.keys(mc)) mcKeySet.add(k);
  }
  const mcKeys = [...mcKeySet].sort((a, b) => parseInt(a, 10) - parseInt(b, 10));
  const nMc = mcKeys.length;

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
  });

  const cutWs = wb.addWorksheet(SHEET_CUTOFFS);
  const totalWs = wb.addWorksheet(SHEET_TOTAL);
  buildTotalSheet(totalWs, {
    numStudents,
    nMc,
    lastStudentRow,
    mcQuoted,
    tallyQuoted,
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
  t.alignment = { vertical: "middle", horizontal: "center" };

  ws.getCell(MC_ANSWER_KEY_ROW, 1).value =
    "Optional row (aligns with MC sheet layout). SA scores start in row below headers.";
  ws.getCell(MC_ANSWER_KEY_ROW, 1).font = { italic: true, size: 9 };

  ws.getCell(HEADER_ROW, 1).value = "Anonymous ID";
  ws.getCell(HEADER_ROW, 1).font = { bold: true };

  const saCols = [];
  for (let i = 1; i <= nSa; i++) {
    saCols.push({ name: saTableColumnName(i) });
  }

  const saDataRows = studentIds.map(() => saCols.map(() => null));

  if (nSa > 0) {
    ws.addTable({
      name: "tblSA",
      displayName: "SA_Data",
      ref: `B${HEADER_ROW}`,
      headerRow: true,
      totalsRow: false,
      style: {
        theme: "TableStyleMedium2",
        showFirstColumn: false,
        showLastColumn: false,
        showRowStripes: true,
        showColumnStripes: false,
      },
      columns: saCols,
      rows: saDataRows,
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

  for (let r = 0; r < numStudents; r++) {
    const row = FIRST_STUDENT_ROW + r;
    ws.getCell(row, 1).value = studentIds[r];
    for (let tc = 0; tc < SA_TRAILER_HEADERS.length; tc++) {
      ws.getCell(row, 2 + nSa + tc).value = null;
    }
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

  const mcCols = mcKeys.map((qid) => ({ name: mcTableColumnName(qid) }));
  const mcDataRows = studentIds.map((id) => {
    const mc = exams.get(id) ?? {};
    return mcKeys.map((q) => mc[q] ?? "");
  });

  if (nMc > 0) {
    ws.addTable({
      name: "tblMC",
      displayName: "MC_Data",
      ref: `B${HEADER_ROW}`,
      headerRow: true,
      totalsRow: false,
      style: {
        theme: "TableStyleMedium9",
        showFirstColumn: false,
        showLastColumn: false,
        showRowStripes: true,
        showColumnStripes: false,
      },
      columns: mcCols,
      rows: mcDataRows,
    });
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

  ws.getColumn(1).width = 18;
  for (let i = 2; i <= lastMcCol; i++) {
    ws.getColumn(i).width = i <= 1 + nMc ? 6 : 10;
  }

  if (nMc > 0) {
    writeMcPerQuestionSummaries(ws, {
      mcKeys,
      studentIds,
      exams,
      lastStudentRow,
    });
  }
}

/**
 * One vertical block per MC question: stats only for that column, with COUNTIF rows
 * for each answer value that appears in the batch for that question (so T/F shows A,B only;
 * five-option MC shows A through E when present).
 * @param {import("exceljs").Worksheet} ws
 */
function writeMcPerQuestionSummaries(ws, ctx) {
  const { mcKeys, studentIds, exams, lastStudentRow } = ctx;
  let row = lastStudentRow + MC_SUMMARY_GAP_ROWS;

  for (let qi = 0; qi < mcKeys.length; qi++) {
    const qid = mcKeys[qi];
    const excelCol = 2 + qi;
    const colName = mcTableColumnName(qid);
    const choices = distinctSortedAnswersForQuestion(qid, studentIds, exams);

    ws.getCell(row, 1).value = `Question ${qid} (summary)`;
    ws.getCell(row, 1).font = { bold: true, italic: true };
    ws.getCell(row, excelCol).value = `Q${qid}`;
    ws.getCell(row, excelCol).font = { bold: true };
    row++;

    const rows = [
      {
        label: "Non-blank",
        formula: `COUNTA(tblMC[[#Data],[${colName}]])`,
      },
      {
        label: "Blank",
        formula: `COUNTBLANK(tblMC[[#Data],[${colName}]])`,
      },
      ...choices.map((ans) => ({
        label: `Count "${ans}"`,
        formula: `COUNTIF(tblMC[[#Data],[${colName}]],"${escapeExcelQuotedLiteral(ans)}")`,
      })),
      {
        label: "Students",
        formula: `ROWS(tblMC[[#Data],[${colName}]])`,
      },
    ];

    for (const spec of rows) {
      ws.getCell(row, 1).value = spec.label;
      ws.getCell(row, 1).font = { italic: true };
      ws.getCell(row, excelCol).value = { formula: spec.formula };
      row++;
    }

    row += MC_SUMMARY_BLOCK_GAP_ROWS;
  }
}

/**
 * MCtally: per-question Ok formulas point at {@link SHEET_MC} keys and answers; column A is total correct;
 * columns M–R hold points-per-question and 1-based MC/TF question index ranges (inclusive) into the MC sheet
 * columns B..; MC pts / TF pts columns use SUMPRODUCT over the local Ok cells in those index ranges.
 * @param {import("exceljs").Worksheet} ws
 */
function buildMcTallySheet(ws, p) {
  const { mcKeys, nMc, numStudents, lastStudentRow, mcQuoted } = p;
  const lastDataCol = nMc > 0 ? 4 + nMc : 6;
  ws.mergeCells(TITLE_ROW, 1, TITLE_ROW, Math.max(lastDataCol, MC_TALLY_TF_LAST_Q_COL));
  const title = ws.getCell(TITLE_ROW, 1);
  title.value = "MC tally (Ok vs keys on MC+TF)";
  title.font = { bold: true, size: 14 };
  title.alignment = { vertical: "middle", horizontal: "center" };

  ws.getCell(1, MC_TALLY_MC_PTS_COL).value = "MC pts each";
  ws.getCell(1, MC_TALLY_TF_PTS_COL).value = "TF pts each";
  ws.getCell(1, MC_TALLY_MC_FIRST_Q_COL).value = "MC Q first (1-based)";
  ws.getCell(1, MC_TALLY_MC_LAST_Q_COL).value = "MC Q last (1-based)";
  ws.getCell(1, MC_TALLY_TF_FIRST_Q_COL).value = "TF Q first (1-based)";
  ws.getCell(1, MC_TALLY_TF_LAST_Q_COL).value = "TF Q last (1-based)";

  ws.getCell(MC_TALLY_INPUT_ROW, MC_TALLY_MC_PTS_COL).value = 1;
  ws.getCell(MC_TALLY_INPUT_ROW, MC_TALLY_TF_PTS_COL).value = 1;
  ws.getCell(MC_TALLY_INPUT_ROW, MC_TALLY_MC_FIRST_Q_COL).value = nMc > 0 ? 1 : 0;
  ws.getCell(MC_TALLY_INPUT_ROW, MC_TALLY_MC_LAST_Q_COL).value = nMc;
  ws.getCell(MC_TALLY_INPUT_ROW, MC_TALLY_TF_FIRST_Q_COL).value = 0;
  ws.getCell(MC_TALLY_INPUT_ROW, MC_TALLY_TF_LAST_Q_COL).value = 0;

  const mL = colIndexToLetters(MC_TALLY_MC_PTS_COL);
  const nL = colIndexToLetters(MC_TALLY_TF_PTS_COL);
  const oL = colIndexToLetters(MC_TALLY_MC_FIRST_Q_COL);
  const pL = colIndexToLetters(MC_TALLY_MC_LAST_Q_COL);
  const qL = colIndexToLetters(MC_TALLY_TF_FIRST_Q_COL);
  const rL = colIndexToLetters(MC_TALLY_TF_LAST_Q_COL);
  const ir = MC_TALLY_INPUT_ROW;

  if (nMc > 0) {
    const mcPtsCol = 3 + nMc;
    const tfPtsCol = 4 + nMc;

    const firstOkL = colIndexToLetters(3);
    const lastOkL = colIndexToLetters(2 + nMc);

    const tallyRows = [];
    for (let r = 0; r < numStudents; r++) {
      const row = FIRST_STUDENT_ROW + r;
      const totalF = `=SUMPRODUCT(0+(${firstOkL}${row}:${lastOkL}${row}))`;
      const idF = `=${mcQuoted}!A${row}`;
      const okCells = [];
      for (let qi = 0; qi < nMc; qi++) {
        const mcCol = 2 + qi;
        okCells.push({
          formula: buildMcKeyMatchFormulaForSheet(`${mcQuoted}!`, mcCol, row),
        });
      }
      const mcPtsF =
        `=$${mL}$${ir}*IF(OR($${oL}$${ir}<1,$${pL}$${ir}<$${oL}$${ir},$${oL}$${ir}>${nMc},$${pL}$${ir}>${nMc}),0,` +
        `SUMPRODUCT(0+(OFFSET($A${row},0,(1+$${oL}$${ir}),1,$${pL}$${ir}-$${oL}$${ir}+1))))`;
      const tfPtsF =
        `=$${nL}$${ir}*IF(OR($${qL}$${ir}=0,$${rL}$${ir}<$${qL}$${ir},$${qL}$${ir}>${nMc},$${rL}$${ir}>${nMc}),0,` +
        `SUMPRODUCT(0+(OFFSET($A${row},0,(1+$${qL}$${ir}),1,$${rL}$${ir}-$${qL}$${ir}+1))))`;

      tallyRows.push([
        { formula: totalF },
        { formula: idF },
        ...okCells,
        { formula: mcPtsF },
        { formula: tfPtsF },
      ]);
    }

    const lastColL = colIndexToLetters(tfPtsCol);
    ws.addTable({
      name: "tblTally",
      displayName: "MC_Tally_Data",
      ref: `A${HEADER_ROW}:${lastColL}${lastStudentRow}`,
      headerRow: true,
      totalsRow: false,
      style: {
        theme: "TableStyleMedium9",
        showFirstColumn: false,
        showLastColumn: false,
        showRowStripes: true,
        showColumnStripes: false,
      },
      columns: [
        { name: "Total_correct" },
        { name: "Anonymous_ID" },
        ...mcKeys.map((qid) => ({
          name: `Ok_Q_${String(qid).replace(/[^0-9A-Za-z_]/g, "_")}`,
        })),
        { name: "MC_pts" },
        { name: "TF_pts" },
      ],
      rows: tallyRows,
    });
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
 * Cutoffs: F25-style ladder (Value, Grade, GPA, Count, Pct), Total N, class GPA, and curve guidance (not enforced).
 * Count uses COUNTIF(letters, Grade); Pct uses Count / Total N. Depends on workbook name `letters` (set after Total).
 * @param {import("exceljs").Worksheet} ws
 */
function buildCutoffsSheet(ws) {
  const lastRow = CUTOFFS_DATA_LAST_ROW;
  const totalNRow = CUTOFFS_TOTAL_N_VALUE_ROW;
  const rows = F25_GRADE_CUTOFF_ROWS.map(([minV, letter, gpa], i) => {
    const row = CUTOFFS_DATA_FIRST_ROW + i;
    return [
      minV,
      letter,
      gpa,
      { formula: `=COUNTIF(letters,B${row})` },
      { formula: `=D${row}/$B$${totalNRow}` },
    ];
  });

  ws.addTable({
    name: "tblCutoffs",
    displayName: "Cutoff_Bands",
    ref: `A${CUTOFFS_HEADER_ROW}:E${lastRow}`,
    headerRow: true,
    totalsRow: false,
    style: {
      theme: "TableStyleMedium2",
      showFirstColumn: false,
      showLastColumn: false,
      showRowStripes: true,
      showColumnStripes: false,
    },
    columns: [
      { name: "Value" },
      { name: "Grade" },
      { name: "GPA" },
      { name: "Count" },
      { name: "Pct" },
    ],
    rows,
  });

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
  bMinusHint.value = "Curve target about 5-20% at B- (guidance only).";
  bMinusHint.font = { italic: true, size: 9 };
  bMinusHint.alignment = { wrapText: true, vertical: "top" };

  const aHint = ws.getCell(10, 6);
  aHint.value = "Curve target about 5-30% at A (guidance only).";
  aHint.font = { italic: true, size: 9 };
  aHint.alignment = { wrapText: true, vertical: "top" };

  const gpaBandHint = ws.getCell(6, 8);
  gpaBandHint.value =
    "Typical class GPA band about 3.2-3.4 before participation adj. (guidance only).";
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

  for (let c = 2; c <= 5; c++) {
    ws.getCell(CUTOFFS_PART_POS_VALUE_ROW, c).value = 0;
  }

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
  for (let c = 2; c <= 5; c++) {
    ws.getCell(CUTOFFS_PART_NEG_VALUE_ROW, c).value = 0;
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
 * Total: one row per student with IDs from MC+TF, tallies from MCtally, percent of max points, letter from Cutoffs.
 * @param {import("exceljs").Worksheet} ws
 */
function buildTotalSheet(ws, p) {
  const { numStudents, nMc, lastStudentRow, mcQuoted, tallyQuoted } = p;
  const lastCol = 6;
  ws.mergeCells(TITLE_ROW, 1, TITLE_ROW, lastCol);
  const title = ws.getCell(TITLE_ROW, 1);
  title.value = "Totals (MC/TF points and letter grade)";
  title.font = { bold: true, size: 14 };
  title.alignment = { vertical: "middle", horizontal: "center" };

  ws.getCell(2, 1).value = "Max points (from MCtally inputs)";
  ws.getCell(2, 1).font = { italic: true, size: 10 };
  const tq = tallyQuoted;
  const mLc = colIndexToLetters(MC_TALLY_MC_PTS_COL);
  const nLc = colIndexToLetters(MC_TALLY_TF_PTS_COL);
  const oLc = colIndexToLetters(MC_TALLY_MC_FIRST_Q_COL);
  const pLc = colIndexToLetters(MC_TALLY_MC_LAST_Q_COL);
  const qLc = colIndexToLetters(MC_TALLY_TF_FIRST_Q_COL);
  const rLc = colIndexToLetters(MC_TALLY_TF_LAST_Q_COL);
  const maxPtsF =
    `=${tq}!$${mLc}$${MC_TALLY_INPUT_ROW}*IF(OR(${tq}!$${oLc}$${MC_TALLY_INPUT_ROW}<1,` +
    `${tq}!$${pLc}$${MC_TALLY_INPUT_ROW}<${tq}!$${oLc}$${MC_TALLY_INPUT_ROW},` +
    `${tq}!$${oLc}$${MC_TALLY_INPUT_ROW}>${nMc},${tq}!$${pLc}$${MC_TALLY_INPUT_ROW}>${nMc}),0,` +
    `${tq}!$${pLc}$${MC_TALLY_INPUT_ROW}-${tq}!$${oLc}$${MC_TALLY_INPUT_ROW}+1)+` +
    `${tq}!$${nLc}$${MC_TALLY_INPUT_ROW}*IF(OR(${tq}!$${qLc}$${MC_TALLY_INPUT_ROW}=0,` +
    `${tq}!$${rLc}$${MC_TALLY_INPUT_ROW}<${tq}!$${qLc}$${MC_TALLY_INPUT_ROW},` +
    `${tq}!$${qLc}$${MC_TALLY_INPUT_ROW}>${nMc},${tq}!$${rLc}$${MC_TALLY_INPUT_ROW}>${nMc}),0,` +
    `${tq}!$${rLc}$${MC_TALLY_INPUT_ROW}-${tq}!$${qLc}$${MC_TALLY_INPUT_ROW}+1)`;
  ws.getCell(2, 2).value = { formula: maxPtsF };

  ws.getCell(HEADER_ROW, 1).value = "Anonymous ID";
  ws.getCell(HEADER_ROW, 1).font = { bold: true };
  ws.getCell(HEADER_ROW, 2).value = "Total correct";
  ws.getCell(HEADER_ROW, 2).font = { bold: true };
  ws.getCell(HEADER_ROW, 3).value = "MC pts";
  ws.getCell(HEADER_ROW, 3).font = { bold: true };
  ws.getCell(HEADER_ROW, 4).value = "TF pts";
  ws.getCell(HEADER_ROW, 4).font = { bold: true };
  ws.getCell(HEADER_ROW, 5).value = "Pct (0-1)";
  ws.getCell(HEADER_ROW, 5).font = { bold: true };
  ws.getCell(HEADER_ROW, 6).value = "Letter";
  ws.getCell(HEADER_ROW, 6).font = { bold: true };

  const mcPtsColL = nMc > 0 ? colIndexToLetters(3 + nMc) : "C";
  const tfPtsColL = nMc > 0 ? colIndexToLetters(4 + nMc) : "D";

  for (let r = 0; r < numStudents; r++) {
    const row = FIRST_STUDENT_ROW + r;
    ws.getCell(row, 1).value = { formula: `=${mcQuoted}!A${row}` };
    if (nMc > 0) {
      ws.getCell(row, 2).value = { formula: `=${tallyQuoted}!A${row}` };
      ws.getCell(row, 3).value = { formula: `=${tallyQuoted}!${mcPtsColL}${row}` };
      ws.getCell(row, 4).value = { formula: `=${tallyQuoted}!${tfPtsColL}${row}` };
      ws.getCell(row, 5).value = {
        formula: `=IF($B$2=0,"",(C${row}+D${row})/$B$2)`,
      };
      ws.getCell(row, 6).value = {
        formula: `=IF(E${row}="","",VLOOKUP(E${row},cutoffs,2,TRUE))`,
      };
    } else {
      ws.getCell(row, 2).value = 0;
      ws.getCell(row, 3).value = 0;
      ws.getCell(row, 4).value = 0;
      ws.getCell(row, 5).value = "";
      ws.getCell(row, 6).value = "";
    }
  }

  ws.getColumn(1).width = 16;
  ws.getColumn(2).width = 14;
  for (let c = 3; c <= 6; c++) {
    ws.getColumn(c).width = 12;
  }
}
