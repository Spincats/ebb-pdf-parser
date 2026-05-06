import ExcelJS from "exceljs";
import { describe, expect, it } from "vitest";
import {
  buildGradesWorkbook,
  buildMcKeyMatchFormula,
  buildMcKeyMatchFormulaForSheet,
  distinctSortedAnswersForQuestion,
  isTfStyleMcColumn,
  sortMcAnswerChoices,
} from "../js/buildGradesWorkbook.js";
import { quoteSheetNameForFormula } from "../js/colUtils.js";
import {
  FIRST_STUDENT_ROW,
  MC_SUMMARY_GAP_ROWS,
  MC_TALLY_INPUT_ROW,
  MC_TALLY_MC_FIRST_Q_COL,
  MC_TALLY_MC_LAST_Q_COL,
  MC_TALLY_TF_FIRST_Q_COL,
  MC_TALLY_TF_LAST_Q_COL,
  SA_TRAILER_HEADERS,
} from "../js/gradesSpec.js";

describe("sortMcAnswerChoices", () => {
  it("orders single letters A-Z before other tokens", () => {
    expect(sortMcAnswerChoices(["B", "a", "AA", "C"])).toEqual(["a", "B", "C", "AA"]);
  });
});

describe("distinctSortedAnswersForQuestion", () => {
  it("returns only values present for that question", () => {
    const exams = new Map([
      ["anon-fake-001", { 1: "A", 2: "E" }],
      ["anon-fake-002", { 1: "B", 2: "D" }],
    ]);
    expect(
      distinctSortedAnswersForQuestion("1", ["anon-fake-001", "anon-fake-002"], exams)
    ).toEqual(["A", "B"]);
    expect(
      distinctSortedAnswersForQuestion("2", ["anon-fake-001", "anon-fake-002"], exams)
    ).toEqual(["D", "E"]);
  });
});

describe("buildMcKeyMatchFormula", () => {
  it("matches any letter in the row-2 key against the student first letter (AB accepts A or B)", () => {
    const f = buildMcKeyMatchFormula(2, 4);
    expect(f).toContain("$2");
    expect(f).toContain("MID");
    expect(f).toContain("SUMPRODUCT");
    expect(f).toContain("LEFT(TRIM(B4),1)");
  });

  it("supports a quoted sheet prefix for cross-sheet refs", () => {
    const q = quoteSheetNameForFormula("MC+TF");
    const f = buildMcKeyMatchFormulaForSheet(`${q}!`, 2, 4);
    expect(f).toContain("'MC+TF'!");
    expect(f).toContain("'MC+TF'!B$2");
    expect(f).toContain("'MC+TF'!B4");
  });
});

describe("buildGradesWorkbook", () => {
  it("writes MC grid and F25-style horizontal summary (COUNTIF/COUNTA over plain ranges)", async () => {
    const exams = new Map([
      ["anon-fake-010", { 1: "A", 2: "B" }],
      ["anon-fake-020", { 1: "A", 2: "A" }],
    ]);
    const blob = await buildGradesWorkbook(
      {
        reconciledCourse: "ZZ-UNIT-9101",
        nSa: 2,
        exams,
      },
      ExcelJS
    );
    expect(blob.type).toContain("spreadsheet");
    const buf = await blob.arrayBuffer();
    const wb = new ExcelJS.Workbook();
    await wb.xlsx.load(buf);
    const mc = wb.getWorksheet("MC+TF");
    expect(mc).toBeTruthy();

    const numStudents = 2;
    const lastStudentRow = FIRST_STUDENT_ROW + numStudents - 1;
    const baseRow = lastStudentRow + MC_SUMMARY_GAP_ROWS;
    const rowA = baseRow + 2;
    const rowB = baseRow + 3;

    const pctAq1 = mc.getCell(rowA, 2);
    expect(
      pctAq1.value &&
        typeof pctAq1.value === "object" &&
        "formula" in pctAq1.value
    ).toBe(true);
    expect(String(pctAq1.value.formula)).toContain("COUNTIF");
    expect(String(pctAq1.value.formula)).toContain("COUNTA");
    expect(String(pctAq1.value.formula)).toContain("$B$" + FIRST_STUDENT_ROW);
    expect(String(pctAq1.value.formula)).toContain("$A$" + rowA);

    const pctBq2 = mc.getCell(rowB, 3);
    expect(String(pctBq2.value.formula)).toContain("COUNTIF");
    expect(String(pctBq2.value.formula)).toContain("$A$" + rowB);

    const tot = wb.getWorksheet("Total");
    expect(tot).toBeTruthy();
    const saRef = tot.getCell(FIRST_STUDENT_ROW, 4).value;
    expect(saRef && typeof saRef === "object" && "formula" in saRef).toBe(true);
    expect(String(saRef.formula)).toContain("'ShortAnswer+Essay'!");
  });

  it("includes E row formulas when five-option answers appear", async () => {
    const exams = new Map([
      ["anon-fake-001", { 1: "A", 2: "E" }],
      ["anon-fake-002", { 1: "C", 2: "E" }],
    ]);
    const blob = await buildGradesWorkbook(
      {
        reconciledCourse: "ZZ-UNIT-9102",
        nSa: 1,
        exams,
      },
      ExcelJS
    );
    const buf = await blob.arrayBuffer();
    const wb = new ExcelJS.Workbook();
    await wb.xlsx.load(buf);
    const mc = wb.getWorksheet("MC+TF");

    const numStudents = 2;
    const lastStudentRow = FIRST_STUDENT_ROW + numStudents - 1;
    const baseRow = lastStudentRow + MC_SUMMARY_GAP_ROWS;
    const rowE = baseRow + 6;
    const f = String(mc.getCell(rowE, 3).value.formula);
    expect(f).toContain("COUNTIF");
    expect(f).toContain("$A$" + rowE);
  });

  it("omits C-E summary formulas for TF-style columns (only A/B observed)", async () => {
    const exams = new Map([
      ["anon-fake-001", { 1: "A", 2: "B" }],
      ["anon-fake-002", { 1: "B", 2: "A" }],
    ]);
    expect(
      isTfStyleMcColumn("1", ["anon-fake-001", "anon-fake-002"], exams)
    ).toBe(true);
    const blob = await buildGradesWorkbook(
      { reconciledCourse: "ZZ-UNIT-9103", nSa: 0, exams },
      ExcelJS
    );
    const buf = await blob.arrayBuffer();
    const wb = new ExcelJS.Workbook();
    await wb.xlsx.load(buf);
    const mc = wb.getWorksheet("MC+TF");
    const numStudents = 2;
    const lastStudentRow = FIRST_STUDENT_ROW + numStudents - 1;
    const baseRow = lastStudentRow + MC_SUMMARY_GAP_ROWS;
    const rowC = baseRow + 4;
    expect(mc.getCell(rowC, 2).value).toBeNull();
  });

  it("does not put instructional text on SA row 2 and writes TOTAL as SUM over SA score columns", async () => {
    const exams = new Map([["anon-fake-001", { 1: "A" }]]);
    const blob = await buildGradesWorkbook(
      { reconciledCourse: "ZZ-UNIT-9104", nSa: 3, exams },
      ExcelJS
    );
    const buf = await blob.arrayBuffer();
    const wb = new ExcelJS.Workbook();
    await wb.xlsx.load(buf);
    const sa = wb.getWorksheet("ShortAnswer+Essay");
    expect(sa).toBeTruthy();
    const r2a = sa.getCell(2, 1).value;
    expect(r2a === null || r2a === "").toBe(true);

    const totalCol = 2 + 3 + SA_TRAILER_HEADERS.indexOf("TOTAL");
    const totalF = sa.getCell(FIRST_STUDENT_ROW, totalCol).value;
    expect(totalF && typeof totalF === "object" && "formula" in totalF).toBe(true);
    expect(String(totalF.formula)).toMatch(/=SUM\(B4:/);
    expect(String(totalF.formula)).not.toContain("tblSA");
  });

  it("adds MC conditional formatting including colorScale on summary proportions", async () => {
    const exams = new Map([
      ["anon-fake-001", { 1: "A", 2: "C" }],
      ["anon-fake-002", { 1: "B", 2: "D" }],
    ]);
    const blob = await buildGradesWorkbook(
      { reconciledCourse: "ZZ-UNIT-9105", nSa: 0, exams },
      ExcelJS
    );
    const buf = await blob.arrayBuffer();
    const wb = new ExcelJS.Workbook();
    await wb.xlsx.load(buf);
    const mc = wb.getWorksheet("MC+TF");
    const cfs = mc.conditionalFormattings ?? [];
    const hasColorScale = cfs.some((c) =>
      (c.rules ?? []).some((r) => r.type === "colorScale")
    );
    expect(hasColorScale).toBe(true);
    expect(cfs.length).toBeGreaterThanOrEqual(2);
  });

  it("puts per-question Ok on MCtally with key-match formulas referencing MC+TF row 2", async () => {
    const exams = new Map([["anon-fake-001", { 1: "A" }]]);
    const blob = await buildGradesWorkbook(
      { reconciledCourse: "ZZ-UNIT-9106", nSa: 0, exams },
      ExcelJS
    );
    const buf = await blob.arrayBuffer();
    const wb = new ExcelJS.Workbook();
    await wb.xlsx.load(buf);
    const tally = wb.getWorksheet("MCtally");
    expect(tally).toBeTruthy();
    const ok = tally.getCell(4, 3);
    expect(ok.value && typeof ok.value === "object" && "formula" in ok.value).toBe(true);
    expect(String(ok.value.formula)).toContain("'MC+TF'!");
    expect(String(ok.value.formula)).toContain("$2");
    expect(String(ok.value.formula)).toContain("SUMPRODUCT");

    const total = tally.getCell(4, 1).value;
    expect(total && typeof total === "object" && "formula" in total).toBe(true);
    expect(String(total.formula)).toMatch(/SUMPRODUCT/);

    const cut = wb.getWorksheet("Cutoffs");
    expect(cut).toBeTruthy();
    expect(cut.getCell(1, 1).value).toBe("Cutoff");
    expect(cut.getCell(2, 1).value).toBe(0);
    expect(cut.getCell(2, 1).numFmt).toBe("0.00%");
    expect(cut.getCell(2, 2).value).toBe("F");
    expect(cut.getCell(2, 3).value).toBe(0);
    expect(cut.getCell(5, 1).value).toBe(0.65);
    expect(cut.getCell(5, 2).value).toBe("C+");
    expect(cut.getCell(5, 3).value).toBe(2.33);
    expect(cut.getCell(6, 1).value).toBe(0.7);
    expect(cut.getCell(6, 2).value).toBe("B-");
    expect(cut.getCell(6, 3).value).toBe(2.67);
    expect(cut.getCell(10, 1).value).toBe(0.9);
    expect(cut.getCell(10, 2).value).toBe("A");
    expect(cut.getCell(10, 3).value).toBe(4);

    const d2 = cut.getCell(2, 4).value;
    expect(d2 && typeof d2 === "object" && "formula" in d2).toBe(true);
    expect(String(d2.formula)).toContain("COUNTIF(letters");
    const e2 = cut.getCell(2, 5).value;
    expect(e2 && typeof e2 === "object" && "formula" in e2).toBe(true);
    expect(String(e2.formula)).toContain("$B$15");

    const b15 = cut.getCell(15, 2).value;
    expect(b15 && typeof b15 === "object" && "formula" in b15).toBe(true);
    expect(String(b15.formula)).toContain("COUNTA(letters");

    const names = wb.definedNames.model ?? [];
    const lettersName = names.find((d) => d.name === "letters");
    expect(lettersName).toBeTruthy();
    expect(lettersName.ranges[0]).toMatch(/Total!\$F\$/);

    const cutoffsName = names.find((d) => d.name === "cutoffs");
    expect(cutoffsName).toBeTruthy();
    expect(cutoffsName.ranges[0]).toMatch(/Cutoffs!\$/);

    const g18 = cut.getCell(18, 7).value;
    expect(g18 && typeof g18 === "object" && "formula" in g18).toBe(true);
    expect(String(g18.formula)).toContain("B18/3");
    expect(String(g18.formula)).toContain("B21");
    expect(cut.getCell(17, 2).value).toBe("Participation Adj. Positive");
    expect(cut.getCell(20, 2).value).toBe("Participation Adj. Negative");
    expect(cut.getCell(18, 2).value).toBe(0);
    expect(cut.getCell(21, 2).value).toBe(0);
    expect(cut.getCell(18, 3).value == null || cut.getCell(18, 3).value === "").toBe(true);

    const tot = wb.getWorksheet("Total");
    expect(tot).toBeTruthy();
    const maxPts = tot.getCell(2, 2).value;
    expect(maxPts && typeof maxPts === "object" && "formula" in maxPts).toBe(true);
    expect(String(maxPts.formula)).toContain("'MCtally'");

    expect(tot.getCell(4, 4).value === "" || tot.getCell(4, 4).value == null).toBe(true);

    const pctCell = tot.getCell(4, 5).value;
    expect(pctCell && typeof pctCell === "object" && "formula" in pctCell).toBe(true);
    expect(String(pctCell.formula)).toContain("$B$2+$D$2");

    const letterCell = tot.getCell(4, 6).value;
    expect(letterCell && typeof letterCell === "object" && "formula" in letterCell).toBe(true);
    expect(String(letterCell.formula)).toContain("VLOOKUP");
    expect(String(letterCell.formula)).toContain("cutoffs");

    // The legacy SPARKLINE-based graph in column H rendered as #NAME? in Excel and was removed;
    // there should be no merged graph cell and no SPARKLINE formula anywhere on the sheet.
    const sparkCellValue = tot.getCell(2, 8).value;
    expect(sparkCellValue == null || sparkCellValue === "").toBe(true);
    tot.eachRow({ includeEmpty: false }, (row) => {
      row.eachCell({ includeEmpty: false }, (cell) => {
        if (cell.value && typeof cell.value === "object" && "formula" in cell.value) {
          expect(String(cell.value.formula)).not.toContain("SPARKLINE");
        }
      });
    });

    // Total correct column is now sourced from MCtally only; the Total sheet header row has no
    // "Total correct" cell, and column B holds the MC points header instead.
    expect(tot.getCell(3, 2).value).toBe("MC pts");
    tot.getRow(3).eachCell({ includeEmpty: false }, (cell) => {
      expect(cell.value).not.toBe("Total correct");
    });
  });

  it("uses plain grids for MCtally and Cutoffs (no ListObject tables)", async () => {
    const exams = new Map([
      ["anon-fake-001", { 1: "A", 2: "B" }],
      ["anon-fake-002", { 1: "B", 2: "A" }],
      ["anon-fake-003", { 1: "A", 2: "A" }],
    ]);
    const blob = await buildGradesWorkbook(
      { reconciledCourse: "ZZ-UNIT-9107", nSa: 0, exams },
      ExcelJS
    );
    const buf = await blob.arrayBuffer();
    const wb = new ExcelJS.Workbook();
    await wb.xlsx.load(buf);
    const tally = wb.getWorksheet("MCtally");
    expect(tally.getTables().length).toBe(0);
    expect(tally.getCell(3, 1).value).toBe("Total correct");
    expect(tally.getCell(3, 3).value).toBe("1");

    const cut = wb.getWorksheet("Cutoffs");
    expect(cut.getTables().length).toBe(0);
    expect(cut.getCell(1, 1).value).toBe("Cutoff");
  });

  it("stores MC/TF question index overrides on MCtally input row (columns O-R)", async () => {
    const exams = new Map([
      ["anon-fake-001", { 1: "A", 2: "B", 3: "C" }],
      ["anon-fake-002", { 1: "B", 2: "A", 3: "C" }],
    ]);
    const ir = MC_TALLY_INPUT_ROW;
    const blob = await buildGradesWorkbook(
      {
        reconciledCourse: "ZZ-UNIT-9200",
        nSa: 0,
        exams,
        tallyMcFirst: 1,
        tallyMcLast: 2,
        tallyTfFirst: 3,
        tallyTfLast: 3,
      },
      ExcelJS
    );
    const buf = await blob.arrayBuffer();
    const wb = new ExcelJS.Workbook();
    await wb.xlsx.load(buf);
    const tally = wb.getWorksheet("MCtally");
    expect(tally.getCell(ir, MC_TALLY_MC_FIRST_Q_COL).value).toBe(1);
    expect(tally.getCell(ir, MC_TALLY_MC_LAST_Q_COL).value).toBe(2);
    expect(tally.getCell(ir, MC_TALLY_TF_FIRST_Q_COL).value).toBe(3);
    expect(tally.getCell(ir, MC_TALLY_TF_LAST_Q_COL).value).toBe(3);
  });

  it("throws when MC and TF tally ranges overlap", async () => {
    const exams = new Map([["anon-fake-001", { 1: "A", 2: "B" }]]);
    await expect(
      buildGradesWorkbook(
        {
          reconciledCourse: "ZZ-UNIT-9201",
          nSa: 0,
          exams,
          tallyMcFirst: 1,
          tallyMcLast: 2,
          tallyTfFirst: 2,
          tallyTfLast: 2,
        },
        ExcelJS
      )
    ).rejects.toThrow(/overlap/);
  });
});
