import ExcelJS from "exceljs";
import { describe, expect, it } from "vitest";
import {
  buildGradesWorkbook,
  buildMcKeyMatchFormula,
  buildMcKeyMatchFormulaForSheet,
  distinctSortedAnswersForQuestion,
  sortMcAnswerChoices,
} from "../js/buildGradesWorkbook.js";
import { quoteSheetNameForFormula } from "../js/colUtils.js";

describe("sortMcAnswerChoices", () => {
  it("orders single letters A-Z before other tokens", () => {
    expect(sortMcAnswerChoices(["B", "a", "AA", "C"])).toEqual(["a", "B", "C", "AA"]);
  });
});

describe("distinctSortedAnswersForQuestion", () => {
  it("returns only values present for that question", () => {
    const exams = new Map([
      ["1", { 1: "A", 2: "E" }],
      ["2", { 1: "B", 2: "D" }],
    ]);
    expect(distinctSortedAnswersForQuestion("1", ["1", "2"], exams)).toEqual(["A", "B"]);
    expect(distinctSortedAnswersForQuestion("2", ["1", "2"], exams)).toEqual(["D", "E"]);
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
  it("writes MC table and per-question summary COUNTIF rows for observed answers only", async () => {
    const exams = new Map([
      ["10", { 1: "A", 2: "B" }],
      ["20", { 1: "A", 2: "A" }],
    ]);
    const blob = await buildGradesWorkbook(
      {
        reconciledCourse: "TEST-101",
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

    // lastStudentRow=5, summary starts 14; Q1 Count "A" at row 17 col B (see writeMcPerQuestionSummaries)
    const countAq1 = mc.getCell(17, 2);
    expect(
      countAq1.value &&
        typeof countAq1.value === "object" &&
        "formula" in countAq1.value
    ).toBe(true);
    expect(String(countAq1.value.formula)).toContain("COUNTIF");
    expect(String(countAq1.value.formula)).toContain("tblMC");
    expect(String(countAq1.value.formula)).toContain('"A"');

    // Q2 block: Count "B" row 24 col C
    const countBq2 = mc.getCell(24, 3);
    expect(String(countBq2.value.formula)).toContain("COUNTIF");
    expect(String(countBq2.value.formula)).toContain('"B"');
  });

  it("includes extra answer letters when present (e.g. five options)", async () => {
    const exams = new Map([
      ["1", { 1: "A", 2: "E" }],
      ["2", { 1: "C", 2: "E" }],
    ]);
    const blob = await buildGradesWorkbook(
      {
        reconciledCourse: "X",
        nSa: 1,
        exams,
      },
      ExcelJS
    );
    const buf = await blob.arrayBuffer();
    const wb = new ExcelJS.Workbook();
    await wb.xlsx.load(buf);
    const mc = wb.getWorksheet("MC+TF");

    const formulas = [];
    for (let r = 1; r <= 80; r++) {
      const v = mc.getCell(r, 3).value;
      if (v && typeof v === "object" && "formula" in v) {
        formulas.push(String(v.formula));
      }
    }
    const hasE = formulas.some((f) => f.includes('"E"') && f.includes("COUNTIF"));
    expect(hasE).toBe(true);
  });

  it("puts per-question Ok on MCtally with key-match formulas referencing MC+TF row 2", async () => {
    const exams = new Map([["1", { 1: "A" }]]);
    const blob = await buildGradesWorkbook(
      { reconciledCourse: "K", nSa: 0, exams },
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
    expect(cut.getCell(2, 1).value).toBe(0);
    expect(cut.getCell(2, 2).value).toBe("F");
    expect(cut.getCell(2, 3).value).toBe(0);
    expect(cut.getCell(5, 2).value).toBe("C+");
    expect(cut.getCell(5, 3).value).toBe(2.33);
    expect(cut.getCell(6, 2).value).toBe("B-");
    expect(cut.getCell(6, 3).value).toBe(2.67);
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
    expect(lettersName.ranges[0]).toMatch(/Total!\$/);

    const cutoffsName = names.find((d) => d.name === "cutoffs");
    expect(cutoffsName).toBeTruthy();
    expect(cutoffsName.ranges[0]).toMatch(/Cutoffs!\$/);

    const g18 = cut.getCell(18, 7).value;
    expect(g18 && typeof g18 === "object" && "formula" in g18).toBe(true);
    expect(String(g18.formula)).toContain("B18/3");
    expect(String(g18.formula)).toContain("B21");
    expect(cut.getCell(17, 2).value).toBe("Participation Adj. Positive");
    expect(cut.getCell(20, 2).value).toBe("Participation Adj. Negative");

    const tot = wb.getWorksheet("Total");
    expect(tot).toBeTruthy();
    const maxPts = tot.getCell(2, 2).value;
    expect(maxPts && typeof maxPts === "object" && "formula" in maxPts).toBe(true);
    expect(String(maxPts.formula)).toContain("'MCtally'");

    const letterCell = tot.getCell(4, 6).value;
    expect(letterCell && typeof letterCell === "object" && "formula" in letterCell).toBe(true);
    expect(String(letterCell.formula)).toContain("VLOOKUP");
    expect(String(letterCell.formula)).toContain("cutoffs");
  });

  it("sizes tblTally and tblCutoffs to student and band counts", async () => {
    const exams = new Map([
      ["1", { 1: "A", 2: "B" }],
      ["2", { 1: "B", 2: "A" }],
      ["3", { 1: "A", 2: "A" }],
    ]);
    const blob = await buildGradesWorkbook(
      { reconciledCourse: "Z", nSa: 0, exams },
      ExcelJS
    );
    const buf = await blob.arrayBuffer();
    const wb = new ExcelJS.Workbook();
    await wb.xlsx.load(buf);
    const tally = wb.getWorksheet("MCtally");
    const tblTally = tally.getTables().find((t) => t.name === "tblTally");
    expect(tblTally).toBeTruthy();
    const tallyRef = tblTally.ref ?? tblTally.model?.tableRef;
    expect(tallyRef).toBe("A3:F6");

    const cut = wb.getWorksheet("Cutoffs");
    const tblCut = cut.getTables().find((t) => t.name === "tblCutoffs");
    expect(tblCut).toBeTruthy();
    const cutRef = tblCut.ref ?? tblCut.model?.tableRef;
    expect(cutRef).toBe("A1:E10");
  });
});
