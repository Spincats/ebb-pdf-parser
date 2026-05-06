import { describe, expect, it } from "vitest";
import {
  applyStudentCountToExams,
  mcQuestionCountForTally,
  validateMcTallyQuestionRanges,
} from "../js/workbookExport.js";

describe("applyStudentCountToExams", () => {
  it("returns a copy with the same order when studentCount is omitted", () => {
    const exams = new Map([
      ["anon-fake-010", { 1: "A" }],
      ["anon-fake-002", { 1: "B" }],
    ]);
    const { exams: out, studentIds } = applyStudentCountToExams(exams, null);
    expect(out.size).toBe(2);
    expect(studentIds).toEqual(["anon-fake-002", "anon-fake-010"]);
  });

  it("truncates to the first N sorted IDs", () => {
    const exams = new Map([
      ["anon-fake-003", { 1: "A" }],
      ["anon-fake-001", { 1: "B" }],
      ["anon-fake-002", { 1: "C" }],
    ]);
    const { exams: out, studentIds } = applyStudentCountToExams(exams, 2);
    expect(out.size).toBe(2);
    expect(studentIds).toEqual(["anon-fake-001", "anon-fake-002"]);
  });

  it("pads with synthetic placeholder rows", () => {
    const exams = new Map([["anon-fake-001", { 1: "A" }]]);
    const { exams: out, studentIds } = applyStudentCountToExams(exams, 3);
    expect(out.size).toBe(3);
    expect(studentIds.some((id) => id.startsWith("__placeholder_"))).toBe(true);
  });
});

describe("validateMcTallyQuestionRanges", () => {
  it("accepts default-like MC 1-n and TF off", () => {
    expect(validateMcTallyQuestionRanges(5, 1, 5, 0, 0)).toBeNull();
  });

  it("rejects overlapping MC and TF ranges", () => {
    const msg = validateMcTallyQuestionRanges(10, 1, 5, 3, 7);
    expect(msg).toBeTruthy();
    expect(msg).toContain("overlap");
  });

  it("accepts disjoint MC and TF ranges", () => {
    expect(validateMcTallyQuestionRanges(10, 1, 4, 5, 8)).toBeNull();
  });
});

describe("mcQuestionCountForTally", () => {
  it("counts distinct question keys", () => {
    const exams = new Map([
      ["anon-fake-001", { 1: "A", 2: "B" }],
      ["anon-fake-002", { 2: "C" }],
    ]);
    expect(mcQuestionCountForTally(exams)).toBe(2);
  });
});
