import { describe, expect, it } from "vitest";
import { reconcileCourse, validateBatch } from "../js/batchValidate.js";

describe("reconcileCourse", () => {
  it("adopts first course", () => {
    expect(reconcileCourse("", "AB")).toEqual({ ok: true, course: "AB" });
  });

  it("promotes short to long when dashed extension matches", () => {
    expect(reconcileCourse("CS", "CS-F25")).toEqual({ ok: true, course: "CS-F25" });
  });

  it("rejects incompatible dashed codes", () => {
    const r = reconcileCourse("CS-FALL", "CS-SPRING");
    expect(r.ok).toBe(false);
  });
});

describe("validateBatch", () => {
  it("passes consistent batch", () => {
    const v = validateBatch([
      {
        fileName: "a.pdf",
        course: "X",
        exam_no: "1",
        mc: { 1: "A" },
        essayCount: 2,
      },
      {
        fileName: "b.pdf",
        course: "X",
        exam_no: "2",
        mc: { 1: "B" },
        essayCount: 2,
      },
    ]);
    expect(v.ok).toBe(true);
    expect(v.nSa).toBe(2);
    expect(v.exams.size).toBe(2);
  });

  it("fails on duplicate exam id", () => {
    const v = validateBatch([
      {
        fileName: "a.pdf",
        course: "X",
        exam_no: "1",
        mc: { 1: "A" },
        essayCount: 1,
      },
      {
        fileName: "b.pdf",
        course: "X",
        exam_no: "1",
        mc: { 1: "B" },
        essayCount: 1,
      },
    ]);
    expect(v.ok).toBe(false);
    expect(v.errors.some((e) => e.includes("duplicate"))).toBe(true);
  });
});
