import { describe, expect, it } from "vitest";
import { reconcileCourse, validateBatch } from "../js/batchValidate.js";

describe("reconcileCourse", () => {
  it("adopts first course", () => {
    expect(reconcileCourse("", "ZZ-SEED-100")).toEqual({ ok: true, course: "ZZ-SEED-100" });
  });

  it("promotes short to long when dashed extension matches", () => {
    expect(reconcileCourse("ZZ", "ZZ-TERM-8801")).toEqual({ ok: true, course: "ZZ-TERM-8801" });
  });

  it("rejects incompatible dashed codes", () => {
    const r = reconcileCourse("ZZ-FALL", "ZZ-SPRING");
    expect(r.ok).toBe(false);
  });
});

describe("validateBatch", () => {
  it("passes consistent batch", () => {
    const v = validateBatch([
      {
        fileName: "a.pdf",
        course: "ZZ-UNIT-9000",
        exam_no: "anon-fake-001",
        mc: { 1: "A" },
        essayCount: 2,
      },
      {
        fileName: "b.pdf",
        course: "ZZ-UNIT-9000",
        exam_no: "anon-fake-002",
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
        course: "ZZ-UNIT-9000",
        exam_no: "anon-fake-duplicate",
        mc: { 1: "A" },
        essayCount: 1,
      },
      {
        fileName: "b.pdf",
        course: "ZZ-UNIT-9000",
        exam_no: "anon-fake-duplicate",
        mc: { 1: "B" },
        essayCount: 1,
      },
    ]);
    expect(v.ok).toBe(false);
    expect(v.errors.some((e) => e.includes("duplicate"))).toBe(true);
  });
});
