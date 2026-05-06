import { describe, expect, it } from "vitest";
import {
  computeModalEssayNa,
  findEssayCountOutliers,
} from "../js/essayCountModal.js";

describe("computeModalEssayNa", () => {
  it("returns mode", () => {
    expect(computeModalEssayNa([5, 13, 5, 5, null])).toBe(5);
  });

  it("tie-break prefers smallest among tied modes", () => {
    expect(computeModalEssayNa([3, 3, 7, 7])).toBe(3);
  });

  it("returns null when no valid counts", () => {
    expect(computeModalEssayNa([null, null])).toBe(null);
  });
});

describe("findEssayCountOutliers", () => {
  it("lists missing and mismatched", () => {
    const o = findEssayCountOutliers(
      [
        { fileName: "a.pdf", essayCount: 5 },
        { fileName: "b.pdf", essayCount: null },
        { fileName: "c.pdf", essayCount: 9 },
      ],
      5
    );
    expect(o).toHaveLength(2);
  });
});
