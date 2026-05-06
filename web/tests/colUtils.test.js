import { describe, expect, it } from "vitest";
import { a1Range, colIndexToLetters } from "../js/colUtils.js";

describe("colIndexToLetters", () => {
  it("maps A and Z", () => {
    expect(colIndexToLetters(1)).toBe("A");
    expect(colIndexToLetters(26)).toBe("Z");
  });

  it("maps AA", () => {
    expect(colIndexToLetters(27)).toBe("AA");
  });
});

describe("a1Range", () => {
  it("builds range", () => {
    expect(a1Range(3, 2, 10, 5)).toBe("B3:E10");
  });
});
