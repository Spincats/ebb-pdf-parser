import { describe, expect, it } from "vitest";
import {
  extractEssayQuestionCount,
  parseExamText,
} from "../js/parseExam.js";

describe("parseExamText", () => {
  it("parses course, exam number, and MC block", () => {
    const text = `
# CS101-01, section
Anonymous Number:  42
--> Multiple Choice Answers
1. A
2. B
--->End Multiple Choice
`;
    const r = parseExamText(text);
    expect(r.course).toBe("CS101-01");
    expect(r.exam_no).toBe("42");
    expect(r.mc["1"]).toBe("A");
    expect(r.mc["2"]).toBe("B");
  });

  it("maps T/F to A/B", () => {
    const text = `
# X, y
Anonymous Number: 1
--> Multiple Choice Answers
1. t
2. F
--->End Multiple Choice
`;
    const r = parseExamText(text);
    expect(r.mc["1"]).toBe("A");
    expect(r.mc["2"]).toBe("B");
  });

  it("strips BLOCKED suffix", () => {
    const text = `
# C, d
Anonymous Number: 9
--> Multiple Choice Answers
1. C BLOCKED extra
--->End Multiple Choice
`;
    const r = parseExamText(text);
    expect(r.mc["1"]).toBe("C");
  });
});

describe("extractEssayQuestionCount", () => {
  it("reads count after label with spaces", () => {
    expect(
      extractEssayQuestionCount("Number of Essay Questions:   13\nother")
    ).toBe(13);
  });

  it("is case-insensitive on label", () => {
    expect(extractEssayQuestionCount("number of essay questions: 5")).toBe(5);
  });

  it("returns null when missing", () => {
    expect(extractEssayQuestionCount("no essay line")).toBe(null);
  });
});
