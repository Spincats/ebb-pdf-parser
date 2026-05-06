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

  it("accepts text after the Multiple Choice Answers header (e.g. PM suffix)", () => {
    const text = `
# K, z
Anonymous Number: 1
-->   Multiple   Choice   Answers   PM
1.   C
2.   D
--->End   Multiple   Choice
`;
    const r = parseExamText(text);
    expect(r.mc["1"]).toBe("C");
    expect(r.mc["2"]).toBe("D");
  });

  it("tolerates pdf.js-style extra spaces (Exams 2)", () => {
    const text = `
Property,   Class   #   1234_X,   with   Y
Anonymous   Number:   911
-->   Multiple   Choice   Answers
1.   B
9.   C
10.   D
20.   A
--->End   Multiple   Choice
`;
    const r = parseExamText(text);
    expect(r.course).toBe("1234_X");
    expect(r.exam_no).toBe("911");
    expect(r.mc["1"]).toBe("B");
    expect(r.mc["9"]).toBe("C");
    expect(r.mc["10"]).toBe("D");
    expect(r.mc["20"]).toBe("A");
  });

  it("maps T/F with extra spaces after the dot", () => {
    const text = `
# P, q
Anonymous   Number:   1
-->   Multiple   Choice   Answers
1.    t
2.    F
--->End   Multiple   Choice
`;
    const r = parseExamText(text);
    expect(r.mc["1"]).toBe("A");
    expect(r.mc["2"]).toBe("B");
  });

  it("re-parses alternate layout when all first-pass MC answers are identical", () => {
    const text = `
# Z, tail
Anonymous Number: 99
--> Multiple Choice Answers
1. A
2. A
3. A
--->End Multiple Choice
-->Question -1-
Multiple Choice
1. B
2. C
3. D
->End of Question
`;
    const r = parseExamText(text);
    expect(r.exam_no).toBe("99");
    expect(r.mc["1"]).toBe("B");
    expect(r.mc["2"]).toBe("C");
    expect(r.mc["3"]).toBe("D");
  });

  it("alternate layout tolerates spaces in question markers (pdf.js)", () => {
    const text = `
# Z, tail
Anonymous Number: 1
--> Multiple Choice Answers
1. X
2. X
--->End Multiple Choice
-->   Question   -1-
Multiple   Choice
1. P
2. Q
->   End   of   Question
`;
    const r = parseExamText(text);
    expect(r.mc["1"]).toBe("P");
    expect(r.mc["2"]).toBe("Q");
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
