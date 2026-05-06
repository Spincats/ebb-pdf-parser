import { computeModalEssayNa, findEssayCountOutliers } from "./essayCountModal.js";

/**
 * Python-compatible course reconciliation (mc_parse.py lines 135-149).
 * @param {string} aggregateCourse
 * @param {string} nextCourse
 * @returns {{ ok: true, course: string } | { ok: false, message: string }}
 */
export function reconcileCourse(aggregateCourse, nextCourse) {
  if (aggregateCourse === "") {
    return { ok: true, course: nextCourse };
  }
  if (aggregateCourse === nextCourse) {
    return { ok: true, course: aggregateCourse };
  }

  const oldFirstDash = aggregateCourse.indexOf("-");
  const newFirstDash = nextCourse.indexOf("-");
  const oldHas = oldFirstDash > -1;
  const newHas = newFirstDash > -1;

  if (oldHas && !newHas) {
    if (aggregateCourse.slice(0, oldFirstDash) !== nextCourse) {
      return {
        ok: false,
        message: `course mismatch: expected family "${aggregateCourse.slice(0, oldFirstDash)}" but got "${nextCourse}"`,
      };
    }
    return { ok: true, course: aggregateCourse };
  }

  if (!oldHas && newHas) {
    if (aggregateCourse !== nextCourse.slice(0, newFirstDash)) {
      return {
        ok: false,
        message: `course mismatch: batch "${aggregateCourse}" vs new "${nextCourse}"`,
      };
    }
    return { ok: true, course: nextCourse };
  }

  return {
    ok: false,
    message: `course mismatch: batch "${aggregateCourse}" vs new "${nextCourse}"`,
  };
}

/**
 * @typedef {object} ParsedExamRow
 * @property {string} fileName
 * @property {string} course
 * @property {string} exam_no
 * @property {Record<string, string>} mc
 * @property {number | null} essayCount
 */

/**
 * @param {ParsedExamRow[]} rows
 * @returns {{ ok: boolean, errors: string[], warnings: string[], reconciledCourse: string, nSa: number | null, exams: Map<string, Record<string, string>> }}
 */
export function validateBatch(rows) {
  const errors = [];
  const warnings = [];

  let reconciledCourse = "";
  for (const row of rows) {
    if (!row.course || !row.course.trim()) {
      errors.push(`${row.fileName}: missing course (no #... line found).`);
    } else {
      const r = reconcileCourse(reconciledCourse, row.course);
      if (!r.ok) {
        errors.push(`${row.fileName}: ${r.message}`);
      } else {
        reconciledCourse = r.course;
      }
    }
  }

  const essayCounts = rows.map((r) => r.essayCount);
  const nSa = computeModalEssayNa(essayCounts);
  if (nSa === null) {
    errors.push(
      "Could not determine Number of Essay Questions from any PDF (line missing or unparsable on every file)."
    );
  } else {
    const outliers = findEssayCountOutliers(
      rows.map((r) => ({ fileName: r.fileName, essayCount: r.essayCount })),
      nSa
    );
    for (const o of outliers) {
      if (o.parsed === null) {
        warnings.push(
          `${o.fileName}: missing or unparsable "Number of Essay Questions" (batch modal ${nSa}).`
        );
      } else {
        warnings.push(
          `${o.fileName}: essay question count ${o.parsed} differs from batch modal ${nSa}.`
        );
      }
    }
  }

  /** @type {Map<string, { fileName: string, mc: Record<string, string> }>} */
  const byExam = new Map();
  for (const row of rows) {
    const id = row.exam_no.trim();
    if (!id) {
      errors.push(`${row.fileName}: missing Anonymous Number (student ID).`);
      continue;
    }
    if (byExam.has(id)) {
      const prev = byExam.get(id);
      errors.push(
        `${row.fileName}: duplicate anonymous ID "${id}" (also in ${prev?.fileName}).`
      );
    } else {
      byExam.set(id, { fileName: row.fileName, mc: row.mc });
    }
    const keys = Object.keys(row.mc);
    if (keys.length === 0) {
      errors.push(`${row.fileName}: no multiple-choice answers found.`);
    }
  }

  const exams = new Map();
  for (const [id, v] of byExam) {
    exams.set(id, v.mc);
  }

  const ok = errors.length === 0 && nSa !== null;
  return { ok, errors, warnings, reconciledCourse, nSa, exams };
}
