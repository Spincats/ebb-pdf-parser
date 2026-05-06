/**
 * Helpers for workbook export: student row counts and MCtally MC/TF question index validation.
 */

/**
 * @param {Iterable<string>} ids
 * @returns {string[]}
 */
export function sortStudentIds(ids) {
  return [...ids].sort((a, b) => {
    const na = parseInt(a, 10);
    const nb = parseInt(b, 10);
    if (Number.isFinite(na) && Number.isFinite(nb)) return na - nb;
    return a.localeCompare(b);
  });
}

/**
 * Sorted distinct MC question ids (string keys) across all exams, same order as MC+TF columns.
 *
 * @param {Map<string, Record<string, string>>} exams
 * @returns {string[]}
 */
export function listMcQuestionIds(exams) {
  const s = new Set();
  for (const mc of exams.values()) {
    for (const k of Object.keys(mc)) s.add(k);
  }
  return [...s].sort((a, b) => parseInt(a, 10) - parseInt(b, 10));
}

/**
 * Number of MC+TF question columns (matches workbook builder).
 *
 * @param {Map<string, Record<string, string>>} exams
 * @returns {number}
 */
export function mcQuestionCountForTally(exams) {
  return listMcQuestionIds(exams).length;
}

/**
 * Builds an exams map with exactly `studentCount` rows: truncates sorted IDs or pads with
 * empty MC rows using synthetic keys `__placeholder_n` (not from PDFs).
 *
 * @param {Map<string, Record<string, string>>} exams
 * @param {number | null | undefined} studentCount positive integer, or nullish to keep `exams`
 * @returns {{ exams: Map<string, Record<string, string>>, studentIds: string[] }}
 */
export function applyStudentCountToExams(exams, studentCount) {
  const sortedIds = sortStudentIds(exams.keys());
  const n = sortedIds.length;
  if (
    studentCount === null ||
    studentCount === undefined ||
    !Number.isFinite(studentCount)
  ) {
    return { exams: new Map(exams), studentIds: sortedIds };
  }
  const target = Math.max(1, Math.min(5000, Math.floor(Number(studentCount))));
  const out = new Map();
  if (target <= n) {
    for (let i = 0; i < target; i++) {
      const id = sortedIds[i];
      out.set(id, { ...(exams.get(id) ?? {}) });
    }
    return { exams: out, studentIds: sortStudentIds(out.keys()) };
  }
  for (const id of sortedIds) {
    out.set(id, { ...(exams.get(id) ?? {}) });
  }
  let pad = 1;
  while (out.size < target) {
    const pid = `__placeholder_${pad}`;
    if (!out.has(pid)) out.set(pid, {});
    pad++;
  }
  return { exams: out, studentIds: sortStudentIds(out.keys()) };
}

/**
 * Validates 1-based inclusive MC and TF question index ranges for MCtally row 2 (columns O-R).
 * TF is disabled only when both first and last are 0. When enabled, ranges must lie in 1..nMc
 * and must not overlap the MC range.
 *
 * @param {number} nMc
 * @param {number} mcFirst
 * @param {number} mcLast
 * @param {number} tfFirst
 * @param {number} tfLast
 * @returns {string | null} error message, or null if valid
 */
export function validateMcTallyQuestionRanges(nMc, mcFirst, mcLast, tfFirst, tfLast) {
  if (nMc <= 0) return null;
  if (!Number.isInteger(mcFirst) || !Number.isInteger(mcLast)) {
    return "MC question range must be whole numbers.";
  }
  if (mcFirst < 1 || mcLast < mcFirst || mcLast > nMc) {
    return `MC question range must satisfy 1 <= first <= last <= ${nMc} (question count on MC+TF).`;
  }

  const tfOff = tfFirst === 0 && tfLast === 0;
  if (tfOff) {
    return null;
  }
  if (!Number.isInteger(tfFirst) || !Number.isInteger(tfLast)) {
    return "TF question range must be whole numbers (or set both to 0 to disable TF).";
  }
  if (tfFirst < 1 || tfLast < tfFirst || tfLast > nMc) {
    return `TF question range must satisfy 1 <= first <= last <= ${nMc}, or use 0 and 0 to disable.`;
  }
  if (!(mcLast < tfFirst || tfLast < mcFirst)) {
    return "MC and TF question ranges must not overlap.";
  }
  return null;
}
