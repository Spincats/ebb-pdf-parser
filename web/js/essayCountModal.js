/**
 * Modal (most frequent) essay question count across a batch.
 * Tie-break: among values tied for highest frequency, choose the smallest count
 * (fewer columns; conservative).
 * @param {(number | null)[]} counts
 * @returns {number | null}
 */
export function computeModalEssayNa(counts) {
  const valid = counts.filter(
    (c) => c !== null && c !== undefined && Number.isFinite(c) && c >= 0
  );
  if (valid.length === 0) return null;

  /** @type {Map<number, number>} */
  const freq = new Map();
  for (const c of valid) {
    freq.set(c, (freq.get(c) || 0) + 1);
  }

  let maxFreq = 0;
  for (const f of freq.values()) {
    if (f > maxFreq) maxFreq = f;
  }

  const modes = [...freq.keys()].filter((k) => freq.get(k) === maxFreq);
  return Math.min(...modes);
}

/**
 * @param {{ fileName: string, essayCount: number | null }[]} perFile
 * @param {number} nSa
 * @returns {{ fileName: string, parsed: number | null }[]}
 */
export function findEssayCountOutliers(perFile, nSa) {
  /** @type {{ fileName: string, parsed: number | null }[]} */
  const out = [];
  for (const row of perFile) {
    if (row.essayCount === null) {
      out.push({ fileName: row.fileName, parsed: null });
      continue;
    }
    if (row.essayCount !== nSa) {
      out.push({ fileName: row.fileName, parsed: row.essayCount });
    }
  }
  return out;
}
