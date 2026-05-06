/**
 * Converts 1-based column index to Excel column letters (A, B, ..., Z, AA, ...).
 * @param {number} colIndex1
 * @returns {string}
 */
export function colIndexToLetters(colIndex1) {
  if (colIndex1 < 1 || !Number.isInteger(colIndex1)) {
    throw new Error(`Invalid column index: ${colIndex1}`);
  }
  let n = colIndex1;
  let s = "";
  while (n > 0) {
    const r = ((n - 1) % 26) + 1;
    s = String.fromCharCode(64 + r) + s;
    n = Math.floor((n - 1) / 26);
  }
  return s;
}

/**
 * @param {number} row1
 * @param {number} col1
 * @param {number} row2
 * @param {number} col2
 * @returns {string} A1-style range
 */
export function a1Range(row1, col1, row2, col2) {
  const c1 = colIndexToLetters(col1);
  const c2 = colIndexToLetters(col2);
  return `${c1}${row1}:${c2}${row2}`;
}

/**
 * Wraps a worksheet name for use in Excel formulas (handles spaces and apostrophes).
 * @param {string} name
 * @returns {string} e.g. 'MC+TF'
 */
export function quoteSheetNameForFormula(name) {
  const escaped = String(name).replace(/'/g, "''");
  return `'${escaped}'`;
}
