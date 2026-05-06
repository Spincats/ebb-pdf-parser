/**
 * Layout and copy derived from the F25-style gradebook (encoded in source; no template XLSX at runtime).
 * Excel 365+ dynamic arrays / LET are not required; formulas use SUM, COUNT, COUNTIF, VLOOKUP, and plain A1 ranges for Calc/Sheets compatibility.
 */

export const SHEET_SA = "ShortAnswer+Essay";
export const SHEET_MC = "MC+TF";
/** F25-aligned tally sheet (Ok columns and points); encoded from plan when F25.xlsx is absent from repo */
export const SHEET_MC_TALLY = "MCtally";
export const SHEET_CUTOFFS = "Cutoffs";
export const SHEET_TOTAL = "Total";

/**
 * Cutoffs sheet layout matches F25 Grades.xlsx: header on row 1, ladder rows 2-10.
 * Column A = min course fraction (0-1) for VLOOKUP approximate match; B = letter; C = GPA.
 */
export const CUTOFFS_HEADER_ROW = 1;
export const CUTOFFS_DATA_FIRST_ROW = 2;
export const CUTOFFS_DATA_LAST_ROW = 10;

/** Cutoffs: "Total N" label row (1-based), same as F25. */
export const CUTOFFS_TOTAL_N_LABEL_ROW = 14;
/** Cutoffs: cell with =COUNTA(letters); Pct column divides by this row in column B. */
export const CUTOFFS_TOTAL_N_VALUE_ROW = 15;

/** Cutoffs: participation adjustment block (F25 rows 17-21). */
export const CUTOFFS_PART_POS_LABEL_ROW = 17;
export const CUTOFFS_PART_POS_VALUE_ROW = 18;
export const CUTOFFS_PART_NEG_LABEL_ROW = 20;
export const CUTOFFS_PART_NEG_VALUE_ROW = 21;
/** Cutoffs: column for "GPA w/ Part" label and formula (1-based). */
export const CUTOFFS_GPA_WITH_PART_COL = 7;

/**
 * Total sheet column (1-based) where final letter grades are written; workbook name `letters` uses this column.
 * Columns: A ID, B MC pts, C TF pts, D SA total, E overall pct, F letter.
 */
export const TOTAL_LETTER_COL = 6;

/**
 * [minFraction, letterGrade, gpa] in ascending min for VLOOKUP approximate match.
 * Ladder: A starts at 90% and each grade level steps down 5%; F is the catch-all at 0%.
 * Grades: F, D, C, C+, B-, B, B+, A-, A.
 */
export const F25_GRADE_CUTOFF_ROWS = [
  [0, "F", 0],
  [0.55, "D", 1],
  [0.6, "C", 2],
  [0.65, "C+", 2.33],
  [0.7, "B-", 2.67],
  [0.75, "B", 3],
  [0.8, "B+", 3.33],
  [0.85, "A-", 3.67],
  [0.9, "A", 4],
];

/** MCtally: row for points and range inputs (1-based) */
export const MC_TALLY_INPUT_ROW = 2;
/** MCtally: column letters for inputs (1-based indices) — M=13 through R=18 */
export const MC_TALLY_MC_PTS_COL = 13;
export const MC_TALLY_TF_PTS_COL = 14;
export const MC_TALLY_MC_FIRST_Q_COL = 15;
export const MC_TALLY_MC_LAST_Q_COL = 16;
export const MC_TALLY_TF_FIRST_Q_COL = 17;
export const MC_TALLY_TF_LAST_Q_COL = 18;

/** Rows between last student row and the MC summary grid (F25 uses a gap of 3: data ends row 37, summary row 41). */
export const MC_SUMMARY_GAP_ROWS = 3;

/**
 * Row where instructors enter the accepted MC/TF answer key per question column
 * (same column as student answers). Multiple letters mean any one is accepted, e.g. AB = A or B.
 */
export const MC_ANSWER_KEY_ROW = 2;

/** Column header row for student data tables (1-based) */
export const HEADER_ROW = 3;

/** First data row for student rows (1-based); must match on SA and MC sheets */
export const FIRST_STUDENT_ROW = 4;

/** Title row (1-based) */
export const TITLE_ROW = 1;

/** Short-answer trailer column headers (last three columns of SA sheet; matches F25 XC / TOTAL / Notes). */
export const SA_TRAILER_HEADERS = ["XC", "TOTAL", "Notes"];

/** MC+TF trailer column headers (last two columns): adjustments first, then notes. */
export const MC_TRAILER_HEADERS = ["Adjustments", "MC notes"];
