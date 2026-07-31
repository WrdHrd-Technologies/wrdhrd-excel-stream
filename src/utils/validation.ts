import { getColIndex } from "./columnName";

/**
 * Validates standard worksheet naming constraints enforced by the Microsoft Excel / OpenXML document specifications.
 * Sheets will reject formatting constraints matching punctuation limits, length constraints, or character signatures.
 */
export function validateWorksheetName(name: string): void {
  if (!name || name.trim().length === 0) {
    throw new Error("Worksheet name context cannot be completely empty string or blank.");
  }

  if (name.length > 31) {
    throw new Error(
      `Worksheet name "${name}" breaks safety boundaries: length cannot exceed 31 characters maximum.`
    );
  }

  // Excel prohibited punctuation tracking check array: [ ] \ * ? / :
  const forbiddenCharsRegex = /[\\\/:\?\*\[\]]/g;
  if (forbiddenCharsRegex.test(name)) {
    throw new Error(
      `Worksheet name "${name}" contains forbidden characters. Do not use structural symbols matching: \\ / : ? * [ ]`
    );
  }

  if (name.startsWith("'") || name.endsWith("'")) {
    throw new Error(
      `Worksheet name "${name}" cannot begin or terminate with a single quote character.`
    );
  }
}

/**
 * Asserts coordinate formatting parameters match alphanumeric cell range specifications (e.g. "A1:C10").
 */
export function validateA1Range(range: string): void {
  const cleanRange = range.replace(/\s/g, "");

  // Captures individual grid configurations matching: Standard Range (A1:B2) or Solo Reference (A1)
  const rangeRegex = /^[A-Z]+[0-9]+(:[A-Z]+[0-9]+)?$/i;
  if (!rangeRegex.test(cleanRange)) {
    throw new Error(
      `Malformed spreadsheet range token parameter supplied: "${range}". Enforce proper A1 matrix standards.`
    );
  }
}

/**
 * Validates maximum structural limits matching OpenXML row and column boundary counts.
 */
export function assertGridBoundaries(rowIndex: number, colIndex: number): void {
  // Maximum capacity index boundaries matching spreadsheet tracking limits: 1,048,576 rows and 16,384 columns
  if (rowIndex < 1 || rowIndex > 1048576) {
    throw new Error(
      `Row index location reference "${rowIndex}" is out of safe bounds. Must exist between 1 and 1,048,576.`
    );
  }

  if (colIndex < 0 || colIndex > 16383) {
    throw new Error(
      `Column array index position reference "${colIndex}" breaks spreadsheet limits. Cannot exceed 16,384 total width dimensions.`
    );
  }
}
