/**
 * @license
 * SPDX-License-Identifier: Apache-2.0
 */

import { CellCoordinate, CellRange } from '../types/excel';

/**
 * Convert 0-based column index to Excel column string (0 -> 'A', 25 -> 'Z', 26 -> 'AA', etc.)
 */
export function colIndexToLetter(colIndex: number): string {
  let letter = '';
  let temp = colIndex;
  while (temp >= 0) {
    letter = String.fromCharCode((temp % 26) + 65) + letter;
    temp = Math.floor(temp / 26) - 1;
  }
  return letter;
}

/**
 * Convert Excel column letters to 0-based index ('A' -> 0, 'Z' -> 25, 'AA' -> 26)
 */
export function letterToColIndex(letters: string): number {
  const upper = letters.toUpperCase();
  let col = 0;
  for (let i = 0; i < upper.length; i++) {
    col = col * 26 + (upper.charCodeAt(i) - 64);
  }
  return col - 1;
}

/**
 * Convert 0-based row and column to cell ID, e.g. (0, 0) -> "A1"
 */
export function coordsToCellId(row: number, col: number): string {
  return `${colIndexToLetter(col)}${row + 1}`;
}

/**
 * Parse cell ID like "A1", "$B$4", or "C$10" to coordinates
 */
export function parseCellId(cellId: string): { row: number; col: number; isColAbsolute: boolean; isRowAbsolute: boolean } | null {
  const match = cellId.trim().match(/^(\$?)([A-Za-z]+)(\$?)(\d+)$/);
  if (!match) return null;
  const isColAbsolute = match[1] === '$';
  const col = letterToColIndex(match[2]);
  const isRowAbsolute = match[3] === '$';
  const row = parseInt(match[4], 10) - 1;
  return { row, col, isColAbsolute, isRowAbsolute };
}

/**
 * Parse a range string like "A1:C5" into a CellRange
 */
export function parseRange(rangeStr: string): CellRange | null {
  const parts = rangeStr.split(':');
  if (parts.length === 1) {
    const coord = parseCellId(parts[0]);
    if (!coord) return null;
    return {
      startRow: coord.row,
      startCol: coord.col,
      endRow: coord.row,
      endCol: coord.col,
    };
  }
  if (parts.length === 2) {
    const c1 = parseCellId(parts[0]);
    const c2 = parseCellId(parts[1]);
    if (!c1 || !c2) return null;
    return {
      startRow: Math.min(c1.row, c2.row),
      startCol: Math.min(c1.col, c2.col),
      endRow: Math.max(c1.row, c2.row),
      endCol: Math.max(c1.col, c2.col),
    };
  }
  return null;
}

/**
 * Format a CellRange as "A1:C5" or "A1"
 */
export function formatRange(range: CellRange): string {
  const startId = coordsToCellId(range.startRow, range.startCol);
  const endId = coordsToCellId(range.endRow, range.endCol);
  if (startId === endId) return startId;
  return `${startId}:${endId}`;
}

/**
 * Check if a cell (row, col) is within a range
 */
export function isCellInRange(row: number, col: number, range: CellRange): boolean {
  return (
    row >= range.startRow &&
    row <= range.endRow &&
    col >= range.startCol &&
    col <= range.endCol
  );
}

/**
 * Return all cell IDs inside a range
 */
export function getCellsInRange(range: CellRange): string[] {
  const cells: string[] = [];
  for (let r = range.startRow; r <= range.endRow; r++) {
    for (let c = range.startCol; c <= range.endCol; c++) {
      cells.push(coordsToCellId(r, c));
    }
  }
  return cells;
}

/**
 * Shifts relative cell references in a formula string by deltaRow and deltaCol
 * e.g. "=A1+B1" shifted by (1, 0) becomes "=A2+B2"
 * "$A$1" stays "$A$1", "A$1" becomes "A$2" only row is shifted
 */
export function shiftFormula(formula: string, deltaRow: number, deltaCol: number): string {
  if (!formula.startsWith('=')) return formula;

  // Regex to match cell references with optional sheet name and $ signs
  // e.g. Sheet1!$A$1 or A1 or $B3 or C$4
  return formula.replace(/(?:(['"]?[\w\s]+['"]?)!)?(\$?)([A-Za-z]+)(\$?)(\d+)/g, (match, sheet, colAbs, colLetters, rowAbs, rowNum) => {
    let col = letterToColIndex(colLetters);
    let row = parseInt(rowNum, 10) - 1;

    if (!colAbs) {
      col = Math.max(0, col + deltaCol);
    }
    if (!rowAbs) {
      row = Math.max(0, row + deltaRow);
    }

    const newColStr = colAbs ? `$${colIndexToLetter(col)}` : colIndexToLetter(col);
    const newRowStr = rowAbs ? `$${row + 1}` : `${row + 1}`;
    const sheetPrefix = sheet ? `${sheet}!` : '';

    return `${sheetPrefix}${newColStr}${newRowStr}`;
  });
}
