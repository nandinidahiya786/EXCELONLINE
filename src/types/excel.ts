/**
 * @license
 * SPDX-License-Identifier: Apache-2.0
 */

export type NumberFormatType = 
  | 'general'
  | 'number'
  | 'currency'
  | 'accounting'
  | 'percent'
  | 'date'
  | 'scientific'
  | 'text';

export type TextAlign = 'left' | 'center' | 'right';
export type VerticalAlign = 'top' | 'middle' | 'bottom';

export interface CellBorders {
  top?: string;
  bottom?: string;
  left?: string;
  right?: string;
}

export interface CellFormat {
  bold?: boolean;
  italic?: boolean;
  underline?: boolean;
  strikethrough?: boolean;
  fontSize?: number;
  fontFamily?: string;
  textColor?: string;
  bgColor?: string;
  align?: TextAlign;
  verticalAlign?: VerticalAlign;
  wrapText?: boolean;
  numberFormat?: NumberFormatType;
  currencySymbol?: string;
  decimals?: number;
  borders?: CellBorders;
}

export interface CellData {
  raw: string; // e.g. "=SUM(A1:A5)" or "4500" or "Revenue"
  value?: any; // computed value (number, string, boolean, date, error)
  format?: CellFormat;
  comment?: string;
}

export interface CellCoordinate {
  row: number; // 0-based
  col: number; // 0-based
}

export interface CellRange {
  startRow: number;
  startCol: number;
  endRow: number;
  endCol: number;
}

export interface ConditionalRule {
  id: string;
  range: string; // e.g. "B2:B20"
  type: 'greaterThan' | 'lessThan' | 'between' | 'equalTo' | 'textContains' | 'colorScale';
  value1?: number | string;
  value2?: number;
  color?: string; // highlight color or scale type: 'green-red' | 'blue-white'
  textColor?: string;
}

export interface DataValidationRule {
  type: 'list' | 'wholeNumber' | 'decimal';
  options?: string[]; // for dropdown list
  min?: number;
  max?: number;
  allowBlank?: boolean;
}

export interface SheetData {
  id: string;
  name: string;
  cells: Record<string, CellData>; // key like "A1", "C12"
  rowCount: number;
  colCount: number;
  colWidths: Record<number, number>;
  rowHeights: Record<number, number>;
  freezeRows: number;
  freezeCols: number;
  tabColor?: string;
  conditionalRules: ConditionalRule[];
  validationRules: Record<string, DataValidationRule>; // cellId or range
  filters?: {
    range: CellRange;
    activeCol?: number;
    criteria?: string;
  };
}

export interface WorkbookData {
  fileName: string;
  activeSheetId: string;
  sheets: SheetData[];
}
