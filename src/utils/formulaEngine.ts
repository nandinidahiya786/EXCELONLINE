/**
 * @license
 * SPDX-License-Identifier: Apache-2.0
 */

import { SheetData, CellData } from '../types/excel';
import { parseCellId, parseRange, getCellsInRange } from './cellUtils';

export interface FunctionDoc {
  name: string;
  category: 'Math' | 'Statistical' | 'Logical' | 'Lookup' | 'Text' | 'Date & Time' | 'Financial';
  syntax: string;
  description: string;
  example: string;
}

export const FUNCTION_CATALOG: FunctionDoc[] = [
  // Math & Trig
  { name: 'SUM', category: 'Math', syntax: 'SUM(number1, [number2], ...)', description: 'Adds all numbers in a range of cells.', example: 'SUM(A1:A10)' },
  { name: 'AVERAGE', category: 'Statistical', syntax: 'AVERAGE(number1, [number2], ...)', description: 'Returns the average (arithmetic mean) of the arguments.', example: 'AVERAGE(B2:B8)' },
  { name: 'COUNT', category: 'Statistical', syntax: 'COUNT(val1, [val2], ...)', description: 'Counts how many cells contain numbers.', example: 'COUNT(A1:D10)' },
  { name: 'COUNTA', category: 'Statistical', syntax: 'COUNTA(val1, [val2], ...)', description: 'Counts how many cells are not empty.', example: 'COUNTA(A1:A20)' },
  { name: 'MIN', category: 'Statistical', syntax: 'MIN(number1, [number2], ...)', description: 'Returns the smallest number in a set of values.', example: 'MIN(C1:C15)' },
  { name: 'MAX', category: 'Statistical', syntax: 'MAX(number1, [number2], ...)', description: 'Returns the largest number in a set of values.', example: 'MAX(C1:C15)' },
  { name: 'MEDIAN', category: 'Statistical', syntax: 'MEDIAN(number1, [number2], ...)', description: 'Returns the median of the given numbers.', example: 'MEDIAN(A1:A9)' },
  { name: 'STDEV', category: 'Statistical', syntax: 'STDEV(number1, [number2], ...)', description: 'Estimates standard deviation based on a sample.', example: 'STDEV(B2:B20)' },
  { name: 'ROUND', category: 'Math', syntax: 'ROUND(number, num_digits)', description: 'Rounds a number to a specified number of digits.', example: 'ROUND(A1, 2)' },
  { name: 'ROUNDUP', category: 'Math', syntax: 'ROUNDUP(number, num_digits)', description: 'Rounds a number up, away from 0.', example: 'ROUNDUP(A1, 0)' },
  { name: 'ROUNDDOWN', category: 'Math', syntax: 'ROUNDDOWN(number, num_digits)', description: 'Rounds a number down, toward 0.', example: 'ROUNDDOWN(A1, 0)' },
  { name: 'ABS', category: 'Math', syntax: 'ABS(number)', description: 'Returns the absolute value of a number.', example: 'ABS(D4)' },
  { name: 'SQRT', category: 'Math', syntax: 'SQRT(number)', description: 'Returns a positive square root.', example: 'SQRT(16)' },
  { name: 'POWER', category: 'Math', syntax: 'POWER(number, power)', description: 'Returns the result of a number raised to a power.', example: 'POWER(A1, 2)' },
  { name: 'MOD', category: 'Math', syntax: 'MOD(number, divisor)', description: 'Returns the remainder after division.', example: 'MOD(10, 3)' },
  { name: 'PRODUCT', category: 'Math', syntax: 'PRODUCT(number1, [number2], ...)', description: 'Multiplies all the numbers given as arguments.', example: 'PRODUCT(A1:A5)' },
  { name: 'SUMIF', category: 'Math', syntax: 'SUMIF(range, criteria, [sum_range])', description: 'Adds cells specified by a given condition or criteria.', example: 'SUMIF(A1:A10, ">50", B1:B10)' },
  { name: 'COUNTIF', category: 'Statistical', syntax: 'COUNTIF(range, criteria)', description: 'Counts the number of cells that meet a criterion.', example: 'COUNTIF(A1:A10, "Approved")' },
  { name: 'AVERAGEIF', category: 'Statistical', syntax: 'AVERAGEIF(range, criteria, [average_range])', description: 'Computes average for cells meeting a criterion.', example: 'AVERAGEIF(B2:B20, ">0")' },

  // Logical
  { name: 'IF', category: 'Logical', syntax: 'IF(logical_test, value_if_true, [value_if_false])', description: 'Specifies a logical test to perform.', example: 'IF(A1>100, "High", "Normal")' },
  { name: 'IFS', category: 'Logical', syntax: 'IFS(test1, val1, [test2, val2], ...)', description: 'Checks whether one or more conditions are met.', example: 'IFS(A1>90, "A", A1>80, "B", TRUE, "C")' },
  { name: 'AND', category: 'Logical', syntax: 'AND(logical1, [logical2], ...)', description: 'Returns TRUE if all arguments are TRUE.', example: 'AND(A1>0, B1>0)' },
  { name: 'OR', category: 'Logical', syntax: 'OR(logical1, [logical2], ...)', description: 'Returns TRUE if any argument is TRUE.', example: 'OR(A1>100, B1>100)' },
  { name: 'NOT', category: 'Logical', syntax: 'NOT(logical)', description: 'Reverses the logic of its argument.', example: 'NOT(A1=0)' },
  { name: 'IFERROR', category: 'Logical', syntax: 'IFERROR(value, value_if_error)', description: 'Returns value_if_error if formula evaluates to an error.', example: 'IFERROR(A1/B1, 0)' },

  // Lookup & Reference
  { name: 'VLOOKUP', category: 'Lookup', syntax: 'VLOOKUP(lookup_value, table_array, col_index_num, [range_lookup])', description: 'Looks for a value in the leftmost column of a table, then returns a value in the same row from a specified column.', example: 'VLOOKUP("ID101", A1:D20, 2, FALSE)' },
  { name: 'HLOOKUP', category: 'Lookup', syntax: 'HLOOKUP(lookup_value, table_array, row_index_num, [range_lookup])', description: 'Looks for a value in top row and returns value from specified row.', example: 'HLOOKUP("Q1", A1:E5, 2, FALSE)' },
  { name: 'INDEX', category: 'Lookup', syntax: 'INDEX(array, row_num, [column_num])', description: 'Returns a value or reference of the cell at the intersection of a particular row and column.', example: 'INDEX(A1:C10, 2, 3)' },
  { name: 'MATCH', category: 'Lookup', syntax: 'MATCH(lookup_value, lookup_array, [match_type])', description: 'Returns the relative position of an item in an array that matches a specified value.', example: 'MATCH("East", A1:A10, 0)' },
  { name: 'XLOOKUP', category: 'Lookup', syntax: 'XLOOKUP(lookup_value, lookup_array, return_array, [if_not_found])', description: 'Searches a range or an array, and returns an item corresponding to the first match.', example: 'XLOOKUP(A1, B1:B10, C1:C10, "Not Found")' },

  // Text
  { name: 'CONCAT', category: 'Text', syntax: 'CONCAT(text1, [text2], ...)', description: 'Combines the text from multiple ranges and/or strings.', example: 'CONCAT(A1, " ", B1)' },
  { name: 'CONCATENATE', category: 'Text', syntax: 'CONCATENATE(text1, [text2], ...)', description: 'Joins two or more text strings into one string.', example: 'CONCATENATE(A1, B1)' },
  { name: 'LEFT', category: 'Text', syntax: 'LEFT(text, [num_chars])', description: 'Returns the first characters in a text string.', example: 'LEFT(A1, 3)' },
  { name: 'RIGHT', category: 'Text', syntax: 'RIGHT(text, [num_chars])', description: 'Returns the last characters in a text string.', example: 'RIGHT(A1, 4)' },
  { name: 'MID', category: 'Text', syntax: 'MID(text, start_num, num_chars)', description: 'Returns characters from the middle of a text string.', example: 'MID(A1, 2, 4)' },
  { name: 'LEN', category: 'Text', syntax: 'LEN(text)', description: 'Returns the number of characters in a text string.', example: 'LEN(A1)' },
  { name: 'UPPER', category: 'Text', syntax: 'UPPER(text)', description: 'Converts text to uppercase.', example: 'UPPER(A1)' },
  { name: 'LOWER', category: 'Text', syntax: 'LOWER(text)', description: 'Converts text to lowercase.', example: 'LOWER(A1)' },
  { name: 'PROPER', category: 'Text', syntax: 'PROPER(text)', description: 'Capitalizes the first letter in each word of a text value.', example: 'PROPER(A1)' },
  { name: 'TRIM', category: 'Text', syntax: 'TRIM(text)', description: 'Removes all spaces from text except for single spaces between words.', example: 'TRIM(A1)' },
  { name: 'SUBSTITUTE', category: 'Text', syntax: 'SUBSTITUTE(text, old_text, new_text)', description: 'Substitutes new_text for old_text in a text string.', example: 'SUBSTITUTE(A1, "2025", "2026")' },

  // Date & Time
  { name: 'TODAY', category: 'Date & Time', syntax: 'TODAY()', description: 'Returns the current date.', example: 'TODAY()' },
  { name: 'NOW', category: 'Date & Time', syntax: 'NOW()', description: 'Returns the current date and time.', example: 'NOW()' },
  { name: 'DATE', category: 'Date & Time', syntax: 'DATE(year, month, day)', description: 'Returns the sequential serial number that represents a particular date.', example: 'DATE(2026, 10, 30)' },
  { name: 'YEAR', category: 'Date & Time', syntax: 'YEAR(serial_number)', description: 'Returns the year of a date.', example: 'YEAR(TODAY())' },
  { name: 'MONTH', category: 'Date & Time', syntax: 'MONTH(serial_number)', description: 'Returns the month of a date (1 to 12).', example: 'MONTH(TODAY())' },
  { name: 'DAY', category: 'Date & Time', syntax: 'DAY(serial_number)', description: 'Returns the day of the month (1 to 31).', example: 'DAY(TODAY())' },

  // Financial
  { name: 'PMT', category: 'Financial', syntax: 'PMT(rate, nper, pv)', description: 'Calculates the payment for a loan based on constant payments and a constant interest rate.', example: 'PMT(0.05/12, 60, -25000)' },
  { name: 'FV', category: 'Financial', syntax: 'FV(rate, nper, pmt, [pv])', description: 'Returns the future value of an investment.', example: 'FV(0.06/12, 120, -200)' },
  { name: 'PV', category: 'Financial', syntax: 'PV(rate, nper, pmt, [fv])', description: 'Returns the present value of an investment.', example: 'PV(0.08/12, 240, 500)' },
];

/**
 * Core Formula Evaluator
 */
export class FormulaEngine {
  private sheets: Record<string, SheetData> = {};
  private activeSheetId: string = '';
  private evaluatingStack: Set<string> = new Set();
  private cache: Record<string, any> = {};

  constructor(sheets: SheetData[], activeSheetId: string) {
    this.updateContext(sheets, activeSheetId);
  }

  public updateContext(sheets: SheetData[], activeSheetId: string) {
    this.sheets = {};
    for (const sheet of sheets) {
      this.sheets[sheet.id] = sheet;
      // Also map by name for cross-sheet lookup (e.g. Sheet1!A1)
      this.sheets[sheet.name.toLowerCase()] = sheet;
    }
    this.activeSheetId = activeSheetId;
    this.evaluatingStack.clear();
    this.cache = {};
  }

  /**
   * Recalculates all formulas in all sheets and updates their computed .value
   */
  public evaluateWorkbook(): { hasChanged: boolean } {
    let hasChanged = false;
    this.cache = {};

    for (const sheet of Object.values(this.sheets)) {
      if (!sheet.cells) continue;
      for (const [cellId, cell] of Object.entries(sheet.cells)) {
        if (!cell || !cell.raw) continue;
        if (cell.raw.startsWith('=')) {
          const key = `${sheet.id}:${cellId}`;
          this.evaluatingStack.clear();
          const computed = this.evaluateCell(sheet.id, cellId);
          if (cell.value !== computed) {
            cell.value = computed;
            hasChanged = true;
          }
        } else {
          // Parse regular values (number, date, string)
          const num = Number(cell.raw);
          const computed = !isNaN(num) && cell.raw.trim() !== '' ? num : cell.raw;
          if (cell.value !== computed) {
            cell.value = computed;
            hasChanged = true;
          }
        }
      }
    }
    return { hasChanged };
  }

  /**
   * Evaluates a single cell
   */
  public evaluateCell(sheetId: string, cellId: string): any {
    const key = `${sheetId}:${cellId.toUpperCase()}`;
    if (this.evaluatingStack.has(key)) {
      return '#CYCLE!';
    }
    if (this.cache[key] !== undefined) {
      return this.cache[key];
    }

    const sheet = this.sheets[sheetId] || this.sheets[sheetId.toLowerCase()];
    if (!sheet || !sheet.cells) return 0;
    const cell = sheet.cells[cellId.toUpperCase()];
    if (!cell || cell.raw === undefined || cell.raw === '') return 0;

    const raw = String(cell.raw).trim();
    if (!raw.startsWith('=')) {
      const num = Number(raw);
      return !isNaN(num) && raw !== '' ? num : raw;
    }

    this.evaluatingStack.add(key);
    try {
      const expr = raw.substring(1).trim();
      const result = this.evaluateExpression(expr, sheetId);
      this.cache[key] = result;
      return result;
    } catch (err: any) {
      return err.message?.startsWith('#') ? err.message : '#VALUE!';
    } finally {
      this.evaluatingStack.delete(key);
    }
  }

  /**
   * Evaluates an expression string like "SUM(A1:A5) + 10" or "IF(B2>10, 'Yes', 'No')"
   */
  public evaluateExpression(expr: string, contextSheetId: string): any {
    if (!expr || expr.trim() === '') return 0;

    // Remove outer parentheses if they wrap the whole expression
    expr = expr.trim();

    // Check for string literals (e.g. "Approved")
    if ((expr.startsWith('"') && expr.endsWith('"')) || (expr.startsWith("'") && expr.endsWith("'"))) {
      return expr.slice(1, -1);
    }

    // Check for boolean literals
    if (expr.toUpperCase() === 'TRUE') return true;
    if (expr.toUpperCase() === 'FALSE') return false;

    // Check if it's a numeric literal
    if (!isNaN(Number(expr)) && expr !== '') {
      return Number(expr);
    }

    // Check if it's a single cell reference: e.g. A1, $A$1, Sheet1!A1
    const cellRefMatch = expr.match(/^((?:['"]?[\w\s]+['"]?)!)?(\$?[A-Za-z]+\$?\d+)$/);
    if (cellRefMatch) {
      const targetSheet = cellRefMatch[1] ? cellRefMatch[1].replace(/[!'"]/g, '') : contextSheetId;
      const cleanCellId = cellRefMatch[2].replace(/\$/g, '').toUpperCase();
      return this.evaluateCell(targetSheet, cleanCellId);
    }

    // Check if it is a function call: e.g. SUM(...) or VLOOKUP(...)
    const funcMatch = expr.match(/^([A-Za-z_][A-Za-z0-9_]*)\s*\((.*)\)$/s);
    if (funcMatch) {
      // Ensure the closing parenthesis corresponds to the opening parenthesis
      if (this.isBalancedCall(expr)) {
        const funcName = funcMatch[1].toUpperCase();
        const argsStr = funcMatch[2];
        const args = this.parseArguments(argsStr);
        return this.executeFunction(funcName, args, contextSheetId);
      }
    }

    // Binary Operators (sorted by precedence: comparisons, +, -, *, /, ^, &)
    const opSplit = this.findLowestPrecedenceOperator(expr);
    if (opSplit) {
      const leftVal = this.evaluateExpression(opSplit.left, contextSheetId);
      const rightVal = this.evaluateExpression(opSplit.right, contextSheetId);
      return this.applyOperator(opSplit.op, leftVal, rightVal);
    }

    return '#VALUE!';
  }

  private isBalancedCall(str: string): boolean {
    const firstParen = str.indexOf('(');
    if (firstParen === -1 || !str.endsWith(')')) return false;

    let depth = 0;
    let inQuotes = false;
    let quoteChar = '';

    for (let i = 0; i < str.length; i++) {
      const ch = str[i];
      if ((ch === '"' || ch === "'") && (i === 0 || str[i - 1] !== '\\')) {
        if (!inQuotes) {
          inQuotes = true;
          quoteChar = ch;
        } else if (ch === quoteChar) {
          inQuotes = false;
        }
      } else if (!inQuotes) {
        if (ch === '(') depth++;
        else if (ch === ')') {
          depth--;
          if (depth === 0 && i !== str.length - 1) return false;
        }
      }
    }
    return depth === 0;
  }

  private parseArguments(argsStr: string): string[] {
    const args: string[] = [];
    if (!argsStr || argsStr.trim() === '') return args;

    let current = '';
    let parenDepth = 0;
    let inQuotes = false;
    let quoteChar = '';

    for (let i = 0; i < argsStr.length; i++) {
      const ch = argsStr[i];
      if ((ch === '"' || ch === "'") && (i === 0 || argsStr[i - 1] !== '\\')) {
        if (!inQuotes) {
          inQuotes = true;
          quoteChar = ch;
        } else if (ch === quoteChar) {
          inQuotes = false;
        }
        current += ch;
      } else if (!inQuotes) {
        if (ch === '(') {
          parenDepth++;
          current += ch;
        } else if (ch === ')') {
          parenDepth--;
          current += ch;
        } else if (ch === ',' && parenDepth === 0) {
          args.push(current.trim());
          current = '';
        } else {
          current += ch;
        }
      } else {
        current += ch;
      }
    }
    if (current.trim() !== '') {
      args.push(current.trim());
    }
    return args;
  }

  private findLowestPrecedenceOperator(expr: string): { op: string; left: string; right: string } | null {
    // Precedence:
    // 1: Comparisons: =, <>, >=, <=, >, <
    // 2: String concat: &
    // 3: Add / Subtract: +, -
    // 4: Multiply / Divide: *, /
    // 5: Power: ^

    const operatorGroups = [
      ['=', '<>', '>=', '<=', '>', '<'],
      ['&'],
      ['+', '-'],
      ['*', '/'],
      ['^']
    ];

    for (const group of operatorGroups) {
      let depth = 0;
      let inQuotes = false;
      let quoteChar = '';

      // Right-to-left scan for left-associative operators
      for (let i = expr.length - 1; i >= 0; i--) {
        const ch = expr[i];

        if ((ch === '"' || ch === "'") && (i === 0 || expr[i - 1] !== '\\')) {
          if (!inQuotes) {
            inQuotes = true;
            quoteChar = ch;
          } else if (ch === quoteChar) {
            inQuotes = false;
          }
        } else if (!inQuotes) {
          if (ch === ')') parenMatchBackwards: depth++;
          else if (ch === '(') depth--;
          else if (depth === 0) {
            // Check for multi-char operators first: <>, >=, <=
            if (i > 0) {
              const twoChar = expr.substring(i - 1, i + 1);
              if (group.includes(twoChar)) {
                return {
                  op: twoChar,
                  left: expr.substring(0, i - 1).trim(),
                  right: expr.substring(i + 1).trim(),
                };
              }
            }
            if (group.includes(ch)) {
              // Ensure it's not a leading negative sign like "-5"
              if (ch === '-' && (i === 0 || ['+', '-', '*', '/', '(', ',', '=', '<', '>'].includes(expr[i - 1].trim()))) {
                continue;
              }
              return {
                op: ch,
                left: expr.substring(0, i).trim(),
                right: expr.substring(i + 1).trim(),
              };
            }
          }
        }
      }
    }

    return null;
  }

  private applyOperator(op: string, left: any, right: any): any {
    if (typeof left === 'string' && left.startsWith('#')) return left;
    if (typeof right === 'string' && right.startsWith('#')) return right;

    if (op === '&') {
      return `${left ?? ''}${right ?? ''}`;
    }

    // Comparison operators
    if (op === '=') return left === right || (String(left).toLowerCase() === String(right).toLowerCase() && !isNaN(Number(left)) && Number(left) === Number(right));
    if (op === '<>') return left !== right;
    if (op === '>') return Number(left) > Number(right);
    if (op === '<') return Number(left) < Number(right);
    if (op === '>=') return Number(left) >= Number(right);
    if (op === '<=') return Number(left) <= Number(right);

    const numLeft = Number(left);
    const numRight = Number(right);

    if (isNaN(numLeft) || isNaN(numRight)) return '#VALUE!';

    switch (op) {
      case '+': return numLeft + numRight;
      case '-': return numLeft - numRight;
      case '*': return numLeft * numRight;
      case '/': return numRight === 0 ? '#DIV/0!' : numLeft / numRight;
      case '^': return Math.pow(numLeft, numRight);
      default: return '#VALUE!';
    }
  }

  /**
   * Resolves arguments that could be a single value or a range (e.g. A1:B10)
   */
  private resolveRangeOrValues(arg: string, contextSheetId: string): any[] {
    const rangeMatch = arg.match(/^((?:['"]?[\w\s]+['"]?)!)?([A-Za-z]+\d+):([A-Za-z]+\d+)$/);
    if (rangeMatch) {
      const sheetId = rangeMatch[1] ? rangeMatch[1].replace(/[!'"]/g, '') : contextSheetId;
      const range = parseRange(`${rangeMatch[2]}:${rangeMatch[3]}`);
      if (!range) return [];

      const cellIds = getCellsInRange(range);
      return cellIds.map(cId => this.evaluateCell(sheetId, cId));
    }

    // Else evaluate single expression
    const val = this.evaluateExpression(arg, contextSheetId);
    return [val];
  }

  private flattenNumbers(args: string[], contextSheetId: string): number[] {
    const nums: number[] = [];
    for (const arg of args) {
      const values = this.resolveRangeOrValues(arg, contextSheetId);
      for (const v of values) {
        if (typeof v === 'number' && !isNaN(v)) {
          nums.push(v);
        } else if (typeof v === 'string' && v.trim() !== '' && !isNaN(Number(v))) {
          nums.push(Number(v));
        }
      }
    }
    return nums;
  }

  private executeFunction(name: string, args: string[], contextSheetId: string): any {
    switch (name) {
      case 'SUM': {
        const nums = this.flattenNumbers(args, contextSheetId);
        return nums.reduce((a, b) => a + b, 0);
      }

      case 'AVERAGE': {
        const nums = this.flattenNumbers(args, contextSheetId);
        if (nums.length === 0) return '#DIV/0!';
        return nums.reduce((a, b) => a + b, 0) / nums.length;
      }

      case 'COUNT': {
        const nums = this.flattenNumbers(args, contextSheetId);
        return nums.length;
      }

      case 'COUNTA': {
        let count = 0;
        for (const arg of args) {
          const values = this.resolveRangeOrValues(arg, contextSheetId);
          for (const v of values) {
            if (v !== '' && v !== null && v !== undefined && v !== 0) count++;
            else if (v === 0) count++;
          }
        }
        return count;
      }

      case 'COUNTBLANK': {
        let count = 0;
        for (const arg of args) {
          const values = this.resolveRangeOrValues(arg, contextSheetId);
          for (const v of values) {
            if (v === '' || v === null || v === undefined) count++;
          }
        }
        return count;
      }

      case 'MIN': {
        const nums = this.flattenNumbers(args, contextSheetId);
        return nums.length > 0 ? Math.min(...nums) : 0;
      }

      case 'MAX': {
        const nums = this.flattenNumbers(args, contextSheetId);
        return nums.length > 0 ? Math.max(...nums) : 0;
      }

      case 'MEDIAN': {
        const nums = this.flattenNumbers(args, contextSheetId).sort((a, b) => a - b);
        if (nums.length === 0) return 0;
        const mid = Math.floor(nums.length / 2);
        return nums.length % 2 !== 0 ? nums[mid] : (nums[mid - 1] + nums[mid]) / 2;
      }

      case 'STDEV': {
        const nums = this.flattenNumbers(args, contextSheetId);
        if (nums.length < 2) return 0;
        const mean = nums.reduce((a, b) => a + b, 0) / nums.length;
        const variance = nums.reduce((acc, val) => acc + Math.pow(val - mean, 2), 0) / (nums.length - 1);
        return Math.sqrt(variance);
      }

      case 'ROUND': {
        const num = Number(this.evaluateExpression(args[0], contextSheetId));
        const digits = args[1] ? Number(this.evaluateExpression(args[1], contextSheetId)) : 0;
        if (isNaN(num)) return '#VALUE!';
        const factor = Math.pow(10, digits);
        return Math.round(num * factor) / factor;
      }

      case 'ROUNDUP': {
        const num = Number(this.evaluateExpression(args[0], contextSheetId));
        const digits = args[1] ? Number(this.evaluateExpression(args[1], contextSheetId)) : 0;
        const factor = Math.pow(10, digits);
        return Math.ceil(num * factor) / factor;
      }

      case 'ROUNDDOWN': {
        const num = Number(this.evaluateExpression(args[0], contextSheetId));
        const digits = args[1] ? Number(this.evaluateExpression(args[1], contextSheetId)) : 0;
        const factor = Math.pow(10, digits);
        return Math.floor(num * factor) / factor;
      }

      case 'ABS': {
        const num = Number(this.evaluateExpression(args[0], contextSheetId));
        return isNaN(num) ? '#VALUE!' : Math.abs(num);
      }

      case 'SQRT': {
        const num = Number(this.evaluateExpression(args[0], contextSheetId));
        return num < 0 ? '#NUM!' : Math.sqrt(num);
      }

      case 'POWER': {
        const base = Number(this.evaluateExpression(args[0], contextSheetId));
        const exp = Number(this.evaluateExpression(args[1], contextSheetId));
        return Math.pow(base, exp);
      }

      case 'MOD': {
        const num = Number(this.evaluateExpression(args[0], contextSheetId));
        const div = Number(this.evaluateExpression(args[1], contextSheetId));
        return div === 0 ? '#DIV/0!' : num % div;
      }

      case 'PRODUCT': {
        const nums = this.flattenNumbers(args, contextSheetId);
        return nums.length > 0 ? nums.reduce((a, b) => a * b, 1) : 0;
      }

      case 'RAND': {
        return Math.random();
      }

      case 'RANDBETWEEN': {
        const min = Number(this.evaluateExpression(args[0], contextSheetId));
        const max = Number(this.evaluateExpression(args[1], contextSheetId));
        return Math.floor(Math.random() * (max - min + 1)) + min;
      }

      case 'IF': {
        const condition = this.evaluateExpression(args[0], contextSheetId);
        if (condition && condition !== '0' && condition !== 0 && condition !== '#VALUE!' && condition !== false) {
          return args[1] ? this.evaluateExpression(args[1], contextSheetId) : true;
        } else {
          return args[2] ? this.evaluateExpression(args[2], contextSheetId) : false;
        }
      }

      case 'IFS': {
        for (let i = 0; i < args.length; i += 2) {
          const cond = this.evaluateExpression(args[i], contextSheetId);
          if (cond && cond !== '0' && cond !== 0 && cond !== false) {
            return args[i + 1] ? this.evaluateExpression(args[i + 1], contextSheetId) : true;
          }
        }
        return '#N/A';
      }

      case 'AND': {
        for (const arg of args) {
          const val = this.evaluateExpression(arg, contextSheetId);
          if (!val || val === 0 || val === '0' || val === false) return false;
        }
        return true;
      }

      case 'OR': {
        for (const arg of args) {
          const val = this.evaluateExpression(arg, contextSheetId);
          if (val && val !== 0 && val !== '0' && val !== false) return true;
        }
        return false;
      }

      case 'NOT': {
        const val = this.evaluateExpression(args[0], contextSheetId);
        return !val;
      }

      case 'IFERROR': {
        try {
          const val = this.evaluateExpression(args[0], contextSheetId);
          if (typeof val === 'string' && val.startsWith('#')) {
            return args[1] ? this.evaluateExpression(args[1], contextSheetId) : 0;
          }
          return val;
        } catch {
          return args[1] ? this.evaluateExpression(args[1], contextSheetId) : 0;
        }
      }

      case 'COUNTIF': {
        const values = this.resolveRangeOrValues(args[0], contextSheetId);
        const criteria = this.evaluateExpression(args[1], contextSheetId);
        return this.filterByCriteria(values, criteria).length;
      }

      case 'SUMIF': {
        const rangeVals = this.resolveRangeOrValues(args[0], contextSheetId);
        const criteria = this.evaluateExpression(args[1], contextSheetId);
        const sumVals = args[2] ? this.resolveRangeOrValues(args[2], contextSheetId) : rangeVals;
        let sum = 0;
        for (let i = 0; i < rangeVals.length; i++) {
          if (this.matchesCriteria(rangeVals[i], criteria)) {
            const num = Number(sumVals[i]);
            if (!isNaN(num)) sum += num;
          }
        }
        return sum;
      }

      case 'AVERAGEIF': {
        const rangeVals = this.resolveRangeOrValues(args[0], contextSheetId);
        const criteria = this.evaluateExpression(args[1], contextSheetId);
        const avgVals = args[2] ? this.resolveRangeOrValues(args[2], contextSheetId) : rangeVals;
        let sum = 0;
        let count = 0;
        for (let i = 0; i < rangeVals.length; i++) {
          if (this.matchesCriteria(rangeVals[i], criteria)) {
            const num = Number(avgVals[i]);
            if (!isNaN(num)) {
              sum += num;
              count++;
            }
          }
        }
        return count > 0 ? sum / count : '#DIV/0!';
      }

      // Lookup functions
      case 'VLOOKUP': {
        const lookupVal = this.evaluateExpression(args[0], contextSheetId);
        const rangeStr = args[1];
        const colIndex = Number(this.evaluateExpression(args[2], contextSheetId));
        if (isNaN(colIndex) || colIndex < 1) return '#VALUE!';

        const range = parseRange(rangeStr);
        if (!range) return '#REF!';

        const sheet = this.sheets[contextSheetId];
        if (!sheet) return '#REF!';

        for (let r = range.startRow; r <= range.endRow; r++) {
          const firstCellId = `${this.colLetter(range.startCol)}${r + 1}`;
          const firstVal = sheet.cells?.[firstCellId]?.value ?? sheet.cells?.[firstCellId]?.raw;

          if (this.looseEqual(firstVal, lookupVal)) {
            const targetCol = range.startCol + colIndex - 1;
            if (targetCol > range.endCol) return '#REF!';
            const targetCellId = `${this.colLetter(targetCol)}${r + 1}`;
            return sheet.cells?.[targetCellId]?.value ?? sheet.cells?.[targetCellId]?.raw ?? '';
          }
        }
        return '#N/A';
      }

      case 'HLOOKUP': {
        const lookupVal = this.evaluateExpression(args[0], contextSheetId);
        const range = parseRange(args[1]);
        const rowIndex = Number(this.evaluateExpression(args[2], contextSheetId));
        if (!range || isNaN(rowIndex) || rowIndex < 1) return '#VALUE!';

        const sheet = this.sheets[contextSheetId];
        if (!sheet) return '#REF!';

        for (let c = range.startCol; c <= range.endCol; c++) {
          const topCellId = `${this.colLetter(c)}${range.startRow + 1}`;
          const topVal = sheet.cells?.[topCellId]?.value ?? sheet.cells?.[topCellId]?.raw;

          if (this.looseEqual(topVal, lookupVal)) {
            const targetRow = range.startRow + rowIndex - 1;
            if (targetRow > range.endRow) return '#REF!';
            const targetCellId = `${this.colLetter(c)}${targetRow + 1}`;
            return sheet.cells?.[targetCellId]?.value ?? sheet.cells?.[targetCellId]?.raw ?? '';
          }
        }
        return '#N/A';
      }

      case 'INDEX': {
        const range = parseRange(args[0]);
        const rowNum = Number(this.evaluateExpression(args[1], contextSheetId));
        const colNum = args[2] ? Number(this.evaluateExpression(args[2], contextSheetId)) : 1;
        if (!range || isNaN(rowNum)) return '#REF!';

        const targetRow = range.startRow + rowNum - 1;
        const targetCol = range.startCol + colNum - 1;

        if (targetRow > range.endRow || targetCol > range.endCol) return '#REF!';
        const cellId = `${this.colLetter(targetCol)}${targetRow + 1}`;
        const sheet = this.sheets[contextSheetId];
        return sheet?.cells?.[cellId]?.value ?? sheet?.cells?.[cellId]?.raw ?? '';
      }

      case 'MATCH': {
        const lookupVal = this.evaluateExpression(args[0], contextSheetId);
        const values = this.resolveRangeOrValues(args[1], contextSheetId);
        for (let i = 0; i < values.length; i++) {
          if (this.looseEqual(values[i], lookupVal)) {
            return i + 1; // 1-based index in Excel
          }
        }
        return '#N/A';
      }

      case 'XLOOKUP': {
        const lookupVal = this.evaluateExpression(args[0], contextSheetId);
        const lookupRangeVals = this.resolveRangeOrValues(args[1], contextSheetId);
        const returnRangeVals = this.resolveRangeOrValues(args[2], contextSheetId);
        const ifNotFound = args[3] ? this.evaluateExpression(args[3], contextSheetId) : '#N/A';

        for (let i = 0; i < lookupRangeVals.length; i++) {
          if (this.looseEqual(lookupRangeVals[i], lookupVal)) {
            return returnRangeVals[i] ?? '';
          }
        }
        return ifNotFound;
      }

      // Text Functions
      case 'CONCAT':
      case 'CONCATENATE': {
        let result = '';
        for (const arg of args) {
          const vals = this.resolveRangeOrValues(arg, contextSheetId);
          for (const v of vals) {
            result += String(v ?? '');
          }
        }
        return result;
      }

      case 'LEFT': {
        const text = String(this.evaluateExpression(args[0], contextSheetId) ?? '');
        const len = args[1] ? Number(this.evaluateExpression(args[1], contextSheetId)) : 1;
        return text.substring(0, len);
      }

      case 'RIGHT': {
        const text = String(this.evaluateExpression(args[0], contextSheetId) ?? '');
        const len = args[1] ? Number(this.evaluateExpression(args[1], contextSheetId)) : 1;
        return text.substring(Math.max(0, text.length - len));
      }

      case 'MID': {
        const text = String(this.evaluateExpression(args[0], contextSheetId) ?? '');
        const start = Number(this.evaluateExpression(args[1], contextSheetId));
        const len = Number(this.evaluateExpression(args[2], contextSheetId));
        return text.substring(Math.max(0, start - 1), Math.max(0, start - 1) + len);
      }

      case 'LEN': {
        const text = String(this.evaluateExpression(args[0], contextSheetId) ?? '');
        return text.length;
      }

      case 'UPPER': {
        const text = String(this.evaluateExpression(args[0], contextSheetId) ?? '');
        return text.toUpperCase();
      }

      case 'LOWER': {
        const text = String(this.evaluateExpression(args[0], contextSheetId) ?? '');
        return text.toLowerCase();
      }

      case 'PROPER': {
        const text = String(this.evaluateExpression(args[0], contextSheetId) ?? '');
        return text.replace(/\b\w/g, l => l.toUpperCase());
      }

      case 'TRIM': {
        const text = String(this.evaluateExpression(args[0], contextSheetId) ?? '');
        return text.trim().replace(/\s+/g, ' ');
      }

      case 'SUBSTITUTE': {
        const text = String(this.evaluateExpression(args[0], contextSheetId) ?? '');
        const oldText = String(this.evaluateExpression(args[1], contextSheetId) ?? '');
        const newText = String(this.evaluateExpression(args[2], contextSheetId) ?? '');
        return text.split(oldText).join(newText);
      }

      // Date / Time
      case 'TODAY': {
        const d = new Date();
        return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}-${String(d.getDate()).padStart(2, '0')}`;
      }

      case 'NOW': {
        const d = new Date();
        return d.toISOString().replace('T', ' ').substring(0, 19);
      }

      case 'DATE': {
        const y = Number(this.evaluateExpression(args[0], contextSheetId));
        const m = Number(this.evaluateExpression(args[1], contextSheetId));
        const d = Number(this.evaluateExpression(args[2], contextSheetId));
        return `${y}-${String(m).padStart(2, '0')}-${String(d).padStart(2, '0')}`;
      }

      case 'YEAR': {
        const dStr = String(this.evaluateExpression(args[0], contextSheetId));
        const d = new Date(dStr);
        return isNaN(d.getFullYear()) ? '#VALUE!' : d.getFullYear();
      }

      case 'MONTH': {
        const dStr = String(this.evaluateExpression(args[0], contextSheetId));
        const d = new Date(dStr);
        return isNaN(d.getMonth()) ? '#VALUE!' : d.getMonth() + 1;
      }

      case 'DAY': {
        const dStr = String(this.evaluateExpression(args[0], contextSheetId));
        const d = new Date(dStr);
        return isNaN(d.getDate()) ? '#VALUE!' : d.getDate();
      }

      // Financial
      case 'PMT': {
        const rate = Number(this.evaluateExpression(args[0], contextSheetId));
        const nper = Number(this.evaluateExpression(args[1], contextSheetId));
        const pv = Number(this.evaluateExpression(args[2], contextSheetId));
        if (rate === 0) return -(pv / nper);
        const pvif = Math.pow(1 + rate, nper);
        return (rate * pv * pvif) / (1 - pvif);
      }

      case 'FV': {
        const rate = Number(this.evaluateExpression(args[0], contextSheetId));
        const nper = Number(this.evaluateExpression(args[1], contextSheetId));
        const pmt = Number(this.evaluateExpression(args[2], contextSheetId));
        const pv = args[3] ? Number(this.evaluateExpression(args[3], contextSheetId)) : 0;
        if (rate === 0) return -(pv + pmt * nper);
        return -(pv * Math.pow(1 + rate, nper) + pmt * ((Math.pow(1 + rate, nper) - 1) / rate));
      }

      default:
        return '#NAME?';
    }
  }

  private matchesCriteria(val: any, criteria: any): boolean {
    if (criteria === undefined || criteria === null) return false;
    const strCriteria = String(criteria).trim();

    if (strCriteria.startsWith('>=')) {
      return Number(val) >= Number(strCriteria.substring(2));
    }
    if (strCriteria.startsWith('<=')) {
      return Number(val) <= Number(strCriteria.substring(2));
    }
    if (strCriteria.startsWith('>')) {
      return Number(val) > Number(strCriteria.substring(1));
    }
    if (strCriteria.startsWith('<')) {
      return Number(val) < Number(strCriteria.substring(1));
    }
    if (strCriteria.startsWith('<>')) {
      return !this.looseEqual(val, strCriteria.substring(2));
    }
    if (strCriteria.startsWith('=')) {
      return this.looseEqual(val, strCriteria.substring(1));
    }
    return this.looseEqual(val, strCriteria);
  }

  private filterByCriteria(values: any[], criteria: any): any[] {
    return values.filter(v => this.matchesCriteria(v, criteria));
  }

  private looseEqual(a: any, b: any): boolean {
    if (a === b) return true;
    if (a === undefined || b === undefined) return false;
    if (typeof a === 'number' && !isNaN(Number(b))) return a === Number(b);
    return String(a).toLowerCase().trim() === String(b).toLowerCase().trim();
  }

  private colLetter(colIndex: number): string {
    let letter = '';
    let temp = colIndex;
    while (temp >= 0) {
      letter = String.fromCharCode((temp % 26) + 65) + letter;
      temp = Math.floor(temp / 26) - 1;
    }
    return letter;
  }
}
