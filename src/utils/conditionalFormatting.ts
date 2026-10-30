/**
 * @license
 * SPDX-License-Identifier: Apache-2.0
 */

import { ConditionalRule, SheetData } from '../types/excel';
import { parseRange, isCellInRange } from './cellUtils';

export interface StyleOverride {
  bgColor?: string;
  textColor?: string;
}

export function evaluateConditionalFormat(
  row: number,
  col: number,
  val: any,
  sheet: SheetData
): StyleOverride | null {
  if (!sheet.conditionalRules || sheet.conditionalRules.length === 0) return null;

  for (const rule of sheet.conditionalRules) {
    const range = parseRange(rule.range);
    if (!range) continue;
    if (!isCellInRange(row, col, range)) continue;

    const numVal = Number(val);
    const isNum = !isNaN(numVal) && typeof val !== 'boolean' && val !== '';

    switch (rule.type) {
      case 'greaterThan': {
        const threshold = Number(rule.value1);
        if (isNum && numVal > threshold) {
          return {
            bgColor: rule.color || '#fee2e2', // light red
            textColor: rule.textColor || '#991b1b', // dark red
          };
        }
        break;
      }

      case 'lessThan': {
        const threshold = Number(rule.value1);
        if (isNum && numVal < threshold) {
          return {
            bgColor: rule.color || '#dcfce7', // light green
            textColor: rule.textColor || '#166534', // dark green
          };
        }
        break;
      }

      case 'between': {
        const min = Number(rule.value1);
        const max = Number(rule.value2);
        if (isNum && numVal >= min && numVal <= max) {
          return {
            bgColor: rule.color || '#fef3c7', // light amber
            textColor: rule.textColor || '#92400e', // dark amber
          };
        }
        break;
      }

      case 'equalTo': {
        if (String(val).toLowerCase() === String(rule.value1).toLowerCase()) {
          return {
            bgColor: rule.color || '#e0f2fe', // light blue
            textColor: rule.textColor || '#075985',
          };
        }
        break;
      }

      case 'textContains': {
        if (String(val).toLowerCase().includes(String(rule.value1).toLowerCase())) {
          return {
            bgColor: rule.color || '#fef08a',
            textColor: rule.textColor || '#854d0e',
          };
        }
        break;
      }

      case 'colorScale': {
        if (!isNum) break;
        // Compute min and max across the range
        let minVal = Infinity;
        let maxVal = -Infinity;
        for (let r = range.startRow; r <= range.endRow; r++) {
          for (let c = range.startCol; c <= range.endCol; c++) {
            const cellId = `${String.fromCharCode(65 + c)}${r + 1}`;
            const cVal = Number(sheet.cells?.[cellId]?.value ?? sheet.cells?.[cellId]?.raw);
            if (!isNaN(cVal)) {
              if (cVal < minVal) minVal = cVal;
              if (cVal > maxVal) maxVal = cVal;
            }
          }
        }
        if (minVal >= maxVal) break;

        const ratio = Math.max(0, Math.min(1, (numVal - minVal) / (maxVal - minVal)));
        // Green-Yellow-Red scale: 0 is soft red (#f87171), 0.5 is soft yellow (#fef08a), 1 is soft green (#86efac)
        let rColor: number, gColor: number, bColor: number;
        if (ratio < 0.5) {
          const t = ratio * 2;
          rColor = Math.round(254 - (254 - 250) * t);
          gColor = Math.round(202 + (240 - 202) * t);
          bColor = Math.round(202 + (138 - 202) * t);
        } else {
          const t = (ratio - 0.5) * 2;
          rColor = Math.round(250 - (250 - 134) * t);
          gColor = Math.round(240 - (240 - 239) * t);
          bColor = Math.round(138 + (172 - 138) * t);
        }
        return {
          bgColor: `rgb(${rColor}, ${gColor}, ${bColor})`,
          textColor: '#0f172a',
        };
      }
    }
  }

  return null;
}
