/**
 * @license
 * SPDX-License-Identifier: Apache-2.0
 */

import { CellFormat } from '../types/excel';

export function formatCellValue(value: any, format?: CellFormat): string {
  if (value === undefined || value === null) return '';
  if (typeof value === 'string' && value.startsWith('#')) return value; // Error state like #DIV/0!

  if (!format) {
    if (typeof value === 'number') {
      return Number.isInteger(value) ? value.toLocaleString() : value.toString();
    }
    return String(value);
  }

  const num = typeof value === 'number' ? value : Number(value);
  const isNumeric = !isNaN(num) && typeof value !== 'boolean' && value !== '';

  const decimals = format.decimals !== undefined ? format.decimals : 2;
  const currencySymbol = format.currencySymbol || '$';

  switch (format.numberFormat) {
    case 'number': {
      if (!isNumeric) return String(value);
      return num.toLocaleString(undefined, {
        minimumFractionDigits: decimals,
        maximumFractionDigits: decimals,
      });
    }

    case 'currency': {
      if (!isNumeric) return String(value);
      const formattedNum = Math.abs(num).toLocaleString(undefined, {
        minimumFractionDigits: decimals,
        maximumFractionDigits: decimals,
      });
      return num < 0 ? `-${currencySymbol}${formattedNum}` : `${currencySymbol}${formattedNum}`;
    }

    case 'accounting': {
      if (!isNumeric) return String(value);
      const formattedNum = Math.abs(num).toLocaleString(undefined, {
        minimumFractionDigits: decimals,
        maximumFractionDigits: decimals,
      });
      return num < 0 ? `(${currencySymbol} ${formattedNum})` : `${currencySymbol} ${formattedNum} `;
    }

    case 'percent': {
      if (!isNumeric) return String(value);
      const percentVal = num * 100;
      return `${percentVal.toFixed(decimals)}%`;
    }

    case 'scientific': {
      if (!isNumeric) return String(value);
      return num.toExponential(decimals).toUpperCase();
    }

    case 'date': {
      if (value instanceof Date) {
        return value.toLocaleDateString();
      }
      const d = new Date(value);
      if (!isNaN(d.getTime())) {
        return d.toLocaleDateString();
      }
      return String(value);
    }

    case 'text':
      return String(value);

    case 'general':
    default: {
      if (isNumeric) {
        return Number.isInteger(num) ? num.toLocaleString() : num.toString();
      }
      return String(value);
    }
  }
}
