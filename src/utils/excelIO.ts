/**
 * @license
 * SPDX-License-Identifier: Apache-2.0
 */

import * as XLSX from 'xlsx';
import { SheetData, WorkbookData, CellData } from '../types/excel';
import { ChartConfig } from '../types/charts';
import { coordsToCellId } from './cellUtils';

/**
 * Directly exports workbook to user's local drive as an authentic .xlsx file
 */
export function exportToExcel(workbook: WorkbookData) {
  const wb = XLSX.utils.book_new();

  for (const sheet of workbook.sheets) {
    const ws: XLSX.WorkSheet = {};
    let maxR = 0;
    let maxC = 0;

    for (const [cellId, cell] of Object.entries(sheet.cells)) {
      if (!cell || (cell.raw === '' && cell.value === undefined)) continue;

      const cellRef = XLSX.utils.decode_cell(cellId);
      if (cellRef.r > maxR) maxR = cellRef.r;
      if (cellRef.c > maxC) maxC = cellRef.c;

      const wsCell: XLSX.CellObject = { v: '', t: 's' };

      if (cell.raw && cell.raw.startsWith('=')) {
        // Formula
        wsCell.f = cell.raw.substring(1);
        if (cell.value !== undefined) {
          wsCell.v = cell.value;
          wsCell.t = typeof cell.value === 'number' ? 'n' : typeof cell.value === 'boolean' ? 'b' : 's';
        } else {
          wsCell.v = 0;
          wsCell.t = 'n';
        }
      } else {
        // Raw value
        const num = Number(cell.raw);
        if (!isNaN(num) && cell.raw.trim() !== '') {
          wsCell.v = num;
          wsCell.t = 'n';
        } else if (cell.raw.toLowerCase() === 'true' || cell.raw.toLowerCase() === 'false') {
          wsCell.v = cell.raw.toLowerCase() === 'true';
          wsCell.t = 'b';
        } else {
          wsCell.v = cell.raw;
          wsCell.t = 's';
        }
      }

      ws[cellId] = wsCell;
    }

    // Set bounding range
    ws['!ref'] = XLSX.utils.encode_range({
      s: { r: 0, c: 0 },
      e: { r: Math.max(maxR, 20), c: Math.max(maxC, 10) }
    });

    // Column widths
    const cols: XLSX.ColInfo[] = [];
    for (let c = 0; c <= Math.max(maxC, 15); c++) {
      const customWidth = sheet.colWidths[c];
      cols.push({ wch: customWidth ? Math.round(customWidth / 8) : 14 });
    }
    ws['!cols'] = cols;

    XLSX.utils.book_append_sheet(wb, ws, sheet.name.substring(0, 31));
  }

  const fileName = workbook.fileName.endsWith('.xlsx') ? workbook.fileName : `${workbook.fileName}.xlsx`;
  XLSX.writeFile(wb, fileName);
}

/**
 * Export active sheet directly as .csv to local drive
 */
export function exportToCSV(sheet: SheetData, fileName: string) {
  const ws: XLSX.WorkSheet = {};
  let maxR = 0;
  let maxC = 0;

  for (const [cellId, cell] of Object.entries(sheet.cells)) {
    if (!cell || (cell.raw === '' && cell.value === undefined)) continue;
    const cellRef = XLSX.utils.decode_cell(cellId);
    if (cellRef.r > maxR) maxR = cellRef.r;
    if (cellRef.c > maxC) maxC = cellRef.c;

    const val = cell.value !== undefined ? cell.value : cell.raw;
    ws[cellId] = { v: val, t: typeof val === 'number' ? 'n' : 's' };
  }

  ws['!ref'] = XLSX.utils.encode_range({
    s: { r: 0, c: 0 },
    e: { r: maxR, c: maxC }
  });

  const csvOutput = XLSX.utils.sheet_to_csv(ws);
  const blob = new Blob([csvOutput], { type: 'text/csv;charset=utf-8;' });
  const url = URL.createObjectURL(blob);
  const link = document.createElement('a');
  link.href = url;
  const safeName = fileName.replace(/\.[^/.]+$/, '') + '.csv';
  link.setAttribute('download', safeName);
  document.body.appendChild(link);
  link.click();
  document.body.removeChild(link);
  URL.revokeObjectURL(url);
}

/**
 * Import an .xlsx or .csv file from user's local drive
 */
export async function importFromFile(file: File): Promise<WorkbookData> {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();

    reader.onload = (e) => {
      try {
        const data = new Uint8Array(e.target?.result as ArrayBuffer);
        const wb = XLSX.read(data, { type: 'array', cellFormula: true, cellStyles: true });

        const sheets: SheetData[] = [];

        wb.SheetNames.forEach((sheetName, index) => {
          const ws = wb.Sheets[sheetName];
          const cells: Record<string, CellData> = {};
          let maxRow = 50;
          let maxCol = 26;

          if (ws['!ref']) {
            const range = XLSX.utils.decode_range(ws['!ref']);
            maxRow = Math.max(50, range.e.r + 15);
            maxCol = Math.max(26, range.e.c + 5);
          }

          // Read all cells
          for (const cellAddress in ws) {
            if (cellAddress.startsWith('!')) continue;
            const cellObj = ws[cellAddress];
            if (!cellObj) continue;

            const hasFormula = Boolean(cellObj.f);
            const raw = hasFormula ? `=${cellObj.f}` : String(cellObj.v ?? '');
            const value = cellObj.v;

            cells[cellAddress.toUpperCase()] = {
              raw,
              value,
              format: {
                bold: cellObj.s?.font?.bold,
                italic: cellObj.s?.font?.italic,
                numberFormat: typeof cellObj.v === 'number' ? 'number' : 'general',
              }
            };
          }

          sheets.push({
            id: `sheet-${index + 1}-${Date.now()}`,
            name: sheetName,
            cells,
            rowCount: maxRow,
            colCount: maxCol,
            colWidths: {},
            rowHeights: {},
            freezeRows: 0,
            freezeCols: 0,
            conditionalRules: [],
            validationRules: {},
          });
        });

        resolve({
          fileName: file.name,
          activeSheetId: sheets[0]?.id || 'sheet-1',
          sheets: sheets.length > 0 ? sheets : [createEmptySheet('Sheet1', 'sheet-1')],
        });
      } catch (err) {
        reject(err);
      }
    };

    reader.onerror = (err) => reject(err);
    reader.readAsArrayBuffer(file);
  });
}

export function createEmptySheet(name: string, id: string): SheetData {
  return {
    id,
    name,
    cells: {},
    rowCount: 60,
    colCount: 26,
    colWidths: {},
    rowHeights: {},
    freezeRows: 0,
    freezeCols: 0,
    conditionalRules: [],
    validationRules: {},
  };
}

/**
 * Built-in Excel Templates for instantaneous exploration & testing
 */
export function getTemplateWorkbook(templateKey: 'executive_sales' | 'financial_model' | 'project_tracker' | 'inventory_mgmt' | 'blank'): { workbook: WorkbookData; charts: ChartConfig[] } {
  switch (templateKey) {
    case 'executive_sales':
      return createExecutiveSalesTemplate();
    case 'financial_model':
      return createFinancialModelTemplate();
    case 'project_tracker':
      return createProjectTrackerTemplate();
    case 'inventory_mgmt':
      return createInventoryTemplate();
    case 'blank':
    default:
      return {
        workbook: {
          fileName: 'Workbook1.xlsx',
          activeSheetId: 'sheet-1',
          sheets: [createEmptySheet('Sheet1', 'sheet-1')],
        },
        charts: []
      };
  }
}

function createExecutiveSalesTemplate(): { workbook: WorkbookData; charts: ChartConfig[] } {
  const sheetId = 'sheet-sales';
  const cells: Record<string, CellData> = {};

  // Title Block
  cells['A1'] = { raw: 'GLOBAL SALES & REVENUE DASHBOARD 2026', format: { bold: true, fontSize: 16, textColor: '#0f172a' } };
  cells['A2'] = { raw: 'Executive Performance Summary by Region & Quarter', format: { italic: true, textColor: '#64748b' } };

  // Headers (Row 4)
  const headers = ['Region', 'Q1 Revenue', 'Q2 Revenue', 'Q3 Revenue', 'Q4 (Target)', 'Annual Total', 'YoY Growth %', 'Target Status'];
  headers.forEach((h, idx) => {
    const cellId = coordsToCellId(3, idx);
    cells[cellId] = {
      raw: h,
      format: { bold: true, bgColor: '#107c41', textColor: '#ffffff', align: idx > 0 ? 'right' : 'left' }
    };
  });

  // Data Rows (Rows 5-9)
  const rows = [
    ['North America', 142000, 168000, 195000, 220000, '=SUM(B5:E5)', '=((E5-B5)/B5)', '=IF(F5>700000,"Exceeded","On Track")'],
    ['Europe (EMEA)', 118000, 129000, 142000, 165000, '=SUM(B6:E6)', '=((E6-B6)/B6)', '=IF(F6>500000,"Exceeded","On Track")'],
    ['Asia-Pacific', 96000, 115000, 138000, 172000, '=SUM(B7:E7)', '=((E7-B7)/B7)', '=IF(F7>500000,"Exceeded","On Track")'],
    ['Latin America', 45000, 52000, 61000, 74000, '=SUM(B8:E8)', '=((E8-B8)/B8)', '=IF(F8>200000,"Exceeded","On Track")'],
    ['Middle East & Africa', 38000, 44000, 53000, 68000, '=SUM(B9:E9)', '=((E9-B9)/B9)', '=IF(F9>200000,"Exceeded","On Track")'],
  ];

  rows.forEach((row, rIdx) => {
    const rowNum = 4 + rIdx; // 0-based index
    row.forEach((val, cIdx) => {
      const cellId = coordsToCellId(rowNum, cIdx);
      const isFormula = String(val).startsWith('=');
      const isCurrency = cIdx >= 1 && cIdx <= 5;
      const isPercent = cIdx === 6;

      cells[cellId] = {
        raw: String(val),
        format: {
          numberFormat: isCurrency ? 'currency' : isPercent ? 'percent' : 'general',
          decimals: isPercent ? 1 : 0,
          align: cIdx > 0 ? 'right' : 'left',
          bgColor: rIdx % 2 === 1 ? '#f8fafc' : '#ffffff',
        }
      };
    });
  });

  // Summary Row (Row 10)
  const summaryRowNum = 9; // 0-based
  cells['A10'] = { raw: 'TOTAL REVENUE', format: { bold: true, bgColor: '#e2e8f0', textColor: '#0f172a' } };
  cells['B10'] = { raw: '=SUM(B5:B9)', format: { bold: true, numberFormat: 'currency', decimals: 0, bgColor: '#e2e8f0', align: 'right' } };
  cells['C10'] = { raw: '=SUM(C5:C9)', format: { bold: true, numberFormat: 'currency', decimals: 0, bgColor: '#e2e8f0', align: 'right' } };
  cells['D10'] = { raw: '=SUM(D5:D9)', format: { bold: true, numberFormat: 'currency', decimals: 0, bgColor: '#e2e8f0', align: 'right' } };
  cells['E10'] = { raw: '=SUM(E5:E9)', format: { bold: true, numberFormat: 'currency', decimals: 0, bgColor: '#e2e8f0', align: 'right' } };
  cells['F10'] = { raw: '=SUM(F5:F9)', format: { bold: true, numberFormat: 'currency', decimals: 0, bgColor: '#e2e8f0', align: 'right' } };
  cells['G10'] = { raw: '=AVERAGE(G5:G9)', format: { bold: true, numberFormat: 'percent', decimals: 1, bgColor: '#e2e8f0', align: 'right' } };
  cells['H10'] = { raw: 'OVERPERFORMING', format: { bold: true, bgColor: '#dcfce7', textColor: '#15803d', align: 'center' } };

  // KPI Block (Row 12-14)
  cells['A12'] = { raw: 'KPI HIGHLIGHTS', format: { bold: true, fontSize: 13, textColor: '#107c41' } };
  cells['A13'] = { raw: 'Average Regional Total:', format: { bold: true } };
  cells['B13'] = { raw: '=AVERAGE(F5:F9)', format: { bold: true, numberFormat: 'currency', decimals: 0 } };
  cells['A14'] = { raw: 'Top Regional Performer:', format: { bold: true } };
  cells['B14'] = { raw: '=INDEX(A5:A9, MATCH(MAX(F5:F9), F5:F9, 0))', format: { bold: true, textColor: '#107c41' } };
  cells['C14'] = { raw: '=MAX(F5:F9)', format: { bold: true, numberFormat: 'currency', decimals: 0 } };

  const salesSheet: SheetData = {
    id: sheetId,
    name: 'Sales & Revenue',
    cells,
    rowCount: 50,
    colCount: 26,
    colWidths: { 0: 170, 1: 120, 2: 120, 3: 120, 4: 125, 5: 140, 6: 120, 7: 130 },
    rowHeights: {},
    freezeRows: 4,
    freezeCols: 1,
    tabColor: '#107c41',
    conditionalRules: [
      { id: 'cr-1', range: 'G5:G9', type: 'colorScale', color: 'green-red' },
      { id: 'cr-2', range: 'F5:F9', type: 'greaterThan', value1: 500000, color: '#ecfdf5', textColor: '#047857' },
    ],
    validationRules: {},
  };

  // Product categories sheet
  const prodSheetId = 'sheet-products';
  const prodCells: Record<string, CellData> = {};
  prodCells['A1'] = { raw: 'Product Category Sales 2026', format: { bold: true, fontSize: 14 } };
  ['Category', 'Units Sold', 'Unit Price', 'Gross Revenue', 'Profit Margin %'].forEach((h, idx) => {
    prodCells[coordsToCellId(2, idx)] = { raw: h, format: { bold: true, bgColor: '#1e293b', textColor: '#ffffff', align: idx > 0 ? 'right' : 'left' } };
  });

  const prodData = [
    ['Enterprise Cloud', 2450, 450, '=B4*C4', 0.68],
    ['AI Workstations', 1120, 1850, '=B5*C5', 0.42],
    ['Consulting Services', 840, 950, '=B6*C6', 0.55],
    ['Security Hardware', 3900, 210, '=B7*C7', 0.35],
    ['Developer Tools', 6200, 120, '=B8*C8', 0.74],
  ];

  prodData.forEach((row, r) => {
    row.forEach((v, c) => {
      const cellId = coordsToCellId(3 + r, c);
      prodCells[cellId] = {
        raw: String(v),
        format: {
          numberFormat: c === 2 || c === 3 ? 'currency' : c === 4 ? 'percent' : c === 1 ? 'number' : 'general',
          decimals: c === 4 ? 1 : 0,
          align: c > 0 ? 'right' : 'left',
        }
      };
    });
  });

  prodCells['A9'] = { raw: 'TOTAL', format: { bold: true, bgColor: '#e2e8f0' } };
  prodCells['B9'] = { raw: '=SUM(B4:B8)', format: { bold: true, numberFormat: 'number', decimals: 0, bgColor: '#e2e8f0', align: 'right' } };
  prodCells['D9'] = { raw: '=SUM(D4:D8)', format: { bold: true, numberFormat: 'currency', decimals: 0, bgColor: '#e2e8f0', align: 'right' } };

  const productsSheet: SheetData = {
    id: prodSheetId,
    name: 'Products & Margins',
    cells: prodCells,
    rowCount: 40,
    colCount: 20,
    colWidths: { 0: 180, 1: 110, 2: 110, 3: 130, 4: 130 },
    rowHeights: {},
    freezeRows: 3,
    freezeCols: 1,
    tabColor: '#0284c7',
    conditionalRules: [],
    validationRules: {},
  };

  // Pre-configured executive dashboard charts
  const charts: ChartConfig[] = [
    {
      id: 'chart-regional-perf',
      title: 'Quarterly Revenue Trends by Region',
      subtitle: 'Q1 through Q4 regional comparison',
      type: 'column',
      sheetId: salesSheet.id,
      dataRange: 'A4:E9',
      labelColumn: 'A',
      seriesColumns: ['B', 'C', 'D', 'E'],
      colors: ['#3b82f6', '#10b981', '#f59e0b', '#8b5cf6'],
      showLegend: true,
      showGrid: true,
      showValues: false,
      isDashboardWidget: true,
    },
    {
      id: 'chart-product-share',
      title: 'Gross Revenue by Product Category',
      subtitle: 'Category contribution to top line',
      type: 'donut',
      sheetId: productsSheet.id,
      dataRange: 'A3:D8',
      labelColumn: 'A',
      seriesColumns: ['D'],
      colors: ['#0284c7', '#10b981', '#6366f1', '#f59e0b', '#ec4899'],
      showLegend: true,
      showGrid: false,
      showValues: true,
      isDashboardWidget: true,
    },
    {
      id: 'chart-growth-trajectory',
      title: 'Annual Total Sales by Territory',
      subtitle: 'Cumulative performance for 2026',
      type: 'bar',
      sheetId: salesSheet.id,
      dataRange: 'A4:F9',
      labelColumn: 'A',
      seriesColumns: ['F'],
      colors: ['#107c41'],
      showLegend: false,
      showGrid: true,
      showValues: true,
      isDashboardWidget: true,
    },
    {
      id: 'kpi-total-rev',
      title: 'Total Global Revenue',
      type: 'kpi',
      sheetId: salesSheet.id,
      dataRange: 'F10',
      labelColumn: 'A',
      seriesColumns: ['F'],
      colors: ['#107c41'],
      showLegend: false,
      showGrid: false,
      showValues: true,
      isDashboardWidget: true,
      kpiDetails: {
        metricCell: 'F10',
        format: 'currency',
        periodLabel: 'Annual Forecast'
      }
    },
    {
      id: 'kpi-avg-growth',
      title: 'Average YoY Growth Rate',
      type: 'kpi',
      sheetId: salesSheet.id,
      dataRange: 'G10',
      labelColumn: 'A',
      seriesColumns: ['G'],
      colors: ['#3b82f6'],
      showLegend: false,
      showGrid: false,
      showValues: true,
      isDashboardWidget: true,
      kpiDetails: {
        metricCell: 'G10',
        format: 'percent',
        periodLabel: 'Across 5 Continents'
      }
    }
  ];

  return {
    workbook: {
      fileName: 'Executive_Sales_Dashboard_2026.xlsx',
      activeSheetId: salesSheet.id,
      sheets: [salesSheet, productsSheet],
    },
    charts,
  };
}

function createFinancialModelTemplate(): { workbook: WorkbookData; charts: ChartConfig[] } {
  const sheetId = 'sheet-pl';
  const cells: Record<string, CellData> = {};

  cells['A1'] = { raw: 'FIVE-YEAR FINANCIAL MODEL & P&L STATEMENT', format: { bold: true, fontSize: 16 } };
  cells['A2'] = { raw: 'In USD ($ Millions)', format: { italic: true, textColor: '#64748b' } };

  const headers = ['Financial Line Item', 'FY2022', 'FY2023', 'FY2024', 'FY2025', 'FY2026E', 'CAGR %'];
  headers.forEach((h, idx) => {
    cells[coordsToCellId(3, idx)] = { raw: h, format: { bold: true, bgColor: '#0f172a', textColor: '#ffffff', align: idx > 0 ? 'right' : 'left' } };
  });

  const pnlRows = [
    ['Gross Revenue', 42.5, 54.8, 71.2, 92.4, 118.6, '=((F5/B5)^(1/4))-1'],
    ['Cost of Goods Sold (COGS)', 15.3, 19.2, 24.1, 30.5, 38.0, '=((F6/B6)^(1/4))-1'],
    ['Gross Profit', '=B5-B6', '=C5-C6', '=D5-D6', '=E5-E6', '=F5-F6', '=((F7/B7)^(1/4))-1'],
    ['Gross Margin %', '=B7/B5', '=C7/C5', '=D7/D5', '=E7/E5', '=F7/F5', ''],
    ['Research & Development (R&D)', 8.2, 10.5, 13.8, 17.5, 22.0, ''],
    ['Sales & Marketing (S&M)', 11.0, 14.2, 18.0, 22.8, 28.5, ''],
    ['General & Admin (G&A)', 4.1, 5.0, 6.2, 7.5, 9.2, ''],
    ['Total Operating Expenses', '=SUM(B9:B11)', '=SUM(C9:C11)', '=SUM(D9:D11)', '=SUM(E9:E11)', '=SUM(F9:F11)', ''],
    ['Operating Income (EBIT)', '=B7-B12', '=C7-C12', '=D7-D12', '=E7-E12', '=F7-F12', ''],
    ['Operating Margin %', '=B13/B5', '=C13/C5', '=D13/D5', '=E13/E5', '=F13/F5', ''],
    ['Income Taxes (21%)', '=B13*0.21', '=C13*0.21', '=D13*0.21', '=E13*0.21', '=F13*0.21', ''],
    ['Net Income', '=B13-B15', '=C13-C15', '=D13-D15', '=E13-E15', '=F13-F15', '=((F16/B16)^(1/4))-1'],
  ];

  pnlRows.forEach((row, r) => {
    const rowNum = 4 + r;
    const isBold = [2, 3, 7, 8, 9, 11].includes(r);
    const isMargin = [3, 9].includes(r);

    row.forEach((v, c) => {
      const cellId = coordsToCellId(rowNum, c);
      const isPercent = isMargin || c === 6;
      cells[cellId] = {
        raw: String(v),
        format: {
          bold: isBold,
          numberFormat: isPercent ? 'percent' : c > 0 ? 'currency' : 'general',
          currencySymbol: '$',
          decimals: isPercent ? 1 : 1,
          align: c > 0 ? 'right' : 'left',
          bgColor: isBold ? '#f1f5f9' : '#ffffff',
        }
      };
    });
  });

  const plSheet: SheetData = {
    id: sheetId,
    name: 'P&L Model',
    cells,
    rowCount: 40,
    colCount: 20,
    colWidths: { 0: 220, 1: 100, 2: 100, 3: 100, 4: 100, 5: 110, 6: 100 },
    rowHeights: {},
    freezeRows: 4,
    freezeCols: 1,
    tabColor: '#3b82f6',
    conditionalRules: [
      { id: 'c-1', range: 'B16:F16', type: 'colorScale', color: 'green-red' }
    ],
    validationRules: {},
  };

  const charts: ChartConfig[] = [
    {
      id: 'chart-rev-growth',
      title: '5-Year Revenue & Net Income Trajectory ($M)',
      type: 'area',
      sheetId,
      dataRange: 'A4:F16',
      labelColumn: 'A',
      seriesColumns: ['B', 'C', 'D', 'E', 'F'],
      colors: ['#0284c7', '#10b981'],
      showLegend: true,
      showGrid: true,
      showValues: false,
      isDashboardWidget: true,
    },
    {
      id: 'kpi-net-income',
      title: 'FY2026 Projected Net Income',
      type: 'kpi',
      sheetId,
      dataRange: 'F16',
      labelColumn: 'A',
      seriesColumns: ['F'],
      colors: ['#107c41'],
      showLegend: false,
      showGrid: false,
      showValues: true,
      isDashboardWidget: true,
      kpiDetails: {
        metricCell: 'F16',
        format: 'currency',
        periodLabel: 'Target Year'
      }
    }
  ];

  return {
    workbook: {
      fileName: 'Corporate_Financial_Model.xlsx',
      activeSheetId: sheetId,
      sheets: [plSheet],
    },
    charts,
  };
}

function createProjectTrackerTemplate(): { workbook: WorkbookData; charts: ChartConfig[] } {
  const sheetId = 'sheet-projects';
  const cells: Record<string, CellData> = {};

  cells['A1'] = { raw: 'PROJECT MILESTONE & SPRINT TRACKER', format: { bold: true, fontSize: 16 } };
  const headers = ['Task ID', 'Work Item / Deliverable', 'Owner', 'Priority', 'Status', 'Start Date', 'Days', 'Budget', 'Spent'];
  headers.forEach((h, idx) => {
    cells[coordsToCellId(2, idx)] = { raw: h, format: { bold: true, bgColor: '#334155', textColor: '#ffffff', align: idx >= 6 ? 'right' : 'left' } };
  });

  const tasks = [
    ['TSK-101', 'Cloud Infrastructure Migration', 'David Chen', 'High', 'Completed', '2026-09-01', 30, 25000, 24200],
    ['TSK-102', 'API Gateway Rate Limiter', 'Sarah Jenkins', 'Medium', 'Completed', '2026-09-15', 14, 12000, 11500],
    ['TSK-103', 'Multi-tenant Auth & RBAC', 'Marcus Vance', 'Critical', 'In Progress', '2026-10-01', 25, 35000, 28000],
    ['TSK-104', 'Data Visualization Engine', 'Elena Rostova', 'High', 'In Progress', '2026-10-10', 21, 28000, 19500],
    ['TSK-105', 'Automated QA Regression Suite', 'Priya Patel', 'Medium', 'Not Started', '2026-11-01', 18, 15000, 0],
    ['TSK-106', 'Executive Security Audit', 'External Partner', 'Critical', 'Not Started', '2026-11-15', 10, 40000, 0],
  ];

  tasks.forEach((t, r) => {
    t.forEach((v, c) => {
      cells[coordsToCellId(3 + r, c)] = {
        raw: String(v),
        format: {
          numberFormat: c >= 7 ? 'currency' : c === 6 ? 'number' : 'general',
          decimals: 0,
          align: c >= 6 ? 'right' : 'left',
          bgColor: r % 2 === 1 ? '#f8fafc' : '#ffffff',
        }
      };
    });
  });

  cells['A10'] = { raw: 'TOTALS', format: { bold: true, bgColor: '#e2e8f0' } };
  cells['H10'] = { raw: '=SUM(H4:H9)', format: { bold: true, numberFormat: 'currency', decimals: 0, bgColor: '#e2e8f0', align: 'right' } };
  cells['I10'] = { raw: '=SUM(I4:I9)', format: { bold: true, numberFormat: 'currency', decimals: 0, bgColor: '#e2e8f0', align: 'right' } };

  const projectSheet: SheetData = {
    id: sheetId,
    name: 'Sprints & Tasks',
    cells,
    rowCount: 40,
    colCount: 20,
    colWidths: { 0: 90, 1: 240, 2: 130, 3: 90, 4: 110, 5: 110, 6: 70, 7: 100, 8: 100 },
    rowHeights: {},
    freezeRows: 3,
    freezeCols: 2,
    tabColor: '#8b5cf6',
    conditionalRules: [
      { id: 'pr-1', range: 'E4:E9', type: 'equalTo', value1: 'Completed', color: '#dcfce7', textColor: '#15803d' },
      { id: 'pr-2', range: 'E4:E9', type: 'equalTo', value1: 'In Progress', color: '#fef3c7', textColor: '#b45309' },
    ],
    validationRules: {},
  };

  const charts: ChartConfig[] = [
    {
      id: 'chart-project-budget',
      title: 'Budget vs Actual Spent by Task',
      type: 'column',
      sheetId,
      dataRange: 'B3:I9',
      labelColumn: 'B',
      seriesColumns: ['H', 'I'],
      colors: ['#3b82f6', '#10b981'],
      showLegend: true,
      showGrid: true,
      showValues: false,
      isDashboardWidget: true,
    }
  ];

  return {
    workbook: {
      fileName: 'Project_Sprint_Tracker.xlsx',
      activeSheetId: sheetId,
      sheets: [projectSheet],
    },
    charts,
  };
}

function createInventoryTemplate(): { workbook: WorkbookData; charts: ChartConfig[] } {
  const sheetId = 'sheet-inventory';
  const cells: Record<string, CellData> = {};

  cells['A1'] = { raw: 'INVENTORY VALUATION & STOCK TRACKER', format: { bold: true, fontSize: 16 } };
  const headers = ['SKU', 'Product Item Name', 'Category', 'In Stock', 'Unit Cost', 'Total Valuation', 'Status'];
  headers.forEach((h, idx) => {
    cells[coordsToCellId(2, idx)] = { raw: h, format: { bold: true, bgColor: '#065f46', textColor: '#ffffff', align: idx >= 3 && idx <= 5 ? 'right' : 'left' } };
  });

  const items = [
    ['INV-001', 'Ultra-Wide 4K Monitor', 'Displays', 140, 380, '=D4*E4', '=IF(D4<50,"Low Stock","Optimal")'],
    ['INV-002', 'Mechanical RGB Keyboard', 'Peripherals', 320, 65, '=D5*E5', '=IF(D5<50,"Low Stock","Optimal")'],
    ['INV-003', 'Wireless Ergonomic Mouse', 'Peripherals', 450, 42, '=D6*E6', '=IF(D6<50,"Low Stock","Optimal")'],
    ['INV-004', 'Thunderbolt 4 Docking Station', 'Accessories', 35, 180, '=D7*E7', '=IF(D7<50,"Low Stock","Optimal")'],
    ['INV-005', 'Noise-Cancelling Studio Headset', 'Audio', 210, 145, '=D8*E8', '=IF(D8<50,"Low Stock","Optimal")'],
  ];

  items.forEach((item, r) => {
    item.forEach((v, c) => {
      cells[coordsToCellId(3 + r, c)] = {
        raw: String(v),
        format: {
          numberFormat: c === 4 || c === 5 ? 'currency' : c === 3 ? 'number' : 'general',
          decimals: c === 4 || c === 5 ? 0 : 0,
          align: c >= 3 && c <= 5 ? 'right' : 'left',
          bgColor: r % 2 === 1 ? '#f8fafc' : '#ffffff',
        }
      };
    });
  });

  cells['A9'] = { raw: 'TOTAL WAREHOUSE VALUE', format: { bold: true, bgColor: '#e2e8f0' } };
  cells['D9'] = { raw: '=SUM(D4:D8)', format: { bold: true, numberFormat: 'number', decimals: 0, bgColor: '#e2e8f0', align: 'right' } };
  cells['F9'] = { raw: '=SUM(F4:F8)', format: { bold: true, numberFormat: 'currency', decimals: 0, bgColor: '#e2e8f0', align: 'right' } };

  const invSheet: SheetData = {
    id: sheetId,
    name: 'Stock Inventory',
    cells,
    rowCount: 40,
    colCount: 20,
    colWidths: { 0: 100, 1: 260, 2: 130, 3: 90, 4: 100, 5: 140, 6: 110 },
    rowHeights: {},
    freezeRows: 3,
    freezeCols: 2,
    tabColor: '#059669',
    conditionalRules: [
      { id: 'ir-1', range: 'G4:G8', type: 'equalTo', value1: 'Low Stock', color: '#fee2e2', textColor: '#b91c1c' }
    ],
    validationRules: {},
  };

  const charts: ChartConfig[] = [
    {
      id: 'chart-inv-value',
      title: 'Inventory Value Distribution by Item',
      type: 'donut',
      sheetId,
      dataRange: 'B3:F8',
      labelColumn: 'B',
      seriesColumns: ['F'],
      colors: ['#059669', '#10b981', '#34d399', '#6ee7b7', '#a7f3d0'],
      showLegend: true,
      showGrid: false,
      showValues: true,
      isDashboardWidget: true,
    }
  ];

  return {
    workbook: {
      fileName: 'Inventory_Valuation_Model.xlsx',
      activeSheetId: sheetId,
      sheets: [invSheet],
    },
    charts,
  };
}
