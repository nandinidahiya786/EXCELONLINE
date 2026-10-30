/**
 * @license
 * SPDX-License-Identifier: Apache-2.0
 */

import React, { useState, useEffect, useRef, useCallback } from 'react';
import { SheetData, WorkbookData, CellCoordinate, CellRange, CellFormat, CellData, DataValidationRule } from './types/excel';
import { ChartConfig } from './types/charts';
import { coordsToCellId, formatRange, parseCellId, shiftFormula, getCellsInRange } from './utils/cellUtils';
import { FormulaEngine } from './utils/formulaEngine';
import { exportToExcel, exportToCSV, importFromFile, getTemplateWorkbook, createEmptySheet } from './utils/excelIO';
import { Header } from './components/Header';
import { Ribbon } from './components/Ribbon';
import { FormulaBar } from './components/FormulaBar';
import { SpreadsheetGrid } from './components/SpreadsheetGrid';
import { SheetTabs } from './components/SheetTabs';
import { StatusBar } from './components/StatusBar';
import { DashboardView } from './components/DashboardView';
import { InsertFunctionModal } from './components/InsertFunctionModal';
import { CreateChartModal } from './components/CreateChartModal';
import { FindReplaceModal } from './components/FindReplaceModal';
import { FormatCellsModal } from './components/FormatCellsModal';
import { DataValidationModal } from './components/DataValidationModal';

const STORAGE_KEY = 'excelpro_studio_workbook_v1';
const CHARTS_STORAGE_KEY = 'excelpro_studio_charts_v1';

export default function App() {
  // Initialize from LocalStorage or Executive Sales Template
  const [workbook, setWorkbook] = useState<WorkbookData>(() => {
    try {
      const saved = localStorage.getItem(STORAGE_KEY);
      if (saved) {
        return JSON.parse(saved);
      }
    } catch {
      // Ignore parse error and load default
    }
    const initial = getTemplateWorkbook('executive_sales');
    return initial.workbook;
  });

  const [charts, setCharts] = useState<ChartConfig[]>(() => {
    try {
      const saved = localStorage.getItem(CHARTS_STORAGE_KEY);
      if (saved) {
        return JSON.parse(saved);
      }
    } catch {
      // Ignore
    }
    const initial = getTemplateWorkbook('executive_sales');
    return initial.charts;
  });

  // Active View: Sheet Grid or Dashboard Canvas
  const [activeView, setActiveView] = useState<'grid' | 'dashboard'>('grid');
  const [activeRibbonTab, setActiveRibbonTab] = useState<'home' | 'insert' | 'formulas' | 'data' | 'dashboard' | 'view'>('home');

  // Navigation & Selection State
  const [activeCell, setActiveCell] = useState<CellCoordinate>({ row: 4, col: 1 }); // B5 by default
  const [selectedRange, setSelectedRange] = useState<CellRange>({ startRow: 4, startCol: 1, endRow: 4, endCol: 1 });

  // Grid View Options
  const [showGridlines, setShowGridlines] = useState(true);
  const [showFormulaBar, setShowFormulaBar] = useState(true);
  const [zoom, setZoom] = useState(1);

  // Undo / Redo stacks
  const [undoStack, setUndoStack] = useState<WorkbookData[]>([]);
  const [redoStack, setRedoStack] = useState<WorkbookData[]>([]);

  // Clipboard
  const [clipboard, setClipboard] = useState<{
    type: 'cut' | 'copy';
    range: CellRange;
    sheetId: string;
    cells: Record<string, CellData>;
  } | null>(null);

  // Modals
  const [isInsertFunctionOpen, setIsInsertFunctionOpen] = useState(false);
  const [isCreateChartOpen, setIsCreateChartOpen] = useState(false);
  const [editingChart, setEditingChart] = useState<ChartConfig | null>(null);
  const [isFindReplaceOpen, setIsFindReplaceOpen] = useState(false);
  const [isFormatCellsOpen, setIsFormatCellsOpen] = useState(false);
  const [isDataValidationOpen, setIsDataValidationOpen] = useState(false);

  // Active Sheet helper
  const activeSheet = workbook.sheets.find(s => s.id === workbook.activeSheetId) || workbook.sheets[0];
  const activeCellId = coordsToCellId(activeCell.row, activeCell.col);
  const activeCellData = activeSheet.cells?.[activeCellId];

  // Formula Engine Ref
  const formulaEngineRef = useRef<FormulaEngine | null>(null);

  // Run full calculation on sheet/workbook
  const recalculateWorkbook = useCallback((wb: WorkbookData): WorkbookData => {
    const engine = new FormulaEngine(wb.sheets, wb.activeSheetId);
    engine.evaluateWorkbook();
    formulaEngineRef.current = engine;
    return { ...wb };
  }, []);

  // Initial calculation on mount
  useEffect(() => {
    setWorkbook((prev) => recalculateWorkbook(prev));
  }, [recalculateWorkbook]);

  // Persist to LocalStorage on update
  useEffect(() => {
    try {
      localStorage.setItem(STORAGE_KEY, JSON.stringify(workbook));
      localStorage.setItem(CHARTS_STORAGE_KEY, JSON.stringify(charts));
    } catch {
      // Storage full or unavailable
    }
  }, [workbook, charts]);

  // Push history snapshot before mutating workbook
  const pushUndo = useCallback((wb: WorkbookData) => {
    setUndoStack((prev) => [...prev.slice(-30), JSON.parse(JSON.stringify(wb))]);
    setRedoStack([]);
  }, []);

  const handleUndo = useCallback(() => {
    if (undoStack.length === 0) return;
    const previous = undoStack[undoStack.length - 1];
    setUndoStack((prev) => prev.slice(0, -1));
    setRedoStack((prev) => [...prev, JSON.parse(JSON.stringify(workbook))]);
    setWorkbook(recalculateWorkbook(previous));
  }, [undoStack, workbook, recalculateWorkbook]);

  const handleRedo = useCallback(() => {
    if (redoStack.length === 0) return;
    const next = redoStack[redoStack.length - 1];
    setRedoStack((prev) => prev.slice(0, -1));
    setUndoStack((prev) => [...prev, JSON.parse(JSON.stringify(workbook))]);
    setWorkbook(recalculateWorkbook(next));
  }, [redoStack, workbook, recalculateWorkbook]);

  // Handle Cell Value Change (from in-cell editing or formula bar)
  const handleCellChange = useCallback((cellId: string, rawVal: string) => {
    pushUndo(workbook);
    setWorkbook((prevWb) => {
      const newSheets = prevWb.sheets.map((sheet) => {
        if (sheet.id !== prevWb.activeSheetId) return sheet;
        const newCells = { ...sheet.cells };
        const existing = newCells[cellId] || { raw: '' };
        newCells[cellId] = {
          ...existing,
          raw: rawVal,
        };
        return { ...sheet, cells: newCells };
      });

      const updated = { ...prevWb, sheets: newSheets };
      return recalculateWorkbook(updated);
    });
  }, [workbook, pushUndo, recalculateWorkbook]);

  // Apply Cell Format
  const handleApplyFormat = useCallback((formatPatch: Partial<CellFormat>) => {
    pushUndo(workbook);
    setWorkbook((prevWb) => {
      const newSheets = prevWb.sheets.map((sheet) => {
        if (sheet.id !== prevWb.activeSheetId) return sheet;
        const newCells = { ...sheet.cells };

        // Apply to all cells in selected range
        for (let r = selectedRange.startRow; r <= selectedRange.endRow; r++) {
          for (let c = selectedRange.startCol; c <= selectedRange.endCol; c++) {
            const cId = coordsToCellId(r, c);
            const existing = newCells[cId] || { raw: '' };
            const existingFormat = existing.format || {};
            newCells[cId] = {
              ...existing,
              format: { ...existingFormat, ...formatPatch },
            };
          }
        }
        return { ...sheet, cells: newCells };
      });

      return { ...prevWb, sheets: newSheets };
    });
  }, [workbook, selectedRange, pushUndo]);

  // Autofill / Drag Series & Formulas
  const handleAutofill = useCallback((sourceRange: CellRange, targetRange: CellRange) => {
    pushUndo(workbook);
    setWorkbook((prevWb) => {
      const newSheets = prevWb.sheets.map((sheet) => {
        if (sheet.id !== prevWb.activeSheetId) return sheet;
        const newCells = { ...sheet.cells };

        const deltaRow = targetRange.endRow - sourceRange.endRow;
        const deltaCol = targetRange.endCol - sourceRange.endCol;

        // Extending vertically downwards
        if (deltaRow > 0) {
          const sourceHeight = sourceRange.endRow - sourceRange.startRow + 1;
          for (let r = sourceRange.endRow + 1; r <= targetRange.endRow; r++) {
            const srcR = sourceRange.startRow + ((r - sourceRange.endRow - 1) % sourceHeight);
            const rowOffset = r - srcR;

            for (let c = sourceRange.startCol; c <= sourceRange.endCol; c++) {
              const srcId = coordsToCellId(srcR, c);
              const destId = coordsToCellId(r, c);
              const srcCell = newCells[srcId];

              if (!srcCell) continue;

              let newRaw = srcCell.raw;
              if (srcCell.raw.startsWith('=')) {
                newRaw = shiftFormula(srcCell.raw, rowOffset, 0);
              } else {
                const num = Number(srcCell.raw);
                if (!isNaN(num) && srcCell.raw.trim() !== '') {
                  newRaw = String(num + ((r - srcR) * 1));
                }
              }

              newCells[destId] = {
                ...srcCell,
                raw: newRaw,
                format: srcCell.format ? { ...srcCell.format } : undefined,
              };
            }
          }
        }

        // Extending horizontally to the right
        if (deltaCol > 0) {
          const sourceWidth = sourceRange.endCol - sourceRange.startCol + 1;
          for (let c = sourceRange.endCol + 1; c <= targetRange.endCol; c++) {
            const srcC = sourceRange.startCol + ((c - sourceRange.endCol - 1) % sourceWidth);
            const colOffset = c - srcC;

            for (let r = sourceRange.startRow; r <= sourceRange.endRow; r++) {
              const srcId = coordsToCellId(r, srcC);
              const destId = coordsToCellId(r, c);
              const srcCell = newCells[srcId];

              if (!srcCell) continue;

              let newRaw = srcCell.raw;
              if (srcCell.raw.startsWith('=')) {
                newRaw = shiftFormula(srcCell.raw, 0, colOffset);
              }

              newCells[destId] = {
                ...srcCell,
                raw: newRaw,
                format: srcCell.format ? { ...srcCell.format } : undefined,
              };
            }
          }
        }

        return { ...sheet, cells: newCells };
      });

      return recalculateWorkbook({ ...prevWb, sheets: newSheets });
    });
  }, [workbook, pushUndo, recalculateWorkbook]);

  // AutoSum function insertion
  const handleAutoSum = useCallback((funcName: 'SUM' | 'AVERAGE' | 'COUNT' | 'MAX' | 'MIN') => {
    // If range is selected above current cell, auto-suggest range
    let targetRangeStr = '';
    if (activeCell.row > 0) {
      const topCell = coordsToCellId(0, activeCell.col);
      const aboveCell = coordsToCellId(activeCell.row - 1, activeCell.col);
      targetRangeStr = `${topCell}:${aboveCell}`;
    } else {
      targetRangeStr = 'A1:A5';
    }

    const formula = `=${funcName}(${targetRangeStr})`;
    handleCellChange(activeCellId, formula);
  }, [activeCell, activeCellId, handleCellChange]);

  // Sort Range
  const handleSort = useCallback((direction: 'asc' | 'desc') => {
    pushUndo(workbook);
    setWorkbook((prevWb) => {
      const sheet = prevWb.sheets.find(s => s.id === prevWb.activeSheetId);
      if (!sheet) return prevWb;

      const newCells = { ...sheet.cells };
      const col = activeCell.col;
      const startR = selectedRange.startRow;
      const endR = selectedRange.endRow;

      if (endR <= startR) return prevWb;

      // Extract rows
      const rows: { r: number; val: any; cells: Record<number, CellData> }[] = [];
      for (let r = startR; r <= endR; r++) {
        const keyCell = newCells[coordsToCellId(r, col)];
        const rowCells: Record<number, CellData> = {};
        for (let c = 0; c < sheet.colCount; c++) {
          const cId = coordsToCellId(r, c);
          if (newCells[cId]) rowCells[c] = newCells[cId];
        }
        rows.push({
          r,
          val: keyCell?.value !== undefined ? keyCell.value : keyCell?.raw ?? '',
          cells: rowCells,
        });
      }

      // Sort
      rows.sort((a, b) => {
        const numA = Number(a.val);
        const numB = Number(b.val);
        if (!isNaN(numA) && !isNaN(numB)) {
          return direction === 'asc' ? numA - numB : numB - numA;
        }
        const strA = String(a.val).toLowerCase();
        const strB = String(b.val).toLowerCase();
        return direction === 'asc' ? strA.localeCompare(strB) : strB.localeCompare(strA);
      });

      // Write back sorted rows
      rows.forEach((rowObj, idx) => {
        const targetR = startR + idx;
        for (let c = 0; c < sheet.colCount; c++) {
          const destId = coordsToCellId(targetR, c);
          if (rowObj.cells[c]) {
            newCells[destId] = rowObj.cells[c];
          } else {
            delete newCells[destId];
          }
        }
      });

      const newSheets = prevWb.sheets.map(s => s.id === sheet.id ? { ...s, cells: newCells } : s);
      return recalculateWorkbook({ ...prevWb, sheets: newSheets });
    });
  }, [workbook, activeCell, selectedRange, pushUndo, recalculateWorkbook]);

  // Insert / Delete Rows & Columns
  const handleInsertRow = useCallback(() => {
    pushUndo(workbook);
    setWorkbook((prevWb) => {
      const newSheets = prevWb.sheets.map((sheet) => {
        if (sheet.id !== prevWb.activeSheetId) return sheet;
        const newCells: Record<string, CellData> = {};
        const insertR = activeCell.row;

        for (const [cellId, cell] of Object.entries(sheet.cells)) {
          const coord = parseCellId(cellId);
          if (!coord) continue;
          if (coord.row < insertR) {
            newCells[cellId] = cell;
          } else {
            const shiftedId = coordsToCellId(coord.row + 1, coord.col);
            newCells[shiftedId] = cell;
          }
        }
        return { ...sheet, rowCount: sheet.rowCount + 1, cells: newCells };
      });
      return recalculateWorkbook({ ...prevWb, sheets: newSheets });
    });
  }, [workbook, activeCell, pushUndo, recalculateWorkbook]);

  const handleInsertCol = useCallback(() => {
    pushUndo(workbook);
    setWorkbook((prevWb) => {
      const newSheets = prevWb.sheets.map((sheet) => {
        if (sheet.id !== prevWb.activeSheetId) return sheet;
        const newCells: Record<string, CellData> = {};
        const insertC = activeCell.col;

        for (const [cellId, cell] of Object.entries(sheet.cells)) {
          const coord = parseCellId(cellId);
          if (!coord) continue;
          if (coord.col < insertC) {
            newCells[cellId] = cell;
          } else {
            const shiftedId = coordsToCellId(coord.row, coord.col + 1);
            newCells[shiftedId] = cell;
          }
        }
        return { ...sheet, colCount: sheet.colCount + 1, cells: newCells };
      });
      return recalculateWorkbook({ ...prevWb, sheets: newSheets });
    });
  }, [workbook, activeCell, pushUndo, recalculateWorkbook]);

  const handleDeleteRow = useCallback(() => {
    pushUndo(workbook);
    setWorkbook((prevWb) => {
      const newSheets = prevWb.sheets.map((sheet) => {
        if (sheet.id !== prevWb.activeSheetId) return sheet;
        const newCells: Record<string, CellData> = {};
        const delR = activeCell.row;

        for (const [cellId, cell] of Object.entries(sheet.cells)) {
          const coord = parseCellId(cellId);
          if (!coord) continue;
          if (coord.row < delR) {
            newCells[cellId] = cell;
          } else if (coord.row > delR) {
            const shiftedId = coordsToCellId(coord.row - 1, coord.col);
            newCells[shiftedId] = cell;
          }
        }
        return { ...sheet, rowCount: Math.max(10, sheet.rowCount - 1), cells: newCells };
      });
      return recalculateWorkbook({ ...prevWb, sheets: newSheets });
    });
  }, [workbook, activeCell, pushUndo, recalculateWorkbook]);

  const handleDeleteCol = useCallback(() => {
    pushUndo(workbook);
    setWorkbook((prevWb) => {
      const newSheets = prevWb.sheets.map((sheet) => {
        if (sheet.id !== prevWb.activeSheetId) return sheet;
        const newCells: Record<string, CellData> = {};
        const delC = activeCell.col;

        for (const [cellId, cell] of Object.entries(sheet.cells)) {
          const coord = parseCellId(cellId);
          if (!coord) continue;
          if (coord.col < delC) {
            newCells[cellId] = cell;
          } else if (coord.col > delC) {
            const shiftedId = coordsToCellId(coord.row, coord.col - 1);
            newCells[shiftedId] = cell;
          }
        }
        return { ...sheet, colCount: Math.max(5, sheet.colCount - 1), cells: newCells };
      });
      return recalculateWorkbook({ ...prevWb, sheets: newSheets });
    });
  }, [workbook, activeCell, pushUndo, recalculateWorkbook]);

  // Clear Contents
  const handleClearContent = useCallback((type: 'all' | 'content' | 'format' = 'content') => {
    pushUndo(workbook);
    setWorkbook((prevWb) => {
      const newSheets = prevWb.sheets.map((sheet) => {
        if (sheet.id !== prevWb.activeSheetId) return sheet;
        const newCells = { ...sheet.cells };

        for (let r = selectedRange.startRow; r <= selectedRange.endRow; r++) {
          for (let c = selectedRange.startCol; c <= selectedRange.endCol; c++) {
            const cId = coordsToCellId(r, c);
            if (type === 'all') {
              delete newCells[cId];
            } else if (type === 'content') {
              if (newCells[cId]) {
                newCells[cId] = { ...newCells[cId], raw: '', value: undefined };
              }
            } else if (type === 'format') {
              if (newCells[cId]) {
                newCells[cId] = { ...newCells[cId], format: undefined };
              }
            }
          }
        }
        return { ...sheet, cells: newCells };
      });
      return recalculateWorkbook({ ...prevWb, sheets: newSheets });
    });
  }, [workbook, selectedRange, pushUndo, recalculateWorkbook]);

  // Cut, Copy, Paste
  const handleCopy = useCallback(() => {
    const copiedCells: Record<string, CellData> = {};
    for (let r = selectedRange.startRow; r <= selectedRange.endRow; r++) {
      for (let c = selectedRange.startCol; c <= selectedRange.endCol; c++) {
        const cId = coordsToCellId(r, c);
        if (activeSheet.cells?.[cId]) {
          copiedCells[cId] = JSON.parse(JSON.stringify(activeSheet.cells[cId]));
        }
      }
    }
    setClipboard({
      type: 'copy',
      range: selectedRange,
      sheetId: activeSheet.id,
      cells: copiedCells,
    });
  }, [selectedRange, activeSheet]);

  const handleCut = useCallback(() => {
    handleCopy();
    if (clipboard) setClipboard({ ...clipboard, type: 'cut' });
  }, [handleCopy, clipboard]);

  const handlePaste = useCallback(() => {
    if (!clipboard) return;
    pushUndo(workbook);

    setWorkbook((prevWb) => {
      const newSheets = prevWb.sheets.map((sheet) => {
        if (sheet.id !== prevWb.activeSheetId) return sheet;
        const newCells = { ...sheet.cells };

        const deltaRow = activeCell.row - clipboard.range.startRow;
        const deltaCol = activeCell.col - clipboard.range.startCol;

        for (const [srcId, cell] of Object.entries(clipboard.cells)) {
          const coord = parseCellId(srcId);
          if (!coord) continue;

          const destId = coordsToCellId(coord.row + deltaRow, coord.col + deltaCol);
          let raw = cell.raw;
          if (raw.startsWith('=')) {
            raw = shiftFormula(raw, deltaRow, deltaCol);
          }

          newCells[destId] = {
            ...cell,
            raw,
            format: cell.format ? { ...cell.format } : undefined,
          };
        }

        // If cut, remove original cells
        if (clipboard.type === 'cut') {
          for (const srcId of Object.keys(clipboard.cells)) {
            delete newCells[srcId];
          }
          setClipboard(null);
        }

        return { ...sheet, cells: newCells };
      });

      return recalculateWorkbook({ ...prevWb, sheets: newSheets });
    });
  }, [clipboard, activeCell, workbook, pushUndo, recalculateWorkbook]);

  // Conditional Formatting
  const handleApplyConditionalScale = useCallback(() => {
    const rangeStr = formatRange(selectedRange);
    setWorkbook((prevWb) => {
      const newSheets = prevWb.sheets.map((sheet) => {
        if (sheet.id !== prevWb.activeSheetId) return sheet;
        const rules = sheet.conditionalRules || [];
        return {
          ...sheet,
          conditionalRules: [
            ...rules,
            { id: `cr-${Date.now()}`, range: rangeStr, type: 'colorScale' as const, color: 'green-red' },
          ],
        };
      });
      return { ...prevWb, sheets: newSheets };
    });
  }, [selectedRange]);

  const handleApplyConditionalRule = useCallback((type: 'greaterThan' | 'lessThan' | 'equalTo', val: number) => {
    const rangeStr = formatRange(selectedRange);
    setWorkbook((prevWb) => {
      const newSheets = prevWb.sheets.map((sheet) => {
        if (sheet.id !== prevWb.activeSheetId) return sheet;
        const rules = sheet.conditionalRules || [];
        return {
          ...sheet,
          conditionalRules: [
            ...rules,
            {
              id: `cr-${Date.now()}`,
              range: rangeStr,
              type,
              value1: val,
              color: type === 'greaterThan' ? '#dcfce7' : '#fee2e2',
              textColor: type === 'greaterThan' ? '#15803d' : '#991b1b',
            },
          ],
        };
      });
      return { ...prevWb, sheets: newSheets };
    });
  }, [selectedRange]);

  // Freeze Panes
  const handleFreezeTopRow = useCallback(() => {
    setWorkbook((prev) => ({
      ...prev,
      sheets: prev.sheets.map(s => s.id === prev.activeSheetId ? { ...s, freezeRows: 1 } : s)
    }));
  }, []);

  const handleFreezeFirstCol = useCallback(() => {
    setWorkbook((prev) => ({
      ...prev,
      sheets: prev.sheets.map(s => s.id === prev.activeSheetId ? { ...s, freezeCols: 1 } : s)
    }));
  }, []);

  const handleUnfreeze = useCallback(() => {
    setWorkbook((prev) => ({
      ...prev,
      sheets: prev.sheets.map(s => s.id === prev.activeSheetId ? { ...s, freezeRows: 0, freezeCols: 0 } : s)
    }));
  }, []);

  // Sheet Tabs Management
  const handleAddSheet = useCallback(() => {
    pushUndo(workbook);
    const newSheetIndex = workbook.sheets.length + 1;
    const newId = `sheet-${Date.now()}`;
    const newSheet = createEmptySheet(`Sheet${newSheetIndex}`, newId);

    setWorkbook((prev) => ({
      ...prev,
      activeSheetId: newId,
      sheets: [...prev.sheets, newSheet],
    }));
  }, [workbook, pushUndo]);

  const handleRenameSheet = useCallback((sheetId: string, newName: string) => {
    setWorkbook((prev) => ({
      ...prev,
      sheets: prev.sheets.map(s => s.id === sheetId ? { ...s, name: newName } : s)
    }));
  }, []);

  const handleDeleteSheet = useCallback((sheetId: string) => {
    pushUndo(workbook);
    setWorkbook((prev) => {
      const remaining = prev.sheets.filter(s => s.id !== sheetId);
      const nextActiveId = prev.activeSheetId === sheetId ? (remaining[0]?.id || 'sheet-1') : prev.activeSheetId;
      return {
        ...prev,
        activeSheetId: nextActiveId,
        sheets: remaining,
      };
    });
  }, [workbook, pushUndo]);

  const handleDuplicateSheet = useCallback((sheetId: string) => {
    pushUndo(workbook);
    setWorkbook((prev) => {
      const source = prev.sheets.find(s => s.id === sheetId);
      if (!source) return prev;
      const clone: SheetData = JSON.parse(JSON.stringify(source));
      clone.id = `sheet-${Date.now()}`;
      clone.name = `${source.name} (Copy)`;
      return {
        ...prev,
        sheets: [...prev.sheets, clone],
      };
    });
  }, [workbook, pushUndo]);

  const handleChangeTabColor = useCallback((sheetId: string, color: string) => {
    setWorkbook((prev) => ({
      ...prev,
      sheets: prev.sheets.map(s => s.id === sheetId ? { ...s, tabColor: color } : s)
    }));
  }, []);

  // Column / Row Resize
  const handleResizeCol = useCallback((colIdx: number, width: number) => {
    setWorkbook((prev) => ({
      ...prev,
      sheets: prev.sheets.map(s => s.id === prev.activeSheetId ? {
        ...s,
        colWidths: { ...s.colWidths, [colIdx]: width }
      } : s)
    }));
  }, []);

  const handleResizeRow = useCallback((rowIdx: number, height: number) => {
    setWorkbook((prev) => ({
      ...prev,
      sheets: prev.sheets.map(s => s.id === prev.activeSheetId ? {
        ...s,
        rowHeights: { ...s.rowHeights, [rowIdx]: height }
      } : s)
    }));
  }, []);

  // Jump to cell
  const handleJumpToCell = useCallback((cellId: string) => {
    const parsed = parseCellId(cellId);
    if (parsed) {
      setActiveCell({ row: parsed.row, col: parsed.col });
      setSelectedRange({ startRow: parsed.row, startCol: parsed.col, endRow: parsed.row, endCol: parsed.col });
    }
  }, []);

  // Direct Excel Export to Local Drive
  const handleExportExcel = useCallback(() => {
    exportToExcel(workbook);
  }, [workbook]);

  // Export CSV
  const handleExportCSV = useCallback(() => {
    exportToCSV(activeSheet, workbook.fileName);
  }, [activeSheet, workbook.fileName]);

  // Import File (.xlsx / .csv)
  const handleImportFile = useCallback(async (file: File) => {
    try {
      const importedWb = await importFromFile(file);
      const calculated = recalculateWorkbook(importedWb);
      setWorkbook(calculated);
      setCharts([]);
      setActiveCell({ row: 0, col: 0 });
      setSelectedRange({ startRow: 0, startCol: 0, endRow: 0, endCol: 0 });
    } catch (err: any) {
      alert(`Could not parse Excel file: ${err?.message || 'Invalid format'}`);
    }
  }, [recalculateWorkbook]);

  // Template Switching
  const handleSelectTemplate = useCallback((templateKey: any) => {
    const data = getTemplateWorkbook(templateKey);
    const calculated = recalculateWorkbook(data.workbook);
    setWorkbook(calculated);
    setCharts(data.charts);
    setActiveCell({ row: 4, col: 1 });
    setSelectedRange({ startRow: 4, startCol: 1, endRow: 4, endCol: 1 });
  }, [recalculateWorkbook]);

  // Chart Management
  const handleSaveChart = useCallback((chartConfig: ChartConfig) => {
    setCharts((prev) => {
      const exists = prev.some(c => c.id === chartConfig.id);
      if (exists) {
        return prev.map(c => c.id === chartConfig.id ? chartConfig : c);
      }
      return [...prev, chartConfig];
    });
  }, []);

  const handleDeleteChart = useCallback((chartId: string) => {
    setCharts((prev) => prev.filter(c => c.id !== chartId));
  }, []);

  return (
    <div className="flex flex-col h-screen w-screen overflow-hidden bg-[#f3f4f6]">
      {/* Top Application Bar */}
      <Header
        workbook={workbook}
        activeView={activeView}
        onViewChange={setActiveView}
        onExportExcel={handleExportExcel}
        onExportCSV={handleExportCSV}
        onImportFile={handleImportFile}
        onSelectTemplate={handleSelectTemplate}
        onFileNameChange={(name) => setWorkbook((prev) => ({ ...prev, fileName: name }))}
        onOpenNewChartModal={() => {
          setEditingChart(null);
          setIsCreateChartOpen(true);
        }}
      />

      {/* Main Workspace Body */}
      {activeView === 'grid' ? (
        <div className="flex flex-col flex-1 min-h-0 overflow-hidden">
          {/* Microsoft Excel Style Ribbon */}
          <Ribbon
            activeTab={activeRibbonTab}
            onTabChange={setActiveRibbonTab}
            activeCellFormat={activeCellData?.format}
            onApplyFormat={handleApplyFormat}
            onAutoSum={handleAutoSum}
            onSort={handleSort}
            onToggleFilter={() => {}}
            onOpenInsertFunction={() => setIsInsertFunctionOpen(true)}
            onOpenFindReplace={() => setIsFindReplaceOpen(true)}
            onOpenDataValidation={() => setIsDataValidationOpen(true)}
            onOpenNewChart={(type) => {
              setEditingChart(type ? {
                id: `chart-${Date.now()}`,
                title: 'Data Chart',
                type,
                sheetId: activeSheet.id,
                dataRange: formatRange(selectedRange),
                labelColumn: 'A',
                seriesColumns: ['B'],
                colors: ['#107c41', '#3b82f6'],
                showLegend: true,
                showGrid: true,
                showValues: true,
                isDashboardWidget: true,
              } : null);
              setIsCreateChartOpen(true);
            }}
            onInsertRow={handleInsertRow}
            onInsertCol={handleInsertCol}
            onDeleteRow={handleDeleteRow}
            onDeleteCol={handleDeleteCol}
            onClearContent={handleClearContent}
            onMergeCenter={() => handleApplyFormat({ align: 'center' })}
            onRecalculate={() => setWorkbook(recalculateWorkbook(workbook))}
            onUndo={handleUndo}
            onRedo={handleRedo}
            canUndo={undoStack.length > 0}
            canRedo={redoStack.length > 0}
            onCopy={handleCopy}
            onCut={handleCut}
            onPaste={handlePaste}
            showGridlines={showGridlines}
            onToggleGridlines={() => setShowGridlines(!showGridlines)}
            showFormulaBar={showFormulaBar}
            onToggleFormulaBar={() => setShowFormulaBar(!showFormulaBar)}
            zoom={zoom}
            onZoomChange={setZoom}
            onApplyConditionalScale={handleApplyConditionalScale}
            onApplyConditionalRule={handleApplyConditionalRule}
            onFreezeTopRow={handleFreezeTopRow}
            onFreezeFirstCol={handleFreezeFirstCol}
            onUnfreeze={handleUnfreeze}
          />

          {/* Formula Bar */}
          {showFormulaBar && (
            <FormulaBar
              selectedCellId={activeCellId}
              selectedRangeStr={formatRange(selectedRange)}
              cellRawValue={activeCellData?.raw || ''}
              onRawValueChange={(val) => handleCellChange(activeCellId, val)}
              onCommit={() => {}}
              onCancel={() => {}}
              onOpenInsertFunction={() => setIsInsertFunctionOpen(true)}
              onJumpToCell={handleJumpToCell}
            />
          )}

          {/* Interactive Spreadsheet Grid */}
          <SpreadsheetGrid
            sheet={activeSheet}
            activeCell={activeCell}
            selectedRange={selectedRange}
            onActiveCellChange={setActiveCell}
            onRangeSelect={setSelectedRange}
            onCellChange={handleCellChange}
            onAutofill={handleAutofill}
            showGridlines={showGridlines}
            zoom={zoom}
            onCut={handleCut}
            onCopy={handleCopy}
            onPaste={handlePaste}
            onInsertRow={handleInsertRow}
            onInsertCol={handleInsertCol}
            onDeleteRow={handleDeleteRow}
            onDeleteCol={handleDeleteCol}
            onClearContent={handleClearContent}
            onCreateChartFromSelection={() => {
              setEditingChart({
                id: `chart-${Date.now()}`,
                title: 'Data Chart',
                type: 'column',
                sheetId: activeSheet.id,
                dataRange: formatRange(selectedRange),
                labelColumn: 'A',
                seriesColumns: ['B', 'C'],
                colors: ['#107c41', '#3b82f6', '#f59e0b'],
                showLegend: true,
                showGrid: true,
                showValues: true,
                isDashboardWidget: true,
              });
              setIsCreateChartOpen(true);
            }}
            onOpenFormatCells={() => setIsFormatCellsOpen(true)}
            onResizeCol={handleResizeCol}
            onResizeRow={handleResizeRow}
          />

          {/* Bottom Sheet Navigation Tabs */}
          <SheetTabs
            sheets={workbook.sheets}
            activeSheetId={workbook.activeSheetId}
            onSelectSheet={(id) => setWorkbook((prev) => ({ ...prev, activeSheetId: id }))}
            onAddSheet={handleAddSheet}
            onRenameSheet={handleRenameSheet}
            onDeleteSheet={handleDeleteSheet}
            onDuplicateSheet={handleDuplicateSheet}
            onChangeTabColor={handleChangeTabColor}
          />
        </div>
      ) : (
        /* Executive Dashboard Studio View */
        <DashboardView
          workbook={workbook}
          charts={charts}
          onOpenNewChart={(type) => {
            setEditingChart(type ? {
              id: `chart-${Date.now()}`,
              title: type === 'kpi' ? 'KPI Summary' : 'Analytics Chart',
              type,
              sheetId: activeSheet.id,
              dataRange: formatRange(selectedRange),
              labelColumn: 'A',
              seriesColumns: ['B'],
              colors: ['#107c41', '#3b82f6'],
              showLegend: true,
              showGrid: true,
              showValues: true,
              isDashboardWidget: true,
            } : null);
            setIsCreateChartOpen(true);
          }}
          onEditChart={(chart) => {
            setEditingChart(chart);
            setIsCreateChartOpen(true);
          }}
          onDeleteChart={handleDeleteChart}
          onReturnToGrid={() => setActiveView('grid')}
        />
      )}

      {/* Excel Status Bar (Ready, Math Averages & Counts, Zoom) */}
      <StatusBar
        sheet={activeSheet}
        selectedRange={selectedRange}
        activeView={activeView}
        onViewChange={setActiveView}
        zoom={zoom}
        onZoomChange={setZoom}
      />

      {/* ================= MODAL DIALOGS ================= */}
      <InsertFunctionModal
        isOpen={isInsertFunctionOpen}
        onClose={() => setIsInsertFunctionOpen(false)}
        onInsertFunction={(formula) => handleCellChange(activeCellId, formula)}
      />

      <CreateChartModal
        isOpen={isCreateChartOpen}
        onClose={() => {
          setIsCreateChartOpen(false);
          setEditingChart(null);
        }}
        sheets={workbook.sheets}
        activeSheetId={workbook.activeSheetId}
        defaultRange={formatRange(selectedRange)}
        initialChart={editingChart}
        onSaveChart={handleSaveChart}
      />

      <FindReplaceModal
        isOpen={isFindReplaceOpen}
        onClose={() => setIsFindReplaceOpen(false)}
        sheet={activeSheet}
        onFindNext={(query, matchCase, exact) => {
          for (let r = 0; r < activeSheet.rowCount; r++) {
            for (let c = 0; c < activeSheet.colCount; c++) {
              const cId = coordsToCellId(r, c);
              const cell = activeSheet.cells?.[cId];
              const txt = String(cell?.value ?? cell?.raw ?? '');
              const matches = exact
                ? (matchCase ? txt === query : txt.toLowerCase() === query.toLowerCase())
                : (matchCase ? txt.includes(query) : txt.toLowerCase().includes(query.toLowerCase()));

              if (matches) {
                setActiveCell({ row: r, col: c });
                setSelectedRange({ startRow: r, startCol: c, endRow: r, endCol: c });
                return;
              }
            }
          }
          alert('No matches found.');
        }}
        onReplaceCurrent={(find, replace) => {
          if (activeCellData?.raw) {
            handleCellChange(activeCellId, activeCellData.raw.replace(find, replace));
          }
        }}
        onReplaceAll={(find, replace, matchCase, exact) => {
          let count = 0;
          pushUndo(workbook);
          setWorkbook((prevWb) => {
            const newSheets = prevWb.sheets.map((sheet) => {
              if (sheet.id !== prevWb.activeSheetId) return sheet;
              const newCells = { ...sheet.cells };
              for (const [cId, cell] of Object.entries(sheet.cells)) {
                if (!cell || !cell.raw) continue;
                const txt = cell.raw;
                const matches = exact
                  ? (matchCase ? txt === find : txt.toLowerCase() === find.toLowerCase())
                  : (matchCase ? txt.includes(find) : txt.toLowerCase().includes(find.toLowerCase()));

                if (matches) {
                  newCells[cId] = {
                    ...cell,
                    raw: exact ? replace : txt.split(find).join(replace)
                  };
                  count++;
                }
              }
              return { ...sheet, cells: newCells };
            });
            return recalculateWorkbook({ ...prevWb, sheets: newSheets });
          });
          alert(`Replaced ${count} occurrences.`);
        }}
      />

      <FormatCellsModal
        isOpen={isFormatCellsOpen}
        onClose={() => setIsFormatCellsOpen(false)}
        currentFormat={activeCellData?.format}
        onApplyFormat={handleApplyFormat}
      />

      <DataValidationModal
        isOpen={isDataValidationOpen}
        onClose={() => setIsDataValidationOpen(false)}
        selectedRangeStr={formatRange(selectedRange)}
        onApplyValidation={(rule: DataValidationRule) => {
          setWorkbook((prev) => ({
            ...prev,
            sheets: prev.sheets.map(s => s.id === prev.activeSheetId ? {
              ...s,
              validationRules: { ...s.validationRules, [formatRange(selectedRange)]: rule }
            } : s)
          }));
        }}
        onClearValidation={() => {
          setWorkbook((prev) => ({
            ...prev,
            sheets: prev.sheets.map(s => s.id === prev.activeSheetId ? {
              ...s,
              validationRules: {}
            } : s)
          }));
        }}
      />
    </div>
  );
}
