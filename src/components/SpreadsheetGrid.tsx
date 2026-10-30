/**
 * @license
 * SPDX-License-Identifier: Apache-2.0
 */

import React, { useState, useRef, useEffect, useCallback } from 'react';
import { SheetData, CellCoordinate, CellRange, CellFormat } from '../types/excel';
import { coordsToCellId, colIndexToLetter, isCellInRange, shiftFormula } from '../utils/cellUtils';
import { formatCellValue } from '../utils/numberFormat';
import { evaluateConditionalFormat } from '../utils/conditionalFormatting';
import { CellContextMenu } from './CellContextMenu';

interface SpreadsheetGridProps {
  sheet: SheetData;
  activeCell: CellCoordinate;
  selectedRange: CellRange;
  onActiveCellChange: (coord: CellCoordinate) => void;
  onRangeSelect: (range: CellRange) => void;
  onCellChange: (cellId: string, raw: string) => void;
  onAutofill: (sourceRange: CellRange, targetRange: CellRange) => void;
  showGridlines: boolean;
  zoom: number;
  onCut: () => void;
  onCopy: () => void;
  onPaste: () => void;
  onInsertRow: () => void;
  onInsertCol: () => void;
  onDeleteRow: () => void;
  onDeleteCol: () => void;
  onClearContent: () => void;
  onCreateChartFromSelection: () => void;
  onOpenFormatCells: () => void;
  onResizeCol: (colIndex: number, width: number) => void;
  onResizeRow: (rowIndex: number, height: number) => void;
}

const DEFAULT_COL_WIDTH = 100;
const DEFAULT_ROW_HEIGHT = 24;
const HEADER_ROW_HEIGHT = 26;
const HEADER_COL_WIDTH = 48;

export const SpreadsheetGrid: React.FC<SpreadsheetGridProps> = ({
  sheet,
  activeCell,
  selectedRange,
  onActiveCellChange,
  onRangeSelect,
  onCellChange,
  onAutofill,
  showGridlines,
  zoom,
  onCut,
  onCopy,
  onPaste,
  onInsertRow,
  onInsertCol,
  onDeleteRow,
  onDeleteCol,
  onClearContent,
  onCreateChartFromSelection,
  onOpenFormatCells,
  onResizeCol,
  onResizeRow,
}) => {
  const gridContainerRef = useRef<HTMLDivElement>(null);
  const cellInputRef = useRef<HTMLInputElement>(null);

  // Editing state
  const [isEditing, setIsEditing] = useState(false);
  const [editValue, setEditValue] = useState('');

  // Drag selection state
  const [isSelecting, setIsSelecting] = useState(false);
  const [selectionAnchor, setSelectionAnchor] = useState<CellCoordinate | null>(null);

  // Autofill drag state
  const [isAutofilling, setIsAutofilling] = useState(false);
  const [autofillEndCoord, setAutofillEndCoord] = useState<CellCoordinate | null>(null);

  // Context menu state
  const [contextMenu, setContextMenu] = useState<{ x: number; y: number } | null>(null);

  // Column / Row resize dragging
  const [resizingCol, setResizingCol] = useState<{ colIndex: number; startX: number; startWidth: number } | null>(null);
  const [resizingRow, setResizingRow] = useState<{ rowIndex: number; startY: number; startHeight: number } | null>(null);

  const activeCellId = coordsToCellId(activeCell.row, activeCell.col);
  const activeCellData = sheet.cells?.[activeCellId];

  // Sync edit value when cell changes
  useEffect(() => {
    if (!isEditing) {
      setEditValue(activeCellData?.raw ?? '');
    }
  }, [activeCellId, activeCellData, isEditing]);

  // Handle in-cell input auto focus
  useEffect(() => {
    if (isEditing && cellInputRef.current) {
      cellInputRef.current.focus();
    }
  }, [isEditing]);

  const commitEdit = useCallback(() => {
    if (isEditing) {
      onCellChange(activeCellId, editValue);
      setIsEditing(false);
    }
  }, [isEditing, activeCellId, editValue, onCellChange]);

  const cancelEdit = useCallback(() => {
    if (isEditing) {
      setEditValue(activeCellData?.raw ?? '');
      setIsEditing(false);
    }
  }, [isEditing, activeCellData]);

  // Global Mouse Up for drag operations
  useEffect(() => {
    const handleGlobalMouseUp = () => {
      if (isSelecting) {
        setIsSelecting(false);
      }
      if (isAutofilling && autofillEndCoord) {
        setIsAutofilling(false);
        const targetRange: CellRange = {
          startRow: Math.min(selectedRange.startRow, autofillEndCoord.row),
          startCol: Math.min(selectedRange.startCol, autofillEndCoord.col),
          endRow: Math.max(selectedRange.endRow, autofillEndCoord.row),
          endCol: Math.max(selectedRange.endCol, autofillEndCoord.col),
        };
        onAutofill(selectedRange, targetRange);
        setAutofillEndCoord(null);
      } else if (isAutofilling) {
        setIsAutofilling(false);
        setAutofillEndCoord(null);
      }

      if (resizingCol) {
        setResizingCol(null);
      }
      if (resizingRow) {
        setResizingRow(null);
      }
    };

    const handleGlobalMouseMove = (e: MouseEvent) => {
      if (resizingCol) {
        const delta = e.clientX - resizingCol.startX;
        const newWidth = Math.max(40, resizingCol.startWidth + delta);
        onResizeCol(resizingCol.colIndex, newWidth);
      } else if (resizingRow) {
        const delta = e.clientY - resizingRow.startY;
        const newHeight = Math.max(18, resizingRow.startHeight + delta);
        onResizeRow(resizingRow.rowIndex, newHeight);
      }
    };

    window.addEventListener('mouseup', handleGlobalMouseUp);
    window.addEventListener('mousemove', handleGlobalMouseMove);
    return () => {
      window.removeEventListener('mouseup', handleGlobalMouseUp);
      window.removeEventListener('mousemove', handleGlobalMouseMove);
    };
  }, [isSelecting, isAutofilling, autofillEndCoord, selectedRange, onAutofill, resizingCol, resizingRow, onResizeCol, onResizeRow]);

  // Keyboard navigation
  const handleKeyDown = (e: React.KeyboardEvent) => {
    if (isEditing) {
      if (e.key === 'Enter') {
        e.preventDefault();
        commitEdit();
        // Move selection down
        if (activeCell.row < sheet.rowCount - 1) {
          const newCoord = { row: activeCell.row + 1, col: activeCell.col };
          onActiveCellChange(newCoord);
          onRangeSelect({ startRow: newCoord.row, startCol: newCoord.col, endRow: newCoord.row, endCol: newCoord.col });
        }
      } else if (e.key === 'Tab') {
        e.preventDefault();
        commitEdit();
        // Move selection right
        if (activeCell.col < sheet.colCount - 1) {
          const newCoord = { row: activeCell.row, col: activeCell.col + 1 };
          onActiveCellChange(newCoord);
          onRangeSelect({ startRow: newCoord.row, startCol: newCoord.col, endRow: newCoord.row, endCol: newCoord.col });
        }
      } else if (e.key === 'Escape') {
        e.preventDefault();
        cancelEdit();
      }
      return;
    }

    // Grid Navigation Keys
    let handled = false;
    let nextRow = activeCell.row;
    let nextCol = activeCell.col;

    switch (e.key) {
      case 'ArrowUp':
        nextRow = Math.max(0, activeCell.row - 1);
        handled = true;
        break;
      case 'ArrowDown':
        nextRow = Math.min(sheet.rowCount - 1, activeCell.row + 1);
        handled = true;
        break;
      case 'ArrowLeft':
        nextCol = Math.max(0, activeCell.col - 1);
        handled = true;
        break;
      case 'ArrowRight':
        nextCol = Math.min(sheet.colCount - 1, activeCell.col + 1);
        handled = true;
        break;
      case 'Tab':
        e.preventDefault();
        if (e.shiftKey) {
          nextCol = Math.max(0, activeCell.col - 1);
        } else {
          nextCol = Math.min(sheet.colCount - 1, activeCell.col + 1);
        }
        handled = true;
        break;
      case 'Enter':
        e.preventDefault();
        if (e.shiftKey) {
          nextRow = Math.max(0, activeCell.row - 1);
        } else {
          nextRow = Math.min(sheet.rowCount - 1, activeCell.row + 1);
        }
        handled = true;
        break;
      case 'Delete':
      case 'Backspace':
        onClearContent();
        handled = true;
        break;
      case 'F2':
        setIsEditing(true);
        handled = true;
        break;
    }

    // Shortcut checks
    if ((e.ctrlKey || e.metaKey) && e.key.toLowerCase() === 'c') {
      onCopy();
      handled = true;
    } else if ((e.ctrlKey || e.metaKey) && e.key.toLowerCase() === 'v') {
      onPaste();
      handled = true;
    } else if ((e.ctrlKey || e.metaKey) && e.key.toLowerCase() === 'x') {
      onCut();
      handled = true;
    }

    if (handled) {
      if (nextRow !== activeCell.row || nextCol !== activeCell.col) {
        const newCoord = { row: nextRow, col: nextCol };
        onActiveCellChange(newCoord);
        if (e.shiftKey && e.key.startsWith('Arrow')) {
          onRangeSelect({
            startRow: Math.min(selectedRange.startRow, nextRow),
            startCol: Math.min(selectedRange.startCol, nextCol),
            endRow: Math.max(selectedRange.endRow, nextRow),
            endCol: Math.max(selectedRange.endCol, nextCol),
          });
        } else {
          onRangeSelect({ startRow: nextRow, startCol: nextCol, endRow: nextRow, endCol: nextCol });
        }
      }
      return;
    }

    // Typing directly into cell starts edit mode!
    if (!e.ctrlKey && !e.metaKey && !e.altKey && e.key.length === 1) {
      setIsEditing(true);
      setEditValue(e.key);
    }
  };

  // Cell Click handlers
  const handleCellMouseDown = (row: number, col: number, e: React.MouseEvent) => {
    if (e.button !== 0) return; // Only primary mouse button
    commitEdit();

    if (e.shiftKey) {
      // Extend range
      onRangeSelect({
        startRow: Math.min(selectedRange.startRow, row),
        startCol: Math.min(selectedRange.startCol, col),
        endRow: Math.max(selectedRange.endRow, row),
        endCol: Math.max(selectedRange.endCol, col),
      });
    } else {
      const coord = { row, col };
      onActiveCellChange(coord);
      onRangeSelect({ startRow: row, startCol: col, endRow: row, endCol: col });
      setSelectionAnchor(coord);
      setIsSelecting(true);
    }
  };

  const handleCellMouseEnter = (row: number, col: number) => {
    if (isSelecting && selectionAnchor) {
      onRangeSelect({
        startRow: Math.min(selectionAnchor.row, row),
        startCol: Math.min(selectionAnchor.col, col),
        endRow: Math.max(selectionAnchor.row, row),
        endCol: Math.max(selectionAnchor.col, col),
      });
    } else if (isAutofilling) {
      setAutofillEndCoord({ row, col });
    }
  };

  // Header Selection
  const handleSelectColumn = (col: number) => {
    commitEdit();
    onActiveCellChange({ row: 0, col });
    onRangeSelect({ startRow: 0, startCol: col, endRow: sheet.rowCount - 1, endCol: col });
  };

  const handleSelectRow = (row: number) => {
    commitEdit();
    onActiveCellChange({ row, col: 0 });
    onRangeSelect({ startRow: row, startCol: 0, endRow: row, endCol: sheet.colCount - 1 });
  };

  const handleSelectAll = () => {
    commitEdit();
    onActiveCellChange({ row: 0, col: 0 });
    onRangeSelect({ startRow: 0, startCol: 0, endRow: sheet.rowCount - 1, endCol: sheet.colCount - 1 });
  };

  // Right-click context menu
  const handleContextMenu = (e: React.MouseEvent, row: number, col: number) => {
    e.preventDefault();
    if (!isCellInRange(row, col, selectedRange)) {
      onActiveCellChange({ row, col });
      onRangeSelect({ startRow: row, startCol: col, endRow: row, endCol: col });
    }
    setContextMenu({ x: e.clientX, y: e.clientY });
  };

  // Calculate cell positioning for selection highlight and autofill handle
  const getColWidth = (c: number) => sheet.colWidths[c] || DEFAULT_COL_WIDTH;
  const getRowHeight = (r: number) => sheet.rowHeights[r] || DEFAULT_ROW_HEIGHT;

  let selectionLeft = HEADER_COL_WIDTH;
  for (let c = 0; c < selectedRange.startCol; c++) {
    selectionLeft += getColWidth(c);
  }

  let selectionTop = HEADER_ROW_HEIGHT;
  for (let r = 0; r < selectedRange.startRow; r++) {
    selectionTop += getRowHeight(r);
  }

  let selectionWidth = 0;
  for (let c = selectedRange.startCol; c <= selectedRange.endCol; c++) {
    selectionWidth += getColWidth(c);
  }

  let selectionHeight = 0;
  for (let r = selectedRange.startRow; r <= selectedRange.endRow; r++) {
    selectionHeight += getRowHeight(r);
  }

  return (
    <div
      ref={gridContainerRef}
      tabIndex={0}
      onKeyDown={handleKeyDown}
      onClick={() => setContextMenu(null)}
      className="flex-1 bg-white overflow-auto relative outline-none select-none"
      style={{ transform: `scale(${zoom})`, transformOrigin: 'top left' }}
    >
      <div className="inline-block min-w-full">
        {/* Table Structure */}
        <table className="border-collapse table-fixed text-xs text-slate-800">
          <thead>
            <tr className="sticky top-0 z-20 bg-[#f8fafc]">
              {/* Select All Top-Left Corner */}
              <th
                onClick={handleSelectAll}
                className="sticky left-0 z-30 w-12 border-b border-r border-slate-300 bg-slate-100 hover:bg-slate-200 cursor-pointer font-normal text-slate-500"
                style={{ width: HEADER_COL_WIDTH, height: HEADER_ROW_HEIGHT }}
              >
                <div className="w-2.5 h-2.5 border-r border-b border-slate-400 m-auto" />
              </th>

              {/* Column Headers A, B, C... */}
              {Array.from({ length: sheet.colCount }).map((_, colIdx) => {
                const colLetter = colIndexToLetter(colIdx);
                const width = getColWidth(colIdx);
                const isSelected = colIdx >= selectedRange.startCol && colIdx <= selectedRange.endCol;

                return (
                  <th
                    key={colIdx}
                    onClick={() => handleSelectColumn(colIdx)}
                    style={{ width, minWidth: width, maxWidth: width, height: HEADER_ROW_HEIGHT }}
                    className={`relative border-b border-r border-slate-300 font-semibold text-center select-none cursor-pointer transition-colors ${
                      isSelected ? 'bg-emerald-100 text-[#107c41]' : 'bg-[#f8fafc] text-slate-600 hover:bg-slate-100'
                    }`}
                  >
                    <span className="truncate block px-1">{colLetter}</span>

                    {/* Column Resizer Divider */}
                    <div
                      onMouseDown={(e) => {
                        e.stopPropagation();
                        setResizingCol({ colIndex: colIdx, startX: e.clientX, startWidth: width });
                      }}
                      onDoubleClick={(e) => {
                        e.stopPropagation();
                        // Auto-fit column based on contents
                        let maxChars = 8;
                        for (let r = 0; r < sheet.rowCount; r++) {
                          const cId = coordsToCellId(r, colIdx);
                          const cell = sheet.cells?.[cId];
                          const txt = cell?.value !== undefined ? String(cell.value) : cell?.raw || '';
                          if (txt.length > maxChars) maxChars = txt.length;
                        }
                        onResizeCol(colIdx, Math.min(320, Math.max(60, maxChars * 9 + 24)));
                      }}
                      className="absolute right-0 top-0 bottom-0 w-1.5 cursor-col-resize hover:bg-[#107c41] z-10"
                    />
                  </th>
                );
              })}
            </tr>
          </thead>

          <tbody>
            {Array.from({ length: sheet.rowCount }).map((_, rowIdx) => {
              const rowHeight = getRowHeight(rowIdx);
              const isRowSelected = rowIdx >= selectedRange.startRow && rowIdx <= selectedRange.endRow;

              return (
                <tr key={rowIdx} style={{ height: rowHeight }}>
                  {/* Sticky Row Header 1, 2, 3... */}
                  <th
                    onClick={() => handleSelectRow(rowIdx)}
                    style={{ width: HEADER_COL_WIDTH, height: rowHeight }}
                    className={`sticky left-0 z-10 border-b border-r border-slate-300 font-mono font-medium text-center select-none cursor-pointer transition-colors ${
                      isRowSelected ? 'bg-emerald-100 text-[#107c41]' : 'bg-[#f8fafc] text-slate-500 hover:bg-slate-100'
                    }`}
                  >
                    <span className="text-[11px] block">{rowIdx + 1}</span>

                    {/* Row Resizer Divider */}
                    <div
                      onMouseDown={(e) => {
                        e.stopPropagation();
                        setResizingRow({ rowIndex: rowIdx, startY: e.clientY, startHeight: rowHeight });
                      }}
                      className="absolute left-0 right-0 bottom-0 h-1.5 cursor-row-resize hover:bg-[#107c41] z-10"
                    />
                  </th>

                  {/* Cell Columns */}
                  {Array.from({ length: sheet.colCount }).map((_, colIdx) => {
                    const cellId = coordsToCellId(rowIdx, colIdx);
                    const cell = sheet.cells?.[cellId];
                    const isActive = rowIdx === activeCell.row && colIdx === activeCell.col;
                    const inRange = isCellInRange(rowIdx, colIdx, selectedRange);

                    // Formatted display value
                    const displayValue = formatCellValue(cell?.value ?? cell?.raw, cell?.format);

                    // Conditional formatting overrides
                    const condStyle = evaluateConditionalFormat(rowIdx, colIdx, cell?.value ?? cell?.raw, sheet);

                    // Cell styles
                    const format = cell?.format || {};
                    const style: React.CSSProperties = {
                      width: getColWidth(colIdx),
                      height: rowHeight,
                      fontWeight: format.bold ? 'bold' : 'normal',
                      fontStyle: format.italic ? 'italic' : 'normal',
                      textDecoration: `${format.underline ? 'underline ' : ''}${format.strikethrough ? 'line-through' : ''}`.trim() || undefined,
                      fontSize: format.fontSize ? `${format.fontSize}px` : undefined,
                      fontFamily: format.fontFamily ? `${format.fontFamily}, sans-serif` : undefined,
                      color: condStyle?.textColor || format.textColor || '#0f172a',
                      backgroundColor: condStyle?.bgColor || format.bgColor || (inRange && !isActive ? '#f0fdf4' : '#ffffff'),
                      textAlign: format.align || (typeof cell?.value === 'number' ? 'right' : 'left'),
                      verticalAlign: format.verticalAlign || 'middle',
                      whiteSpace: format.wrapText ? 'normal' : 'nowrap',
                    };

                    // Borders
                    if (format.borders) {
                      if (format.borders.top) style.borderTop = `1px solid ${format.borders.top}`;
                      if (format.borders.bottom) style.borderBottom = `1px solid ${format.borders.bottom}`;
                      if (format.borders.left) style.borderLeft = `1px solid ${format.borders.left}`;
                      if (format.borders.right) style.borderRight = `1px solid ${format.borders.right}`;
                    }

                    return (
                      <td
                        key={colIdx}
                        onMouseDown={(e) => handleCellMouseDown(rowIdx, colIdx, e)}
                        onMouseEnter={() => handleCellMouseEnter(rowIdx, colIdx)}
                        onDoubleClick={() => setIsEditing(true)}
                        onContextMenu={(e) => handleContextMenu(e, rowIdx, colIdx)}
                        style={style}
                        className={`relative px-1.5 py-0.5 truncate overflow-hidden ${
                          showGridlines ? 'border-b border-r border-slate-200' : ''
                        }`}
                      >
                        {isActive && isEditing ? (
                          <input
                            ref={cellInputRef}
                            type="text"
                            value={editValue}
                            onChange={(e) => setEditValue(e.target.value)}
                            onBlur={commitEdit}
                            style={{
                              textAlign: style.textAlign,
                              fontWeight: style.fontWeight,
                              fontFamily: style.fontFamily,
                            }}
                            className="absolute inset-0 w-full h-full px-1.5 font-mono text-xs bg-white text-slate-900 border-2 border-[#107c41] outline-none z-30 shadow-md"
                          />
                        ) : (
                          <span className={typeof cell?.value === 'number' ? 'tabular-nums' : ''}>
                            {displayValue}
                          </span>
                        )}
                      </td>
                    );
                  })}
                </tr>
              );
            })}
          </tbody>
        </table>

        {/* Selection Rectangle & Excel Fill Handle */}
        <div
          className="absolute pointer-events-none z-10 transition-all duration-75"
          style={{
            left: selectionLeft,
            top: selectionTop,
            width: selectionWidth,
            height: selectionHeight,
            border: '2px solid #107c41',
            backgroundColor: 'rgba(16, 124, 65, 0.04)',
          }}
        >
          {/* Fill Handle (Bottom-Right green square like real Excel) */}
          <div
            onMouseDown={(e) => {
              e.stopPropagation();
              e.preventDefault();
              setIsAutofilling(true);
            }}
            className="absolute -right-1 -bottom-1 w-2.5 h-2.5 bg-[#107c41] border border-white cursor-crosshair pointer-events-auto shadow-xs hover:scale-125 transition-transform"
            title="Autofill series / copy formula"
          />
        </div>

        {/* Autofill Ghost Box while dragging handle */}
        {isAutofilling && autofillEndCoord && (
          <div
            className="absolute pointer-events-none z-10 border-2 border-dashed border-emerald-600 bg-emerald-500/10"
            style={{
              left: Math.min(selectionLeft, HEADER_COL_WIDTH + (autofillEndCoord.col * DEFAULT_COL_WIDTH)),
              top: Math.min(selectionTop, HEADER_ROW_HEIGHT + (autofillEndCoord.row * DEFAULT_ROW_HEIGHT)),
              width: Math.max(selectionWidth, (autofillEndCoord.col - selectedRange.startCol + 1) * DEFAULT_COL_WIDTH),
              height: Math.max(selectionHeight, (autofillEndCoord.row - selectedRange.startRow + 1) * DEFAULT_ROW_HEIGHT),
            }}
          />
        )}
      </div>

      {/* Right-click Context Menu */}
      {contextMenu && (
        <CellContextMenu
          x={contextMenu.x}
          y={contextMenu.y}
          onClose={() => setContextMenu(null)}
          onCut={onCut}
          onCopy={onCopy}
          onPaste={onPaste}
          onInsertRow={onInsertRow}
          onInsertCol={onInsertCol}
          onDeleteRow={onDeleteRow}
          onDeleteCol={onDeleteCol}
          onClearContent={onClearContent}
          onCreateChartFromSelection={onCreateChartFromSelection}
          onFormatCells={onOpenFormatCells}
        />
      )}
    </div>
  );
};
