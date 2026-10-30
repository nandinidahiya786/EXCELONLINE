/**
 * @license
 * SPDX-License-Identifier: Apache-2.0
 */

import React, { useMemo } from 'react';
import { SheetData, CellRange } from '../types/excel';
import { coordsToCellId } from '../utils/cellUtils';
import { Table, LayoutDashboard, ZoomIn, ZoomOut } from 'lucide-react';

interface StatusBarProps {
  sheet: SheetData;
  selectedRange: CellRange;
  activeView: 'grid' | 'dashboard';
  onViewChange: (view: 'grid' | 'dashboard') => void;
  zoom: number;
  onZoomChange: (zoom: number) => void;
}

export const StatusBar: React.FC<StatusBarProps> = ({
  sheet,
  selectedRange,
  activeView,
  onViewChange,
  zoom,
  onZoomChange,
}) => {
  // Compute summary stats for current selection
  const stats = useMemo(() => {
    let totalCount = 0;
    let numericCount = 0;
    let sum = 0;
    let min = Infinity;
    let max = -Infinity;

    for (let r = selectedRange.startRow; r <= selectedRange.endRow; r++) {
      for (let c = selectedRange.startCol; c <= selectedRange.endCol; c++) {
        totalCount++;
        const cellId = coordsToCellId(r, c);
        const cell = sheet.cells?.[cellId];
        if (!cell || (cell.raw === '' && cell.value === undefined)) continue;

        const val = cell.value !== undefined ? cell.value : Number(cell.raw);
        if (typeof val === 'number' && !isNaN(val)) {
          numericCount++;
          sum += val;
          if (val < min) min = val;
          if (val > max) max = val;
        }
      }
    }

    const average = numericCount > 0 ? sum / numericCount : 0;
    return {
      totalCount,
      numericCount,
      sum,
      average,
      min: min === Infinity ? 0 : min,
      max: max === -Infinity ? 0 : max,
    };
  }, [sheet, selectedRange]);

  const hasMultipleSelection = (selectedRange.endRow - selectedRange.startRow > 0) || (selectedRange.endCol - selectedRange.startCol > 0);

  return (
    <footer className="bg-slate-100 border-t border-slate-200 px-3 py-1 flex items-center justify-between text-xs text-slate-600 select-none">
      {/* Left: Mode / Status */}
      <div className="flex items-center gap-2">
        <span className="font-medium text-[#107c41]">Ready</span>
        <span className="text-slate-300">|</span>
        <span className="text-[11px] text-slate-400">100% Client-Side Engine</span>
      </div>

      {/* Center: Live Excel Statistics Bar (Average, Count, Sum, Min, Max) */}
      <div className="flex items-center gap-4 font-mono tabular-nums text-[11px]">
        {hasMultipleSelection && (
          <>
            {stats.numericCount > 0 && (
              <>
                <div>
                  <span className="text-slate-400 font-sans mr-1">AVERAGE:</span>
                  <span className="font-semibold text-slate-700">
                    {stats.average.toLocaleString(undefined, { maximumFractionDigits: 2 })}
                  </span>
                </div>
                <div>
                  <span className="text-slate-400 font-sans mr-1">COUNT:</span>
                  <span className="font-semibold text-slate-700">{stats.totalCount}</span>
                </div>
                <div>
                  <span className="text-slate-400 font-sans mr-1">MIN:</span>
                  <span className="font-semibold text-slate-700">{stats.min.toLocaleString()}</span>
                </div>
                <div>
                  <span className="text-slate-400 font-sans mr-1">MAX:</span>
                  <span className="font-semibold text-slate-700">{stats.max.toLocaleString()}</span>
                </div>
                <div>
                  <span className="text-slate-400 font-sans mr-1">SUM:</span>
                  <span className="font-semibold text-[#107c41]">
                    {stats.sum.toLocaleString(undefined, { maximumFractionDigits: 2 })}
                  </span>
                </div>
              </>
            )}
            {stats.numericCount === 0 && stats.totalCount > 1 && (
              <div>
                <span className="text-slate-400 font-sans mr-1">COUNT:</span>
                <span className="font-semibold text-slate-700">{stats.totalCount}</span>
              </div>
            )}
          </>
        )}
      </div>

      {/* Right: View Toggles & Zoom Slider */}
      <div className="flex items-center gap-3">
        <div className="flex items-center gap-1">
          <button
            onClick={() => onViewChange('grid')}
            className={`p-1 rounded transition-colors ${activeView === 'grid' ? 'bg-white text-[#107c41] shadow-xs' : 'hover:bg-slate-200 text-slate-500'}`}
            title="Normal Grid View"
          >
            <Table className="w-3.5 h-3.5" />
          </button>
          <button
            onClick={() => onViewChange('dashboard')}
            className={`p-1 rounded transition-colors ${activeView === 'dashboard' ? 'bg-white text-[#107c41] shadow-xs' : 'hover:bg-slate-200 text-slate-500'}`}
            title="Dashboard Studio View"
          >
            <LayoutDashboard className="w-3.5 h-3.5" />
          </button>
        </div>

        <div className="h-3 w-[1px] bg-slate-300" />

        {/* Zoom Slider */}
        <div className="flex items-center gap-1.5">
          <button
            onClick={() => onZoomChange(Math.max(0.5, Math.round((zoom - 0.1) * 10) / 10))}
            className="p-0.5 hover:text-slate-900"
            title="Zoom Out"
          >
            <ZoomOut className="w-3 h-3" />
          </button>
          <input
            type="range"
            min="0.5"
            max="1.5"
            step="0.05"
            value={zoom}
            onChange={(e) => onZoomChange(Number(e.target.value))}
            className="w-16 accent-[#107c41] h-1 bg-slate-300 rounded cursor-pointer"
          />
          <button
            onClick={() => onZoomChange(Math.min(1.5, Math.round((zoom + 0.1) * 10) / 10))}
            className="p-0.5 hover:text-slate-900"
            title="Zoom In"
          >
            <ZoomIn className="w-3 h-3" />
          </button>
          <span className="w-8 font-mono text-[10px] text-right font-medium text-slate-600">
            {Math.round(zoom * 100)}%
          </span>
        </div>
      </div>
    </footer>
  );
};
