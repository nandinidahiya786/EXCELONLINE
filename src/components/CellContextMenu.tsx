/**
 * @license
 * SPDX-License-Identifier: Apache-2.0
 */

import React from 'react';
import { 
  Scissors, 
  Copy, 
  ClipboardPaste, 
  Plus, 
  Trash2, 
  BarChart2, 
  Sparkles, 
  Eraser,
  SlidersHorizontal 
} from 'lucide-react';

interface CellContextMenuProps {
  x: number;
  y: number;
  onClose: () => void;
  onCut: () => void;
  onCopy: () => void;
  onPaste: () => void;
  onInsertRow: () => void;
  onInsertCol: () => void;
  onDeleteRow: () => void;
  onDeleteCol: () => void;
  onClearContent: () => void;
  onCreateChartFromSelection: () => void;
  onFormatCells: () => void;
}

export const CellContextMenu: React.FC<CellContextMenuProps> = ({
  x,
  y,
  onClose,
  onCut,
  onCopy,
  onPaste,
  onInsertRow,
  onInsertCol,
  onDeleteRow,
  onDeleteCol,
  onClearContent,
  onCreateChartFromSelection,
  onFormatCells,
}) => {
  return (
    <div
      className="fixed bg-white border border-slate-200 rounded-lg shadow-2xl py-1.5 z-50 text-xs text-slate-700 w-52 select-none animate-in fade-in zoom-in-95 duration-100"
      style={{ left: Math.min(x, window.innerWidth - 220), top: Math.min(y, window.innerHeight - 300) }}
      onClick={(e) => e.stopPropagation()}
    >
      <button
        onClick={() => { onCut(); onClose(); }}
        className="w-full text-left px-3 py-1.5 hover:bg-slate-100 flex items-center justify-between"
      >
        <span className="flex items-center gap-2"><Scissors className="w-3.5 h-3.5 text-slate-500" /> Cut</span>
        <span className="text-[10px] text-slate-400 font-mono">Ctrl+X</span>
      </button>
      <button
        onClick={() => { onCopy(); onClose(); }}
        className="w-full text-left px-3 py-1.5 hover:bg-slate-100 flex items-center justify-between"
      >
        <span className="flex items-center gap-2"><Copy className="w-3.5 h-3.5 text-slate-500" /> Copy</span>
        <span className="text-[10px] text-slate-400 font-mono">Ctrl+C</span>
      </button>
      <button
        onClick={() => { onPaste(); onClose(); }}
        className="w-full text-left px-3 py-1.5 hover:bg-slate-100 flex items-center justify-between"
      >
        <span className="flex items-center gap-2"><ClipboardPaste className="w-3.5 h-3.5 text-slate-500" /> Paste</span>
        <span className="text-[10px] text-slate-400 font-mono">Ctrl+V</span>
      </button>

      <div className="my-1 border-t border-slate-100" />

      <button
        onClick={() => { onInsertRow(); onClose(); }}
        className="w-full text-left px-3 py-1.5 hover:bg-slate-100 flex items-center gap-2"
      >
        <Plus className="w-3.5 h-3.5 text-emerald-600" /> Insert Row Above
      </button>
      <button
        onClick={() => { onInsertCol(); onClose(); }}
        className="w-full text-left px-3 py-1.5 hover:bg-slate-100 flex items-center gap-2"
      >
        <Plus className="w-3.5 h-3.5 text-emerald-600" /> Insert Column Left
      </button>
      <button
        onClick={() => { onDeleteRow(); onClose(); }}
        className="w-full text-left px-3 py-1.5 hover:bg-slate-100 flex items-center gap-2 text-rose-600"
      >
        <Trash2 className="w-3.5 h-3.5" /> Delete Row
      </button>
      <button
        onClick={() => { onDeleteCol(); onClose(); }}
        className="w-full text-left px-3 py-1.5 hover:bg-slate-100 flex items-center gap-2 text-rose-600"
      >
        <Trash2 className="w-3.5 h-3.5" /> Delete Column
      </button>

      <div className="my-1 border-t border-slate-100" />

      <button
        onClick={() => { onCreateChartFromSelection(); onClose(); }}
        className="w-full text-left px-3 py-1.5 hover:bg-slate-100 flex items-center gap-2 text-[#107c41] font-semibold"
      >
        <BarChart2 className="w-3.5 h-3.5" /> Create Chart from Selection
      </button>
      <button
        onClick={() => { onFormatCells(); onClose(); }}
        className="w-full text-left px-3 py-1.5 hover:bg-slate-100 flex items-center gap-2"
      >
        <SlidersHorizontal className="w-3.5 h-3.5 text-slate-500" /> Format Cells...
      </button>
      <button
        onClick={() => { onClearContent(); onClose(); }}
        className="w-full text-left px-3 py-1.5 hover:bg-slate-100 flex items-center gap-2 text-slate-600"
      >
        <Eraser className="w-3.5 h-3.5 text-slate-400" /> Clear Contents
      </button>
    </div>
  );
};
