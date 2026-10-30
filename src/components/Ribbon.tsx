/**
 * @license
 * SPDX-License-Identifier: Apache-2.0
 */

import React, { useState } from 'react';
import {
  Bold,
  Italic,
  Underline,
  Strikethrough,
  AlignLeft,
  AlignCenter,
  AlignRight,
  AlignJustify,
  DollarSign,
  Percent,
  Calculator,
  Filter,
  ArrowDownAZ,
  ArrowUpZA,
  Search,
  Grid3X3,
  PaintBucket,
  Type,
  Maximize2,
  WrapText,
  BarChart2,
  PieChart,
  LineChart,
  TrendingUp,
  Table,
  Plus,
  Trash2,
  Sparkles,
  Layers,
  HelpCircle,
  Undo2,
  Redo2,
  Copy,
  Scissors,
  ClipboardPaste,
  ChevronDown
} from 'lucide-react';
import { CellFormat, NumberFormatType, TextAlign } from '../types/excel';

interface RibbonProps {
  activeTab: 'home' | 'insert' | 'formulas' | 'data' | 'dashboard' | 'view';
  onTabChange: (tab: 'home' | 'insert' | 'formulas' | 'data' | 'dashboard' | 'view') => void;
  activeCellFormat?: CellFormat;
  onApplyFormat: (format: Partial<CellFormat>) => void;
  onAutoSum: (funcName: 'SUM' | 'AVERAGE' | 'COUNT' | 'MAX' | 'MIN') => void;
  onSort: (direction: 'asc' | 'desc') => void;
  onToggleFilter: () => void;
  onOpenInsertFunction: () => void;
  onOpenFindReplace: () => void;
  onOpenDataValidation: () => void;
  onOpenNewChart: (type?: any) => void;
  onInsertRow: () => void;
  onInsertCol: () => void;
  onDeleteRow: () => void;
  onDeleteCol: () => void;
  onClearContent: (type: 'all' | 'content' | 'format') => void;
  onMergeCenter: () => void;
  onRecalculate: () => void;
  onUndo: () => void;
  onRedo: () => void;
  canUndo: boolean;
  canRedo: boolean;
  onCopy: () => void;
  onCut: () => void;
  onPaste: () => void;
  showGridlines: boolean;
  onToggleGridlines: () => void;
  showFormulaBar: boolean;
  onToggleFormulaBar: () => void;
  zoom: number;
  onZoomChange: (z: number) => void;
  onApplyConditionalScale: () => void;
  onApplyConditionalRule: (type: 'greaterThan' | 'lessThan' | 'equalTo', val: number) => void;
  onFreezeTopRow: () => void;
  onFreezeFirstCol: () => void;
  onUnfreeze: () => void;
}

export const Ribbon: React.FC<RibbonProps> = ({
  activeTab,
  onTabChange,
  activeCellFormat = {},
  onApplyFormat,
  onAutoSum,
  onSort,
  onToggleFilter,
  onOpenInsertFunction,
  onOpenFindReplace,
  onOpenDataValidation,
  onOpenNewChart,
  onInsertRow,
  onInsertCol,
  onDeleteRow,
  onDeleteCol,
  onClearContent,
  onMergeCenter,
  onRecalculate,
  onUndo,
  onRedo,
  canUndo,
  canRedo,
  onCopy,
  onCut,
  onPaste,
  showGridlines,
  onToggleGridlines,
  showFormulaBar,
  onToggleFormulaBar,
  zoom,
  onZoomChange,
  onApplyConditionalScale,
  onApplyConditionalRule,
  onFreezeTopRow,
  onFreezeFirstCol,
  onUnfreeze,
}) => {
  const [showColorPicker, setShowColorPicker] = useState<'fill' | 'text' | null>(null);
  const [showAutoSumMenu, setShowAutoSumMenu] = useState(false);
  const [showCondMenu, setShowCondMenu] = useState(false);
  const [showCurrencyMenu, setShowCurrencyMenu] = useState(false);
  const [showClearMenu, setShowClearMenu] = useState(false);
  const [showBordersMenu, setShowBordersMenu] = useState(false);

  const colors = [
    '#ffffff', '#f8fafc', '#f1f5f9', '#e2e8f0', '#cbd5e1', '#64748b', '#334155', '#0f172a',
    '#fee2e2', '#fecaca', '#f87171', '#ef4444', '#dc2626', '#b91c1c', '#991b1b', '#7f1d1d',
    '#ffedd5', '#fed7aa', '#fb923c', '#f97316', '#ea580c', '#c2410c', '#9a3412', '#7c2d12',
    '#fef3c7', '#fde68a', '#facc15', '#eab308', '#ca8a04', '#a16207', '#854d0e', '#713f12',
    '#dcfce7', '#bbf7d0', '#86efac', '#4ade80', '#22c55e', '#16a34a', '#15803d', '#107c41',
    '#e0f2fe', '#bae6fd', '#7dd3fc', '#38bdf8', '#0ea5e9', '#0284c7', '#0369a1', '#075985',
    '#ede9fe', '#ddd6fe', '#c4b5fd', '#a78bfa', '#8b5cf6', '#7c3aed', '#6d28d9', '#5b21b6',
  ];

  return (
    <div className="bg-white border-b border-slate-200 select-none shadow-xs">
      {/* Ribbon Navigation Tabs */}
      <div className="flex items-center px-2 bg-slate-50 border-b border-slate-200 text-xs font-medium text-slate-600 gap-1">
        {(['home', 'insert', 'formulas', 'data', 'dashboard', 'view'] as const).map((tab) => {
          const labels: Record<string, string> = {
            home: 'Home',
            insert: 'Insert',
            formulas: 'Formulas',
            data: 'Data',
            dashboard: 'Dashboard Studio',
            view: 'View',
          };
          const isActive = activeTab === tab;
          return (
            <button
              key={tab}
              onClick={() => onTabChange(tab)}
              className={`px-3 py-1.5 transition-colors relative font-semibold ${
                isActive
                  ? 'text-[#107c41] bg-white border-b-2 border-[#107c41]'
                  : 'text-slate-600 hover:text-slate-900 hover:bg-slate-100'
              }`}
            >
              {labels[tab]}
            </button>
          );
        })}
      </div>

      {/* Ribbon Content Panel */}
      <div className="px-3 py-2 flex items-center gap-3 overflow-x-auto min-h-[58px]">
        {/* ===================== HOME TAB ===================== */}
        {activeTab === 'home' && (
          <>
            {/* Undo / Redo / Clipboard Group */}
            <div className="flex items-center gap-1 border-r border-slate-200 pr-3">
              <button
                onClick={onUndo}
                disabled={!canUndo}
                className="p-1.5 rounded hover:bg-slate-100 disabled:opacity-30 disabled:hover:bg-transparent text-slate-700"
                title="Undo (Ctrl+Z)"
              >
                <Undo2 className="w-4 h-4" />
              </button>
              <button
                onClick={onRedo}
                disabled={!canRedo}
                className="p-1.5 rounded hover:bg-slate-100 disabled:opacity-30 disabled:hover:bg-transparent text-slate-700"
                title="Redo (Ctrl+Y)"
              >
                <Redo2 className="w-4 h-4" />
              </button>
              <div className="h-5 w-[1px] bg-slate-200 mx-0.5" />
              <button
                onClick={onPaste}
                className="p-1.5 rounded hover:bg-slate-100 text-slate-700"
                title="Paste (Ctrl+V)"
              >
                <ClipboardPaste className="w-4 h-4" />
              </button>
              <button
                onClick={onCopy}
                className="p-1.5 rounded hover:bg-slate-100 text-slate-700"
                title="Copy (Ctrl+C)"
              >
                <Copy className="w-4 h-4" />
              </button>
              <button
                onClick={onCut}
                className="p-1.5 rounded hover:bg-slate-100 text-slate-700"
                title="Cut (Ctrl+X)"
              >
                <Scissors className="w-4 h-4" />
              </button>
            </div>

            {/* Font Styling Group */}
            <div className="flex items-center gap-1 border-r border-slate-200 pr-3">
              {/* Font Family */}
              <select
                value={activeCellFormat.fontFamily || 'Plus Jakarta Sans'}
                onChange={(e) => onApplyFormat({ fontFamily: e.target.value })}
                className="text-xs bg-slate-50 border border-slate-200 rounded px-1.5 py-1 text-slate-800 outline-none w-28"
              >
                <option value="Plus Jakarta Sans">Jakarta Sans</option>
                <option value="JetBrains Mono">JetBrains Mono</option>
                <option value="Arial">Arial</option>
                <option value="Calibri">Calibri</option>
                <option value="Times New Roman">Times New Roman</option>
                <option value="Georgia">Georgia</option>
              </select>

              {/* Font Size */}
              <select
                value={activeCellFormat.fontSize || 12}
                onChange={(e) => onApplyFormat({ fontSize: Number(e.target.value) })}
                className="text-xs bg-slate-50 border border-slate-200 rounded px-1 py-1 text-slate-800 outline-none w-14"
              >
                {[9, 10, 11, 12, 14, 16, 18, 20, 24].map((sz) => (
                  <option key={sz} value={sz}>{sz} pt</option>
                ))}
              </select>

              {/* B, I, U, S */}
              <button
                onClick={() => onApplyFormat({ bold: !activeCellFormat.bold })}
                className={`p-1.5 rounded transition-colors ${activeCellFormat.bold ? 'bg-[#107c41] text-white' : 'hover:bg-slate-100 text-slate-700'}`}
                title="Bold (Ctrl+B)"
              >
                <Bold className="w-3.5 h-3.5" />
              </button>
              <button
                onClick={() => onApplyFormat({ italic: !activeCellFormat.italic })}
                className={`p-1.5 rounded transition-colors ${activeCellFormat.italic ? 'bg-[#107c41] text-white' : 'hover:bg-slate-100 text-slate-700'}`}
                title="Italic (Ctrl+I)"
              >
                <Italic className="w-3.5 h-3.5" />
              </button>
              <button
                onClick={() => onApplyFormat({ underline: !activeCellFormat.underline })}
                className={`p-1.5 rounded transition-colors ${activeCellFormat.underline ? 'bg-[#107c41] text-white' : 'hover:bg-slate-100 text-slate-700'}`}
                title="Underline (Ctrl+U)"
              >
                <Underline className="w-3.5 h-3.5" />
              </button>
              <button
                onClick={() => onApplyFormat({ strikethrough: !activeCellFormat.strikethrough })}
                className={`p-1.5 rounded transition-colors ${activeCellFormat.strikethrough ? 'bg-[#107c41] text-white' : 'hover:bg-slate-100 text-slate-700'}`}
                title="Strikethrough"
              >
                <Strikethrough className="w-3.5 h-3.5" />
              </button>

              {/* Fill Color Picker */}
              <div className="relative">
                <button
                  onClick={() => setShowColorPicker(showColorPicker === 'fill' ? null : 'fill')}
                  className="p-1.5 rounded hover:bg-slate-100 text-slate-700 flex items-center gap-0.5"
                  title="Fill Color"
                >
                  <PaintBucket className="w-3.5 h-3.5" />
                  <div
                    className="w-3 h-1 border border-slate-300"
                    style={{ backgroundColor: activeCellFormat.bgColor || '#ffffff' }}
                  />
                </button>
                {showColorPicker === 'fill' && (
                  <div className="absolute top-8 left-0 p-2 bg-white border border-slate-200 rounded-lg shadow-xl grid grid-cols-8 gap-1 z-50 w-52">
                    {colors.map((c) => (
                      <button
                        key={c}
                        onClick={() => {
                          onApplyFormat({ bgColor: c });
                          setShowColorPicker(null);
                        }}
                        style={{ backgroundColor: c }}
                        className="w-5 h-5 rounded-xs border border-slate-200 hover:scale-110 transition-transform"
                      />
                    ))}
                    <button
                      onClick={() => {
                        onApplyFormat({ bgColor: undefined });
                        setShowColorPicker(null);
                      }}
                      className="col-span-8 text-[11px] py-1 text-slate-600 hover:bg-slate-100 rounded text-center mt-1"
                    >
                      No Fill
                    </button>
                  </div>
                )}
              </div>

              {/* Text Color Picker */}
              <div className="relative">
                <button
                  onClick={() => setShowColorPicker(showColorPicker === 'text' ? null : 'text')}
                  className="p-1.5 rounded hover:bg-slate-100 text-slate-700 flex items-center gap-0.5"
                  title="Font Color"
                >
                  <Type className="w-3.5 h-3.5" />
                  <div
                    className="w-3 h-1"
                    style={{ backgroundColor: activeCellFormat.textColor || '#0f172a' }}
                  />
                </button>
                {showColorPicker === 'text' && (
                  <div className="absolute top-8 left-0 p-2 bg-white border border-slate-200 rounded-lg shadow-xl grid grid-cols-8 gap-1 z-50 w-52">
                    {colors.map((c) => (
                      <button
                        key={c}
                        onClick={() => {
                          onApplyFormat({ textColor: c });
                          setShowColorPicker(null);
                        }}
                        style={{ backgroundColor: c }}
                        className="w-5 h-5 rounded-xs border border-slate-200 hover:scale-110 transition-transform"
                      />
                    ))}
                  </div>
                )}
              </div>

              {/* Borders */}
              <div className="relative">
                <button
                  onClick={() => setShowBordersMenu(!showBordersMenu)}
                  className="p-1.5 rounded hover:bg-slate-100 text-slate-700"
                  title="Borders"
                >
                  <Grid3X3 className="w-3.5 h-3.5" />
                </button>
                {showBordersMenu && (
                  <div className="absolute top-8 left-0 bg-white border border-slate-200 rounded shadow-xl py-1 z-50 w-36 text-xs text-slate-700">
                    <button
                      onClick={() => {
                        onApplyFormat({ borders: { top: '#cbd5e1', bottom: '#cbd5e1', left: '#cbd5e1', right: '#cbd5e1' } });
                        setShowBordersMenu(false);
                      }}
                      className="w-full text-left px-3 py-1.5 hover:bg-slate-100"
                    >
                      All Borders
                    </button>
                    <button
                      onClick={() => {
                        onApplyFormat({ borders: { bottom: '#0f172a' } });
                        setShowBordersMenu(false);
                      }}
                      className="w-full text-left px-3 py-1.5 hover:bg-slate-100"
                    >
                      Bottom Border
                    </button>
                    <button
                      onClick={() => {
                        onApplyFormat({ borders: { bottom: '2px solid #0f172a' } });
                        setShowBordersMenu(false);
                      }}
                      className="w-full text-left px-3 py-1.5 hover:bg-slate-100"
                    >
                      Thick Bottom
                    </button>
                    <button
                      onClick={() => {
                        onApplyFormat({ borders: undefined });
                        setShowBordersMenu(false);
                      }}
                      className="w-full text-left px-3 py-1.5 hover:bg-slate-100 text-red-600"
                    >
                      No Borders
                    </button>
                  </div>
                )}
              </div>
            </div>

            {/* Alignment Group */}
            <div className="flex items-center gap-1 border-r border-slate-200 pr-3">
              <button
                onClick={() => onApplyFormat({ align: 'left' })}
                className={`p-1.5 rounded transition-colors ${activeCellFormat.align === 'left' ? 'bg-[#107c41] text-white' : 'hover:bg-slate-100 text-slate-700'}`}
                title="Align Left"
              >
                <AlignLeft className="w-3.5 h-3.5" />
              </button>
              <button
                onClick={() => onApplyFormat({ align: 'center' })}
                className={`p-1.5 rounded transition-colors ${activeCellFormat.align === 'center' ? 'bg-[#107c41] text-white' : 'hover:bg-slate-100 text-slate-700'}`}
                title="Align Center"
              >
                <AlignCenter className="w-3.5 h-3.5" />
              </button>
              <button
                onClick={() => onApplyFormat({ align: 'right' })}
                className={`p-1.5 rounded transition-colors ${activeCellFormat.align === 'right' ? 'bg-[#107c41] text-white' : 'hover:bg-slate-100 text-slate-700'}`}
                title="Align Right"
              >
                <AlignRight className="w-3.5 h-3.5" />
              </button>
              <button
                onClick={() => onApplyFormat({ wrapText: !activeCellFormat.wrapText })}
                className={`p-1.5 rounded transition-colors ${activeCellFormat.wrapText ? 'bg-[#107c41] text-white' : 'hover:bg-slate-100 text-slate-700'}`}
                title="Wrap Text"
              >
                <WrapText className="w-3.5 h-3.5" />
              </button>
              <button
                onClick={onMergeCenter}
                className="px-2 py-1 text-xs font-medium text-slate-700 hover:bg-slate-100 rounded border border-slate-200 transition-colors"
                title="Merge & Center Selection"
              >
                Merge & Center
              </button>
            </div>

            {/* Number Formatting Group */}
            <div className="flex items-center gap-1 border-r border-slate-200 pr-3">
              <select
                value={activeCellFormat.numberFormat || 'general'}
                onChange={(e) => onApplyFormat({ numberFormat: e.target.value as NumberFormatType })}
                className="text-xs bg-slate-50 border border-slate-200 rounded px-1.5 py-1 text-slate-800 outline-none w-24"
              >
                <option value="general">General</option>
                <option value="number">Number</option>
                <option value="currency">Currency</option>
                <option value="accounting">Accounting</option>
                <option value="percent">Percentage</option>
                <option value="date">Date</option>
                <option value="scientific">Scientific</option>
                <option value="text">Plain Text</option>
              </select>

              {/* Currency symbol */}
              <div className="relative">
                <button
                  onClick={() => setShowCurrencyMenu(!showCurrencyMenu)}
                  className="p-1.5 rounded hover:bg-slate-100 text-slate-700 flex items-center"
                  title="Currency Format"
                >
                  <DollarSign className="w-3.5 h-3.5" />
                  <ChevronDown className="w-2.5 h-2.5 ml-0.5 text-slate-400" />
                </button>
                {showCurrencyMenu && (
                  <div className="absolute top-8 left-0 bg-white border border-slate-200 rounded shadow-xl py-1 z-50 w-28 text-xs text-slate-700">
                    {['$', '€', '£', '¥', '₹'].map((sym) => (
                      <button
                        key={sym}
                        onClick={() => {
                          onApplyFormat({ numberFormat: 'currency', currencySymbol: sym });
                          setShowCurrencyMenu(false);
                        }}
                        className="w-full text-left px-3 py-1 hover:bg-slate-100 font-mono"
                      >
                        {sym} Currency
                      </button>
                    ))}
                  </div>
                )}
              </div>

              {/* Percent */}
              <button
                onClick={() => onApplyFormat({ numberFormat: 'percent', decimals: 1 })}
                className="p-1.5 rounded hover:bg-slate-100 text-slate-700"
                title="Percent Format (%)"
              >
                <Percent className="w-3.5 h-3.5" />
              </button>

              {/* Decimals +/- */}
              <button
                onClick={() => onApplyFormat({ decimals: Math.min(6, (activeCellFormat.decimals || 0) + 1) })}
                className="px-1.5 py-1 text-[11px] font-mono font-semibold text-slate-700 hover:bg-slate-100 rounded border border-slate-200"
                title="Increase Decimal"
              >
                .00+
              </button>
              <button
                onClick={() => onApplyFormat({ decimals: Math.max(0, (activeCellFormat.decimals || 2) - 1) })}
                className="px-1.5 py-1 text-[11px] font-mono font-semibold text-slate-700 hover:bg-slate-100 rounded border border-slate-200"
                title="Decrease Decimal"
              >
                .0-
              </button>
            </div>

            {/* Styles & Conditional Formatting */}
            <div className="flex items-center gap-1 border-r border-slate-200 pr-3 relative">
              <button
                onClick={() => setShowCondMenu(!showCondMenu)}
                className="flex items-center gap-1 px-2.5 py-1 text-xs font-medium text-slate-700 hover:bg-slate-100 rounded border border-slate-200 transition-colors"
                title="Conditional Formatting"
              >
                <Sparkles className="w-3.5 h-3.5 text-amber-500" />
                <span>Conditional</span>
                <ChevronDown className="w-3 h-3 text-slate-400" />
              </button>

              {showCondMenu && (
                <div className="absolute top-9 left-0 bg-white border border-slate-200 rounded-lg shadow-xl py-1.5 z-50 w-52 text-xs text-slate-700">
                  <div className="px-3 py-1 text-[10px] uppercase font-bold text-slate-400">Color Scales</div>
                  <button
                    onClick={() => {
                      onApplyConditionalScale();
                      setShowCondMenu(false);
                    }}
                    className="w-full text-left px-3 py-1.5 hover:bg-slate-50 flex items-center justify-between"
                  >
                    <span>Green-Yellow-Red Scale</span>
                    <div className="w-12 h-3 rounded bg-gradient-to-r from-emerald-300 via-amber-200 to-red-300" />
                  </button>
                  <div className="my-1 border-t border-slate-100" />
                  <div className="px-3 py-1 text-[10px] uppercase font-bold text-slate-400">Highlight Rules</div>
                  <button
                    onClick={() => {
                      const val = prompt('Highlight cells greater than:', '100000');
                      if (val) onApplyConditionalRule('greaterThan', Number(val));
                      setShowCondMenu(false);
                    }}
                    className="w-full text-left px-3 py-1.5 hover:bg-slate-50"
                  >
                    Greater Than...
                  </button>
                  <button
                    onClick={() => {
                      const val = prompt('Highlight cells less than:', '50000');
                      if (val) onApplyConditionalRule('lessThan', Number(val));
                      setShowCondMenu(false);
                    }}
                    className="w-full text-left px-3 py-1.5 hover:bg-slate-50"
                  >
                    Less Than...
                  </button>
                </div>
              )}
            </div>

            {/* Cells & Editing Operations */}
            <div className="flex items-center gap-1">
              <button
                onClick={onInsertRow}
                className="px-2 py-1 text-xs font-medium text-slate-700 hover:bg-slate-100 rounded border border-slate-200"
                title="Insert Row Above"
              >
                + Row
              </button>
              <button
                onClick={onInsertCol}
                className="px-2 py-1 text-xs font-medium text-slate-700 hover:bg-slate-100 rounded border border-slate-200"
                title="Insert Column Left"
              >
                + Col
              </button>
              <button
                onClick={onDeleteRow}
                className="p-1.5 rounded hover:bg-red-50 text-slate-700 hover:text-red-600"
                title="Delete Row"
              >
                <Trash2 className="w-3.5 h-3.5" />
              </button>

              {/* AutoSum */}
              <div className="relative ml-1">
                <button
                  onClick={() => setShowAutoSumMenu(!showAutoSumMenu)}
                  className="flex items-center gap-1 px-2.5 py-1 text-xs font-medium bg-[#107c41]/10 text-[#107c41] hover:bg-[#107c41]/20 rounded border border-[#107c41]/30 transition-colors"
                >
                  <Calculator className="w-3.5 h-3.5" />
                  <span>AutoSum</span>
                  <ChevronDown className="w-3 h-3" />
                </button>
                {showAutoSumMenu && (
                  <div className="absolute top-8 left-0 bg-white border border-slate-200 rounded shadow-xl py-1 z-50 w-32 text-xs text-slate-700">
                    {(['SUM', 'AVERAGE', 'COUNT', 'MAX', 'MIN'] as const).map((fn) => (
                      <button
                        key={fn}
                        onClick={() => {
                          onAutoSum(fn);
                          setShowAutoSumMenu(false);
                        }}
                        className="w-full text-left px-3 py-1.5 hover:bg-slate-100 font-mono"
                      >
                        ={fn}()
                      </button>
                    ))}
                  </div>
                )}
              </div>

              {/* Sort & Filter */}
              <button
                onClick={() => onSort('asc')}
                className="p-1.5 rounded hover:bg-slate-100 text-slate-700"
                title="Sort Ascending A-Z"
              >
                <ArrowDownAZ className="w-3.5 h-3.5" />
              </button>
              <button
                onClick={() => onSort('desc')}
                className="p-1.5 rounded hover:bg-slate-100 text-slate-700"
                title="Sort Descending Z-A"
              >
                <ArrowUpZA className="w-3.5 h-3.5" />
              </button>
              <button
                onClick={onToggleFilter}
                className="p-1.5 rounded hover:bg-slate-100 text-slate-700"
                title="Toggle Filter Headers"
              >
                <Filter className="w-3.5 h-3.5" />
              </button>

              {/* Find & Replace */}
              <button
                onClick={onOpenFindReplace}
                className="p-1.5 rounded hover:bg-slate-100 text-slate-700"
                title="Find and Replace (Ctrl+F)"
              >
                <Search className="w-3.5 h-3.5" />
              </button>
            </div>
          </>
        )}

        {/* ===================== INSERT TAB ===================== */}
        {activeTab === 'insert' && (
          <div className="flex items-center gap-3">
            <div className="flex items-center gap-1.5 border-r border-slate-200 pr-3">
              <span className="text-[11px] font-bold text-slate-400 uppercase tracking-wider mr-1">Charts:</span>
              <button
                onClick={() => onOpenNewChart('column')}
                className="flex items-center gap-1.5 px-2.5 py-1.5 text-xs font-medium text-slate-700 bg-white hover:bg-slate-100 rounded border border-slate-200"
              >
                <BarChart2 className="w-4 h-4 text-blue-600" />
                <span>Column</span>
              </button>
              <button
                onClick={() => onOpenNewChart('bar')}
                className="flex items-center gap-1.5 px-2.5 py-1.5 text-xs font-medium text-slate-700 bg-white hover:bg-slate-100 rounded border border-slate-200"
              >
                <BarChart2 className="w-4 h-4 text-emerald-600 rotate-90" />
                <span>Bar</span>
              </button>
              <button
                onClick={() => onOpenNewChart('line')}
                className="flex items-center gap-1.5 px-2.5 py-1.5 text-xs font-medium text-slate-700 bg-white hover:bg-slate-100 rounded border border-slate-200"
              >
                <LineChart className="w-4 h-4 text-amber-500" />
                <span>Line</span>
              </button>
              <button
                onClick={() => onOpenNewChart('area')}
                className="flex items-center gap-1.5 px-2.5 py-1.5 text-xs font-medium text-slate-700 bg-white hover:bg-slate-100 rounded border border-slate-200"
              >
                <TrendingUp className="w-4 h-4 text-indigo-600" />
                <span>Area</span>
              </button>
              <button
                onClick={() => onOpenNewChart('pie')}
                className="flex items-center gap-1.5 px-2.5 py-1.5 text-xs font-medium text-slate-700 bg-white hover:bg-slate-100 rounded border border-slate-200"
              >
                <PieChart className="w-4 h-4 text-rose-500" />
                <span>Pie / Donut</span>
              </button>
            </div>

            <div className="flex items-center gap-1.5">
              <button
                onClick={() => onOpenNewChart('kpi')}
                className="flex items-center gap-1.5 px-2.5 py-1.5 text-xs font-semibold text-emerald-800 bg-emerald-50 hover:bg-emerald-100 rounded border border-emerald-200"
              >
                <Sparkles className="w-3.5 h-3.5 text-emerald-600" />
                <span>KPI Metric Card</span>
              </button>
              <button
                onClick={onOpenInsertFunction}
                className="flex items-center gap-1.5 px-2.5 py-1.5 text-xs font-medium text-slate-700 bg-white hover:bg-slate-100 rounded border border-slate-200"
              >
                <span className="font-serif italic font-bold">fx</span>
                <span>Insert Function</span>
              </button>
            </div>
          </div>
        )}

        {/* ===================== FORMULAS TAB ===================== */}
        {activeTab === 'formulas' && (
          <div className="flex items-center gap-3">
            <button
              onClick={onOpenInsertFunction}
              className="flex items-center gap-1.5 px-3 py-1.5 text-xs font-bold text-white bg-[#107c41] hover:bg-[#0c6233] rounded shadow-xs"
            >
              <span className="font-serif italic text-sm">fx</span>
              <span>Insert Function</span>
            </button>

            <div className="flex items-center gap-1.5 border-l border-slate-200 pl-3">
              {(['SUM', 'AVERAGE', 'IF', 'VLOOKUP', 'COUNT', 'MIN', 'MAX', 'TODAY'] as const).map((fn) => (
                <button
                  key={fn}
                  onClick={() => onAutoSum(fn as any)}
                  className="px-2 py-1 text-xs font-mono font-semibold text-slate-700 hover:bg-slate-100 rounded border border-slate-200"
                >
                  ={fn}
                </button>
              ))}
            </div>

            <button
              onClick={onRecalculate}
              className="ml-auto flex items-center gap-1.5 px-2.5 py-1.5 text-xs font-medium text-slate-700 bg-white hover:bg-slate-100 rounded border border-slate-200"
              title="Force full recalculation of all formulas"
            >
              <Calculator className="w-3.5 h-3.5 text-slate-500" />
              <span>Calculate Now</span>
            </button>
          </div>
        )}

        {/* ===================== DATA TAB ===================== */}
        {activeTab === 'data' && (
          <div className="flex items-center gap-3">
            <div className="flex items-center gap-1.5 border-r border-slate-200 pr-3">
              <button
                onClick={() => onSort('asc')}
                className="flex items-center gap-1.5 px-2.5 py-1.5 text-xs font-medium text-slate-700 bg-white hover:bg-slate-100 rounded border border-slate-200"
              >
                <ArrowDownAZ className="w-4 h-4 text-emerald-600" />
                <span>Sort A to Z</span>
              </button>
              <button
                onClick={() => onSort('desc')}
                className="flex items-center gap-1.5 px-2.5 py-1.5 text-xs font-medium text-slate-700 bg-white hover:bg-slate-100 rounded border border-slate-200"
              >
                <ArrowUpZA className="w-4 h-4 text-emerald-600" />
                <span>Sort Z to A</span>
              </button>
              <button
                onClick={onToggleFilter}
                className="flex items-center gap-1.5 px-2.5 py-1.5 text-xs font-medium text-slate-700 bg-white hover:bg-slate-100 rounded border border-slate-200"
              >
                <Filter className="w-4 h-4 text-blue-600" />
                <span>Filter Headers</span>
              </button>
            </div>

            <div className="flex items-center gap-1.5">
              <button
                onClick={onOpenDataValidation}
                className="flex items-center gap-1.5 px-2.5 py-1.5 text-xs font-medium text-slate-700 bg-white hover:bg-slate-100 rounded border border-slate-200"
              >
                <Table className="w-4 h-4 text-purple-600" />
                <span>Data Validation (Dropdowns)</span>
              </button>
            </div>
          </div>
        )}

        {/* ===================== DASHBOARD STUDIO TAB ===================== */}
        {activeTab === 'dashboard' && (
          <div className="flex items-center gap-3">
            <button
              onClick={() => onOpenNewChart()}
              className="flex items-center gap-1.5 px-3 py-1.5 text-xs font-bold text-white bg-[#107c41] hover:bg-[#0c6233] rounded shadow-xs"
            >
              <Plus className="w-3.5 h-3.5" />
              <span>Create Chart Widget</span>
            </button>
            <button
              onClick={() => onOpenNewChart('kpi')}
              className="flex items-center gap-1.5 px-2.5 py-1.5 text-xs font-medium text-slate-700 bg-white hover:bg-slate-100 rounded border border-slate-200"
            >
              <Sparkles className="w-3.5 h-3.5 text-amber-500" />
              <span>Add KPI Scorecard</span>
            </button>
            <span className="text-xs text-slate-400">· Charts synchronize live with worksheet calculations</span>
          </div>
        )}

        {/* ===================== VIEW TAB ===================== */}
        {activeTab === 'view' && (
          <div className="flex items-center gap-3 text-xs text-slate-700">
            <label className="flex items-center gap-1.5 cursor-pointer">
              <input
                type="checkbox"
                checked={showGridlines}
                onChange={onToggleGridlines}
                className="rounded text-[#107c41] focus:ring-0"
              />
              <span>Gridlines</span>
            </label>
            <label className="flex items-center gap-1.5 cursor-pointer">
              <input
                type="checkbox"
                checked={showFormulaBar}
                onChange={onToggleFormulaBar}
                className="rounded text-[#107c41] focus:ring-0"
              />
              <span>Formula Bar</span>
            </label>

            <div className="h-5 w-[1px] bg-slate-200 mx-1" />

            {/* Freeze Panes */}
            <div className="flex items-center gap-1">
              <button
                onClick={onFreezeTopRow}
                className="px-2 py-1 bg-slate-50 hover:bg-slate-100 border border-slate-200 rounded text-xs"
              >
                Freeze Top Row
              </button>
              <button
                onClick={onFreezeFirstCol}
                className="px-2 py-1 bg-slate-50 hover:bg-slate-100 border border-slate-200 rounded text-xs"
              >
                Freeze First Column
              </button>
              <button
                onClick={onUnfreeze}
                className="px-2 py-1 bg-slate-50 hover:bg-slate-100 border border-slate-200 rounded text-xs text-slate-500"
              >
                Unfreeze Panes
              </button>
            </div>

            <div className="h-5 w-[1px] bg-slate-200 mx-1" />

            {/* Zoom */}
            <div className="flex items-center gap-1">
              <span>Zoom:</span>
              <select
                value={zoom}
                onChange={(e) => onZoomChange(Number(e.target.value))}
                className="bg-slate-50 border border-slate-200 rounded px-1.5 py-1 text-xs"
              >
                <option value={0.5}>50%</option>
                <option value={0.75}>75%</option>
                <option value={1}>100%</option>
                <option value={1.25}>125%</option>
                <option value={1.5}>150%</option>
              </select>
            </div>
          </div>
        )}
      </div>
    </div>
  );
};
