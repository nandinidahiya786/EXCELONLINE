/**
 * @license
 * SPDX-License-Identifier: Apache-2.0
 */

import React, { useState } from 'react';
import { X, Check } from 'lucide-react';
import { CellFormat, NumberFormatType, TextAlign, VerticalAlign } from '../types/excel';

interface FormatCellsModalProps {
  isOpen: boolean;
  onClose: () => void;
  currentFormat?: CellFormat;
  onApplyFormat: (format: Partial<CellFormat>) => void;
}

export const FormatCellsModal: React.FC<FormatCellsModalProps> = ({
  isOpen,
  onClose,
  currentFormat = {},
  onApplyFormat,
}) => {
  if (!isOpen) return null;

  const [activeTab, setActiveTab] = useState<'number' | 'alignment' | 'font' | 'border' | 'fill'>('number');
  const [numFormat, setNumFormat] = useState<NumberFormatType>(currentFormat.numberFormat || 'general');
  const [decimals, setDecimals] = useState<number>(currentFormat.decimals ?? 2);
  const [currencySymbol, setCurrencySymbol] = useState<string>(currentFormat.currencySymbol || '$');
  const [align, setAlign] = useState<TextAlign>(currentFormat.align || 'left');
  const [verticalAlign, setVerticalAlign] = useState<VerticalAlign>(currentFormat.verticalAlign || 'middle');
  const [wrapText, setWrapText] = useState<boolean>(currentFormat.wrapText ?? false);
  const [bold, setBold] = useState<boolean>(currentFormat.bold ?? false);
  const [italic, setItalic] = useState<boolean>(currentFormat.italic ?? false);
  const [underline, setUnderline] = useState<boolean>(currentFormat.underline ?? false);
  const [strikethrough, setStrikethrough] = useState<boolean>(currentFormat.strikethrough ?? false);
  const [fontSize, setFontSize] = useState<number>(currentFormat.fontSize ?? 12);
  const [fontFamily, setFontFamily] = useState<string>(currentFormat.fontFamily || 'Plus Jakarta Sans');
  const [textColor, setTextColor] = useState<string>(currentFormat.textColor || '#0f172a');
  const [bgColor, setBgColor] = useState<string>(currentFormat.bgColor || '#ffffff');
  const [borderPreset, setBorderPreset] = useState<'none' | 'all' | 'bottom'>('none');

  const handleSubmit = (e: React.FormEvent) => {
    e.preventDefault();
    const updated: Partial<CellFormat> = {
      numberFormat: numFormat,
      decimals,
      currencySymbol,
      align,
      verticalAlign,
      wrapText,
      bold,
      italic,
      underline,
      strikethrough,
      fontSize,
      fontFamily,
      textColor,
      bgColor: bgColor === '#ffffff' ? undefined : bgColor,
    };

    if (borderPreset === 'all') {
      updated.borders = { top: '#cbd5e1', bottom: '#cbd5e1', left: '#cbd5e1', right: '#cbd5e1' };
    } else if (borderPreset === 'bottom') {
      updated.borders = { bottom: '#0f172a' };
    } else if (borderPreset === 'none') {
      updated.borders = undefined;
    }

    onApplyFormat(updated);
    onClose();
  };

  return (
    <div className="fixed inset-0 bg-slate-900/40 backdrop-blur-xs z-50 flex items-center justify-center p-4 select-none">
      <div className="bg-white border border-slate-200 rounded-xl shadow-2xl w-full max-w-lg overflow-hidden flex flex-col">
        {/* Header */}
        <div className="px-5 py-3 border-b border-slate-200 flex items-center justify-between bg-slate-50">
          <h2 className="text-sm font-bold text-slate-800">Format Cells</h2>
          <button onClick={onClose} className="p-1 text-slate-400 hover:text-slate-600 rounded">
            <X className="w-4 h-4" />
          </button>
        </div>

        {/* Tab Navigator */}
        <div className="flex border-b border-slate-200 bg-slate-50 text-xs font-semibold px-4 gap-2">
          {(['number', 'alignment', 'font', 'border', 'fill'] as const).map((tab) => (
            <button
              key={tab}
              onClick={() => setActiveTab(tab)}
              className={`py-2 px-2.5 capitalize transition-colors ${
                activeTab === tab ? 'text-[#107c41] border-b-2 border-[#107c41]' : 'text-slate-500 hover:text-slate-800'
              }`}
            >
              {tab}
            </button>
          ))}
        </div>

        {/* Form Body */}
        <form onSubmit={handleSubmit} className="p-5 space-y-4 text-xs">
          {/* NUMBER TAB */}
          {activeTab === 'number' && (
            <div className="grid grid-cols-2 gap-4">
              <div>
                <label className="block font-semibold text-slate-700 mb-1.5">Category</label>
                <div className="border border-slate-200 rounded-lg overflow-hidden max-h-48 overflow-y-auto">
                  {(['general', 'number', 'currency', 'accounting', 'percent', 'date', 'scientific', 'text'] as const).map((cat) => (
                    <button
                      key={cat}
                      type="button"
                      onClick={() => setNumFormat(cat)}
                      className={`w-full text-left px-3 py-1.5 capitalize ${
                        numFormat === cat ? 'bg-emerald-50 text-[#107c41] font-semibold' : 'hover:bg-slate-50 text-slate-700'
                      }`}
                    >
                      {cat}
                    </button>
                  ))}
                </div>
              </div>

              <div className="space-y-3">
                {numFormat !== 'general' && numFormat !== 'text' && numFormat !== 'date' && (
                  <div>
                    <label className="block font-semibold text-slate-700 mb-1">Decimal Places</label>
                    <input
                      type="number"
                      min="0"
                      max="6"
                      value={decimals}
                      onChange={(e) => setDecimals(Number(e.target.value))}
                      className="w-full px-2 py-1.5 border border-slate-300 rounded"
                    />
                  </div>
                )}

                {(numFormat === 'currency' || numFormat === 'accounting') && (
                  <div>
                    <label className="block font-semibold text-slate-700 mb-1">Currency Symbol</label>
                    <select
                      value={currencySymbol}
                      onChange={(e) => setCurrencySymbol(e.target.value)}
                      className="w-full px-2 py-1.5 border border-slate-300 rounded"
                    >
                      <option value="$">$ (USD)</option>
                      <option value="€">€ (EUR)</option>
                      <option value="£">£ (GBP)</option>
                      <option value="¥">¥ (JPY/CNY)</option>
                      <option value="₹">₹ (INR)</option>
                    </select>
                  </div>
                )}
              </div>
            </div>
          )}

          {/* ALIGNMENT TAB */}
          {activeTab === 'alignment' && (
            <div className="space-y-3">
              <div className="grid grid-cols-2 gap-3">
                <div>
                  <label className="block font-semibold text-slate-700 mb-1">Horizontal</label>
                  <select
                    value={align}
                    onChange={(e) => setAlign(e.target.value as any)}
                    className="w-full px-2 py-1.5 border border-slate-300 rounded"
                  >
                    <option value="left">Left</option>
                    <option value="center">Center</option>
                    <option value="right">Right</option>
                  </select>
                </div>
                <div>
                  <label className="block font-semibold text-slate-700 mb-1">Vertical</label>
                  <select
                    value={verticalAlign}
                    onChange={(e) => setVerticalAlign(e.target.value as any)}
                    className="w-full px-2 py-1.5 border border-slate-300 rounded"
                  >
                    <option value="top">Top</option>
                    <option value="middle">Middle</option>
                    <option value="bottom">Bottom</option>
                  </select>
                </div>
              </div>

              <label className="flex items-center gap-2 cursor-pointer pt-2">
                <input
                  type="checkbox"
                  checked={wrapText}
                  onChange={(e) => setWrapText(e.target.checked)}
                  className="rounded text-[#107c41]"
                />
                <span>Wrap Text</span>
              </label>
            </div>
          )}

          {/* FONT TAB */}
          {activeTab === 'font' && (
            <div className="space-y-3">
              <div className="grid grid-cols-2 gap-3">
                <div>
                  <label className="block font-semibold text-slate-700 mb-1">Font Family</label>
                  <select
                    value={fontFamily}
                    onChange={(e) => setFontFamily(e.target.value)}
                    className="w-full px-2 py-1.5 border border-slate-300 rounded"
                  >
                    <option value="Plus Jakarta Sans">Jakarta Sans</option>
                    <option value="JetBrains Mono">JetBrains Mono</option>
                    <option value="Arial">Arial</option>
                    <option value="Calibri">Calibri</option>
                    <option value="Times New Roman">Times New Roman</option>
                  </select>
                </div>
                <div>
                  <label className="block font-semibold text-slate-700 mb-1">Font Size</label>
                  <select
                    value={fontSize}
                    onChange={(e) => setFontSize(Number(e.target.value))}
                    className="w-full px-2 py-1.5 border border-slate-300 rounded"
                  >
                    {[9, 10, 11, 12, 14, 16, 18, 20, 24].map((sz) => (
                      <option key={sz} value={sz}>{sz} pt</option>
                    ))}
                  </select>
                </div>
              </div>

              <div className="flex gap-4 pt-1">
                <label className="flex items-center gap-1.5 cursor-pointer">
                  <input type="checkbox" checked={bold} onChange={(e) => setBold(e.target.checked)} className="rounded text-[#107c41]" />
                  <span className="font-bold">Bold</span>
                </label>
                <label className="flex items-center gap-1.5 cursor-pointer">
                  <input type="checkbox" checked={italic} onChange={(e) => setItalic(e.target.checked)} className="rounded text-[#107c41]" />
                  <span className="italic">Italic</span>
                </label>
                <label className="flex items-center gap-1.5 cursor-pointer">
                  <input type="checkbox" checked={underline} onChange={(e) => setUnderline(e.target.checked)} className="rounded text-[#107c41]" />
                  <span className="underline">Underline</span>
                </label>
              </div>
            </div>
          )}

          {/* BORDER TAB */}
          {activeTab === 'border' && (
            <div className="space-y-3">
              <label className="block font-semibold text-slate-700 mb-1">Presets</label>
              <div className="grid grid-cols-3 gap-2">
                {[
                  { id: 'none', label: 'None' },
                  { id: 'all', label: 'All Borders' },
                  { id: 'bottom', label: 'Bottom Only' },
                ].map((p) => (
                  <button
                    key={p.id}
                    type="button"
                    onClick={() => setBorderPreset(p.id as any)}
                    className={`py-2 px-3 rounded-lg border text-center transition-colors ${
                      borderPreset === p.id ? 'border-[#107c41] bg-emerald-50 text-[#107c41] font-bold' : 'border-slate-200 hover:bg-slate-50'
                    }`}
                  >
                    {p.label}
                  </button>
                ))}
              </div>
            </div>
          )}

          {/* FILL TAB */}
          {activeTab === 'fill' && (
            <div className="space-y-2">
              <label className="block font-semibold text-slate-700 mb-1">Background Color</label>
              <div className="grid grid-cols-6 gap-2">
                {['#ffffff', '#f8fafc', '#fee2e2', '#fef3c7', '#dcfce7', '#e0f2fe', '#ede9fe', '#f1f5f9'].map((c) => (
                  <button
                    key={c}
                    type="button"
                    onClick={() => setBgColor(c)}
                    style={{ backgroundColor: c }}
                    className={`h-8 rounded-lg border transition-all ${bgColor === c ? 'ring-2 ring-[#107c41] scale-105' : 'border-slate-200 hover:scale-105'}`}
                  />
                ))}
              </div>
            </div>
          )}

          {/* Footer */}
          <div className="pt-4 border-t border-slate-200 flex justify-end gap-2">
            <button
              type="button"
              onClick={onClose}
              className="px-4 py-2 text-slate-600 hover:bg-slate-100 rounded-lg transition-colors"
            >
              Cancel
            </button>
            <button
              type="submit"
              className="px-4 py-2 font-semibold text-white bg-[#107c41] hover:bg-[#0c6233] rounded-lg shadow-sm transition-colors flex items-center gap-1.5"
            >
              <Check className="w-3.5 h-3.5" />
              <span>Apply Formatting</span>
            </button>
          </div>
        </form>
      </div>
    </div>
  );
};
