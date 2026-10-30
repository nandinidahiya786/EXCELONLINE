/**
 * @license
 * SPDX-License-Identifier: Apache-2.0
 */

import React, { useState, useEffect, useRef } from 'react';
import { X, Check } from 'lucide-react';
import { FUNCTION_CATALOG } from '../utils/formulaEngine';

interface FormulaBarProps {
  selectedCellId: string;
  selectedRangeStr?: string;
  cellRawValue: string;
  onRawValueChange: (val: string) => void;
  onCommit: () => void;
  onCancel: () => void;
  onOpenInsertFunction: () => void;
  onJumpToCell: (cellId: string) => void;
}

export const FormulaBar: React.FC<FormulaBarProps> = ({
  selectedCellId,
  selectedRangeStr,
  cellRawValue,
  onRawValueChange,
  onCommit,
  onCancel,
  onOpenInsertFunction,
  onJumpToCell,
}) => {
  const [inputValue, setInputValue] = useState(cellRawValue);
  const [nameBoxValue, setNameBoxValue] = useState(selectedRangeStr || selectedCellId);
  const [isEditingNameBox, setIsEditingNameBox] = useState(false);
  const [suggestedFunc, setSuggestedFunc] = useState<string | null>(null);
  const inputRef = useRef<HTMLInputElement>(null);

  useEffect(() => {
    setInputValue(cellRawValue);
  }, [cellRawValue, selectedCellId]);

  useEffect(() => {
    setNameBoxValue(selectedRangeStr || selectedCellId);
  }, [selectedCellId, selectedRangeStr]);

  const handleChange = (e: React.ChangeEvent<HTMLInputElement>) => {
    const val = e.target.value;
    setInputValue(val);
    onRawValueChange(val);

    // Function tooltip helper if starts with =
    if (val.startsWith('=')) {
      const match = val.match(/^=([A-Za-z]+)\(?/);
      if (match) {
        const query = match[1].toUpperCase();
        const found = FUNCTION_CATALOG.find(f => f.name.startsWith(query));
        setSuggestedFunc(found ? `${found.syntax} — ${found.description}` : null);
      } else {
        setSuggestedFunc(null);
      }
    } else {
      setSuggestedFunc(null);
    }
  };

  const handleKeyDown = (e: React.KeyboardEvent<HTMLInputElement>) => {
    if (e.key === 'Enter') {
      e.preventDefault();
      setSuggestedFunc(null);
      onCommit();
    } else if (e.key === 'Escape') {
      e.preventDefault();
      setSuggestedFunc(null);
      setInputValue(cellRawValue);
      onCancel();
    }
  };

  const handleNameBoxKeyDown = (e: React.KeyboardEvent<HTMLInputElement>) => {
    if (e.key === 'Enter') {
      setIsEditingNameBox(false);
      onJumpToCell(nameBoxValue.trim().toUpperCase());
    }
  };

  return (
    <div className="bg-white border-b border-slate-200 px-3 py-1 flex items-center gap-2 select-none text-xs relative">
      {/* Name Box (Cell Address / Range Indicator) */}
      <div className="w-20 shrink-0">
        {isEditingNameBox ? (
          <input
            type="text"
            value={nameBoxValue}
            onChange={(e) => setNameBoxValue(e.target.value)}
            onKeyDown={handleNameBoxKeyDown}
            onBlur={() => setIsEditingNameBox(false)}
            autoFocus
            className="w-full font-mono text-xs font-semibold px-1.5 py-0.5 border border-emerald-500 rounded outline-none text-center bg-white shadow-xs"
          />
        ) : (
          <div
            onClick={() => setIsEditingNameBox(true)}
            className="font-mono text-xs font-semibold px-2 py-0.5 border border-slate-200 hover:border-slate-300 rounded text-center cursor-pointer bg-slate-50 text-slate-800 truncate"
            title="Name Box — Click to jump to cell"
          >
            {nameBoxValue}
          </div>
        )}
      </div>

      {/* Action Buttons: Cancel, Confirm, Insert Function */}
      <div className="flex items-center gap-0.5 shrink-0 border-r border-slate-200 pr-2">
        <button
          onClick={() => {
            setInputValue(cellRawValue);
            setSuggestedFunc(null);
            onCancel();
          }}
          className="p-1 rounded hover:bg-slate-100 text-slate-400 hover:text-red-600 transition-colors"
          title="Cancel (Esc)"
        >
          <X className="w-3.5 h-3.5" />
        </button>
        <button
          onClick={() => {
            setSuggestedFunc(null);
            onCommit();
          }}
          className="p-1 rounded hover:bg-slate-100 text-slate-400 hover:text-emerald-600 transition-colors"
          title="Enter (Commit)"
        >
          <Check className="w-3.5 h-3.5" />
        </button>
        <button
          onClick={onOpenInsertFunction}
          className="px-1.5 py-0.5 rounded hover:bg-slate-100 text-slate-700 font-serif italic font-bold hover:text-[#107c41] transition-colors"
          title="Insert Function Dialog"
        >
          fx
        </button>
      </div>

      {/* Formula Input */}
      <div className="flex-1 relative">
        <input
          ref={inputRef}
          type="text"
          value={inputValue}
          onChange={handleChange}
          onKeyDown={handleKeyDown}
          placeholder="Enter a value or formula (e.g., =SUM(A1:A10), =IF(B2>100,'Yes','No'))"
          className="w-full font-mono text-xs text-slate-800 bg-transparent px-2 py-1 outline-none focus:bg-emerald-50/30 rounded transition-colors"
        />

        {/* Inline Function Helper Tooltip */}
        {suggestedFunc && (
          <div className="absolute left-0 -bottom-8 bg-slate-900 text-slate-100 text-[11px] font-mono px-2.5 py-1 rounded shadow-lg z-50 pointer-events-none whitespace-nowrap animate-in fade-in duration-75">
            {suggestedFunc}
          </div>
        )}
      </div>
    </div>
  );
};
