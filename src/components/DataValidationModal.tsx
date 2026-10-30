/**
 * @license
 * SPDX-License-Identifier: Apache-2.0
 */

import React, { useState } from 'react';
import { X, Check, ListFilter } from 'lucide-react';
import { DataValidationRule } from '../types/excel';

interface DataValidationModalProps {
  isOpen: boolean;
  onClose: () => void;
  selectedRangeStr: string;
  onApplyValidation: (rule: DataValidationRule) => void;
  onClearValidation: () => void;
}

export const DataValidationModal: React.FC<DataValidationModalProps> = ({
  isOpen,
  onClose,
  selectedRangeStr,
  onApplyValidation,
  onClearValidation,
}) => {
  if (!isOpen) return null;

  const [validationType, setValidationType] = useState<'list' | 'wholeNumber' | 'decimal'>('list');
  const [listOptions, setListOptions] = useState('High, Medium, Low');
  const [minValue, setMinValue] = useState('0');
  const [maxValue, setMaxValue] = useState('100');

  const handleSubmit = (e: React.FormEvent) => {
    e.preventDefault();
    if (validationType === 'list') {
      const opts = listOptions.split(',').map(s => s.trim()).filter(Boolean);
      onApplyValidation({ type: 'list', options: opts });
    } else {
      onApplyValidation({
        type: validationType,
        min: Number(minValue),
        max: Number(maxValue),
      });
    }
    onClose();
  };

  return (
    <div className="fixed inset-0 bg-slate-900/40 backdrop-blur-xs z-50 flex items-center justify-center p-4 select-none">
      <div className="bg-white border border-slate-200 rounded-xl shadow-2xl w-full max-w-md overflow-hidden">
        <div className="px-5 py-3 border-b border-slate-200 flex items-center justify-between bg-slate-50">
          <div className="flex items-center gap-2">
            <ListFilter className="w-4 h-4 text-[#107c41]" />
            <h2 className="text-sm font-bold text-slate-800">Data Validation</h2>
          </div>
          <button onClick={onClose} className="p-1 text-slate-400 hover:text-slate-600 rounded">
            <X className="w-4 h-4" />
          </button>
        </div>

        <form onSubmit={handleSubmit} className="p-5 space-y-4 text-xs">
          <div>
            <span className="text-slate-500 block mb-1">Applying to range:</span>
            <span className="font-mono font-bold text-[#107c41] bg-emerald-50 px-2 py-0.5 rounded border border-emerald-200">
              {selectedRangeStr}
            </span>
          </div>

          <div>
            <label className="block font-semibold text-slate-700 mb-1">Validation Criteria</label>
            <select
              value={validationType}
              onChange={(e) => setValidationType(e.target.value as any)}
              className="w-full px-2.5 py-1.5 border border-slate-300 rounded-md outline-none"
            >
              <option value="list">List (Dropdown Menu)</option>
              <option value="wholeNumber">Whole Number</option>
              <option value="decimal">Decimal</option>
            </select>
          </div>

          {validationType === 'list' ? (
            <div>
              <label className="block font-semibold text-slate-700 mb-1">
                Dropdown Source (comma-separated):
              </label>
              <input
                type="text"
                value={listOptions}
                onChange={(e) => setListOptions(e.target.value)}
                required
                className="w-full px-2.5 py-1.5 border border-slate-300 rounded-md outline-none focus:border-[#107c41]"
                placeholder="e.g. Approved, Pending, Rejected"
              />
            </div>
          ) : (
            <div className="grid grid-cols-2 gap-3">
              <div>
                <label className="block font-semibold text-slate-700 mb-1">Minimum</label>
                <input
                  type="number"
                  value={minValue}
                  onChange={(e) => setMinValue(e.target.value)}
                  className="w-full px-2.5 py-1.5 border border-slate-300 rounded-md outline-none"
                />
              </div>
              <div>
                <label className="block font-semibold text-slate-700 mb-1">Maximum</label>
                <input
                  type="number"
                  value={maxValue}
                  onChange={(e) => setMaxValue(e.target.value)}
                  className="w-full px-2.5 py-1.5 border border-slate-300 rounded-md outline-none"
                />
              </div>
            </div>
          )}

          <div className="pt-4 border-t border-slate-200 flex justify-between items-center">
            <button
              type="button"
              onClick={() => { onClearValidation(); onClose(); }}
              className="text-xs text-rose-600 hover:underline"
            >
              Clear All Validation
            </button>
            <div className="flex gap-2">
              <button
                type="button"
                onClick={onClose}
                className="px-3 py-1.5 text-slate-600 hover:bg-slate-100 rounded-lg transition-colors"
              >
                Cancel
              </button>
              <button
                type="submit"
                className="px-4 py-1.5 font-semibold text-white bg-[#107c41] hover:bg-[#0c6233] rounded-lg shadow-sm transition-colors flex items-center gap-1.5"
              >
                <Check className="w-3.5 h-3.5" />
                <span>Apply Rule</span>
              </button>
            </div>
          </div>
        </form>
      </div>
    </div>
  );
};
