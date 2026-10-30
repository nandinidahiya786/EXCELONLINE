/**
 * @license
 * SPDX-License-Identifier: Apache-2.0
 */

import React, { useState } from 'react';
import { X, Search, Replace, Check } from 'lucide-react';
import { SheetData } from '../types/excel';

interface FindReplaceModalProps {
  isOpen: boolean;
  onClose: () => void;
  sheet: SheetData;
  onFindNext: (query: string, matchCase: boolean, exact: boolean) => void;
  onReplaceCurrent: (find: string, replace: string) => void;
  onReplaceAll: (find: string, replace: string, matchCase: boolean, exact: boolean) => void;
}

export const FindReplaceModal: React.FC<FindReplaceModalProps> = ({
  isOpen,
  onClose,
  sheet,
  onFindNext,
  onReplaceCurrent,
  onReplaceAll,
}) => {
  if (!isOpen) return null;

  const [activeTab, setActiveTab] = useState<'find' | 'replace'>('find');
  const [findText, setFindText] = useState('');
  const [replaceText, setReplaceText] = useState('');
  const [matchCase, setMatchCase] = useState(false);
  const [exactMatch, setExactMatch] = useState(false);

  const handleFindSubmit = (e: React.FormEvent) => {
    e.preventDefault();
    if (findText) {
      onFindNext(findText, matchCase, exactMatch);
    }
  };

  return (
    <div className="fixed inset-0 bg-slate-900/30 backdrop-blur-xs z-50 flex items-center justify-center p-4 select-none">
      <div className="bg-white border border-slate-200 rounded-xl shadow-2xl w-full max-w-md overflow-hidden">
        {/* Header Tabs */}
        <div className="flex items-center justify-between px-4 pt-3 border-b border-slate-200 bg-slate-50">
          <div className="flex items-center gap-4 text-xs font-semibold">
            <button
              onClick={() => setActiveTab('find')}
              className={`pb-2.5 transition-colors ${activeTab === 'find' ? 'text-[#107c41] border-b-2 border-[#107c41]' : 'text-slate-500 hover:text-slate-800'}`}
            >
              Find
            </button>
            <button
              onClick={() => setActiveTab('replace')}
              className={`pb-2.5 transition-colors ${activeTab === 'replace' ? 'text-[#107c41] border-b-2 border-[#107c41]' : 'text-slate-500 hover:text-slate-800'}`}
            >
              Replace
            </button>
          </div>
          <button onClick={onClose} className="p-1 text-slate-400 hover:text-slate-600 rounded">
            <X className="w-4 h-4" />
          </button>
        </div>

        {/* Form Body */}
        <form onSubmit={handleFindSubmit} className="p-5 space-y-4 text-xs">
          <div>
            <label className="block font-semibold text-slate-700 mb-1">Find what:</label>
            <input
              type="text"
              value={findText}
              onChange={(e) => setFindText(e.target.value)}
              required
              autoFocus
              className="w-full px-2.5 py-1.5 border border-slate-300 rounded-md outline-none focus:border-[#107c41]"
              placeholder="Text or number to search..."
            />
          </div>

          {activeTab === 'replace' && (
            <div>
              <label className="block font-semibold text-slate-700 mb-1">Replace with:</label>
              <input
                type="text"
                value={replaceText}
                onChange={(e) => setReplaceText(e.target.value)}
                className="w-full px-2.5 py-1.5 border border-slate-300 rounded-md outline-none focus:border-[#107c41]"
                placeholder="Replacement text..."
              />
            </div>
          )}

          {/* Options */}
          <div className="space-y-1.5 pt-1 text-slate-600">
            <label className="flex items-center gap-2 cursor-pointer">
              <input
                type="checkbox"
                checked={matchCase}
                onChange={(e) => setMatchCase(e.target.checked)}
                className="rounded text-[#107c41]"
              />
              <span>Match case</span>
            </label>
            <label className="flex items-center gap-2 cursor-pointer">
              <input
                type="checkbox"
                checked={exactMatch}
                onChange={(e) => setExactMatch(e.target.checked)}
                className="rounded text-[#107c41]"
              />
              <span>Match entire cell contents</span>
            </label>
          </div>

          {/* Action Buttons */}
          <div className="pt-3 border-t border-slate-200 flex items-center justify-end gap-2">
            {activeTab === 'replace' && (
              <>
                <button
                  type="button"
                  onClick={() => onReplaceAll(findText, replaceText, matchCase, exactMatch)}
                  disabled={!findText}
                  className="px-3 py-1.5 text-slate-700 bg-slate-100 hover:bg-slate-200 disabled:opacity-40 rounded-md transition-colors"
                >
                  Replace All
                </button>
                <button
                  type="button"
                  onClick={() => onReplaceCurrent(findText, replaceText)}
                  disabled={!findText}
                  className="px-3 py-1.5 text-slate-700 bg-slate-100 hover:bg-slate-200 disabled:opacity-40 rounded-md transition-colors"
                >
                  Replace
                </button>
              </>
            )}
            <button
              type="submit"
              disabled={!findText}
              className="px-4 py-1.5 font-semibold text-white bg-[#107c41] hover:bg-[#0c6233] disabled:opacity-40 rounded-md shadow-xs transition-colors"
            >
              Find Next
            </button>
          </div>
        </form>
      </div>
    </div>
  );
};
