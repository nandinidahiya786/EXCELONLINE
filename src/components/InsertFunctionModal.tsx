/**
 * @license
 * SPDX-License-Identifier: Apache-2.0
 */

import React, { useState } from 'react';
import { X, Search, HelpCircle, Check } from 'lucide-react';
import { FUNCTION_CATALOG, FunctionDoc } from '../utils/formulaEngine';

interface InsertFunctionModalProps {
  isOpen: boolean;
  onClose: () => void;
  onInsertFunction: (formulaSyntax: string) => void;
}

export const InsertFunctionModal: React.FC<InsertFunctionModalProps> = ({
  isOpen,
  onClose,
  onInsertFunction,
}) => {
  if (!isOpen) return null;

  const [searchQuery, setSearchQuery] = useState('');
  const [selectedCategory, setSelectedCategory] = useState<string>('All');
  const [selectedFunc, setSelectedFunc] = useState<FunctionDoc>(FUNCTION_CATALOG[0]);

  const categories = ['All', 'Math', 'Statistical', 'Logical', 'Lookup', 'Text', 'Date & Time', 'Financial'];

  const filteredFunctions = FUNCTION_CATALOG.filter((f) => {
    const matchesCategory = selectedCategory === 'All' || f.category === selectedCategory;
    const matchesQuery = f.name.toLowerCase().includes(searchQuery.toLowerCase()) ||
                         f.description.toLowerCase().includes(searchQuery.toLowerCase());
    return matchesCategory && matchesQuery;
  });

  const handleInsert = (func: FunctionDoc) => {
    onInsertFunction(`=${func.name}()`);
    onClose();
  };

  return (
    <div className="fixed inset-0 bg-slate-900/40 backdrop-blur-xs z-50 flex items-center justify-center p-4 animate-in fade-in duration-100 select-none">
      <div className="bg-white border border-slate-200 rounded-xl shadow-2xl w-full max-w-2xl overflow-hidden flex flex-col max-h-[85vh]">
        {/* Modal Header */}
        <div className="px-5 py-3 border-b border-slate-200 flex items-center justify-between bg-slate-50">
          <div className="flex items-center gap-2">
            <span className="font-serif italic font-bold text-base text-[#107c41]">fx</span>
            <h2 className="text-sm font-bold text-slate-800">Insert Function Library</h2>
          </div>
          <button
            onClick={onClose}
            className="p-1 text-slate-400 hover:text-slate-600 rounded hover:bg-slate-200 transition-colors"
          >
            <X className="w-4 h-4" />
          </button>
        </div>

        {/* Search and Category Filter */}
        <div className="p-4 border-b border-slate-100 flex flex-col sm:flex-row gap-3 bg-white">
          <div className="relative flex-1">
            <Search className="w-4 h-4 text-slate-400 absolute left-3 top-2.5" />
            <input
              type="text"
              placeholder="Search functions (e.g. SUM, VLOOKUP, IF)..."
              value={searchQuery}
              onChange={(e) => setSearchQuery(e.target.value)}
              className="w-full text-xs pl-9 pr-3 py-2 bg-slate-50 border border-slate-200 rounded-lg outline-none focus:border-[#107c41] focus:bg-white transition-colors"
              autoFocus
            />
          </div>

          <select
            value={selectedCategory}
            onChange={(e) => setSelectedCategory(e.target.value)}
            className="text-xs px-3 py-2 bg-slate-50 border border-slate-200 rounded-lg text-slate-700 outline-none cursor-pointer"
          >
            {categories.map((c) => (
              <option key={c} value={c}>{c} Category</option>
            ))}
          </select>
        </div>

        {/* Function List & Details Split */}
        <div className="flex-1 overflow-hidden grid grid-cols-1 md:grid-cols-2 divide-y md:divide-y-0 md:divide-x divide-slate-200 min-h-[300px]">
          {/* Left Column: Function List */}
          <div className="overflow-y-auto p-2 max-h-[340px]">
            {filteredFunctions.map((fn) => {
              const isSelected = selectedFunc?.name === fn.name;
              return (
                <div
                  key={fn.name}
                  onClick={() => setSelectedFunc(fn)}
                  onDoubleClick={() => handleInsert(fn)}
                  className={`px-3 py-2 rounded-lg cursor-pointer transition-colors flex items-center justify-between text-xs mb-1 ${
                    isSelected ? 'bg-emerald-50 text-[#107c41] font-semibold' : 'hover:bg-slate-50 text-slate-700'
                  }`}
                >
                  <span className="font-mono">{fn.name}</span>
                  <span className="text-[10px] text-slate-400">{fn.category}</span>
                </div>
              );
            })}
            {filteredFunctions.length === 0 && (
              <div className="text-center py-10 text-xs text-slate-400">
                No matching functions found
              </div>
            )}
          </div>

          {/* Right Column: Function Documentation & Syntax */}
          {selectedFunc && (
            <div className="p-5 flex flex-col justify-between overflow-y-auto bg-slate-50/50">
              <div className="space-y-4">
                <div>
                  <div className="flex items-center gap-2 mb-1">
                    <span className="text-base font-bold font-mono text-slate-900">{selectedFunc.name}</span>
                    <span className="text-[11px] font-medium bg-slate-200 text-slate-700 px-2 py-0.5 rounded">
                      {selectedFunc.category}
                    </span>
                  </div>
                  <p className="text-xs text-slate-600 leading-relaxed">{selectedFunc.description}</p>
                </div>

                <div className="space-y-1.5">
                  <span className="text-[11px] font-bold text-slate-500 uppercase tracking-wider">Syntax:</span>
                  <div className="p-2.5 bg-white border border-slate-200 rounded-md font-mono text-xs text-[#107c41] select-text">
                    {selectedFunc.syntax}
                  </div>
                </div>

                <div className="space-y-1.5">
                  <span className="text-[11px] font-bold text-slate-500 uppercase tracking-wider">Example Usage:</span>
                  <div className="p-2.5 bg-white border border-slate-200 rounded-md font-mono text-xs text-slate-700 select-text">
                    ={selectedFunc.example}
                  </div>
                </div>
              </div>

              <div className="pt-4 border-t border-slate-200 mt-6 flex justify-end gap-2">
                <button
                  onClick={onClose}
                  className="px-4 py-2 text-xs font-medium text-slate-600 hover:bg-slate-200 rounded-lg transition-colors"
                >
                  Cancel
                </button>
                <button
                  onClick={() => handleInsert(selectedFunc)}
                  className="px-4 py-2 text-xs font-semibold text-white bg-[#107c41] hover:bg-[#0c6233] rounded-lg shadow-sm transition-colors flex items-center gap-1.5"
                >
                  <Check className="w-3.5 h-3.5" />
                  <span>Insert ={selectedFunc.name}()</span>
                </button>
              </div>
            </div>
          )}
        </div>
      </div>
    </div>
  );
};
