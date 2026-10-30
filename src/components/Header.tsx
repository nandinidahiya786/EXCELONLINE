/**
 * @license
 * SPDX-License-Identifier: Apache-2.0
 */

import React, { useRef, useState } from 'react';
import { 
  FileSpreadsheet, 
  Download, 
  Upload, 
  LayoutDashboard, 
  Table, 
  Plus, 
  FileDown, 
  Sparkles,
  ChevronDown,
  Check,
  FolderOpen
} from 'lucide-react';
import { WorkbookData } from '../types/excel';

interface HeaderProps {
  workbook: WorkbookData;
  activeView: 'grid' | 'dashboard';
  onViewChange: (view: 'grid' | 'dashboard') => void;
  onExportExcel: () => void;
  onExportCSV: () => void;
  onImportFile: (file: File) => void;
  onSelectTemplate: (templateKey: 'executive_sales' | 'financial_model' | 'project_tracker' | 'inventory_mgmt' | 'blank') => void;
  onFileNameChange: (name: string) => void;
  onOpenNewChartModal: () => void;
}

export const Header: React.FC<HeaderProps> = ({
  workbook,
  activeView,
  onViewChange,
  onExportExcel,
  onExportCSV,
  onImportFile,
  onSelectTemplate,
  onFileNameChange,
  onOpenNewChartModal,
}) => {
  const fileInputRef = useRef<HTMLInputElement>(null);
  const [showExportMenu, setShowExportMenu] = useState(false);
  const [showTemplatesMenu, setShowTemplatesMenu] = useState(false);
  const [isEditingName, setIsEditingName] = useState(false);
  const [fileNameInput, setFileNameInput] = useState(workbook.fileName);

  const handleFileChange = (e: React.ChangeEvent<HTMLInputElement>) => {
    if (e.target.files && e.target.files[0]) {
      onImportFile(e.target.files[0]);
    }
  };

  const handleNameSubmit = () => {
    setIsEditingName(false);
    if (fileNameInput.trim()) {
      onFileNameChange(fileNameInput.trim());
    }
  };

  return (
    <header className="bg-white border-b border-slate-200 select-none">
      {/* Top Application Bar */}
      <div className="flex items-center justify-between px-3 py-1.5 gap-4">
        {/* Left Zone: Excel Brand & Document Name */}
        <div className="flex items-center gap-3 min-w-0">
          <div className="flex items-center gap-2 bg-[#107c41] text-white px-2.5 py-1 rounded shadow-sm shrink-0">
            <FileSpreadsheet className="w-5 h-5" />
            <span className="font-bold text-sm tracking-tight">ExcelPro</span>
          </div>

          <div className="flex items-center gap-2 min-w-0">
            {isEditingName ? (
              <input
                type="text"
                value={fileNameInput}
                onChange={(e) => setFileNameInput(e.target.value)}
                onBlur={handleNameSubmit}
                onKeyDown={(e) => e.key === 'Enter' && handleNameSubmit()}
                autoFocus
                className="text-sm font-semibold text-slate-800 bg-white border border-emerald-500 rounded px-2 py-0.5 outline-none w-64 shadow-xs"
              />
            ) : (
              <div
                onClick={() => {
                  setFileNameInput(workbook.fileName);
                  setIsEditingName(true);
                }}
                className="group flex items-center gap-1.5 cursor-pointer px-2 py-0.5 rounded hover:bg-slate-100 transition-colors"
                title="Click to rename workbook"
              >
                <span className="text-sm font-semibold text-slate-800 truncate max-w-[260px]">
                  {workbook.fileName}
                </span>
                <span className="text-[10px] text-slate-400 opacity-0 group-hover:opacity-100 transition-opacity">
                  rename
                </span>
              </div>
            )}
            <span className="text-xs text-slate-400">· Saved locally</span>
          </div>
        </div>

        {/* Center Zone: Mode Switcher Tabs */}
        <div className="flex items-center bg-slate-100 p-0.5 rounded-lg border border-slate-200 shrink-0">
          <button
            onClick={() => onViewChange('grid')}
            className={`flex items-center gap-1.5 px-3 py-1 text-xs font-semibold rounded-md transition-colors ${
              activeView === 'grid'
                ? 'bg-white text-[#107c41] shadow-xs'
                : 'text-slate-600 hover:text-slate-900'
            }`}
          >
            <Table className="w-3.5 h-3.5" />
            <span>Sheet Grid</span>
          </button>
          <button
            onClick={() => onViewChange('dashboard')}
            className={`flex items-center gap-1.5 px-3 py-1 text-xs font-semibold rounded-md transition-colors ${
              activeView === 'dashboard'
                ? 'bg-white text-[#107c41] shadow-xs'
                : 'text-slate-600 hover:text-slate-900'
            }`}
          >
            <LayoutDashboard className="w-3.5 h-3.5" />
            <span>Dashboard Studio</span>
          </button>
        </div>

        {/* Right Zone: Primary Actions */}
        <div className="flex items-center gap-2 shrink-0">
          {/* Templates Selector */}
          <div className="relative">
            <button
              onClick={() => {
                setShowTemplatesMenu(!showTemplatesMenu);
                setShowExportMenu(false);
              }}
              className="flex items-center gap-1.5 px-2.5 py-1.5 text-xs font-medium text-slate-700 bg-slate-50 hover:bg-slate-100 border border-slate-200 rounded-md transition-colors"
            >
              <Sparkles className="w-3.5 h-3.5 text-amber-500" />
              <span>Templates</span>
              <ChevronDown className="w-3 h-3 text-slate-400" />
            </button>

            {showTemplatesMenu && (
              <div 
                className="absolute right-0 mt-1 w-64 bg-white border border-slate-200 rounded-lg shadow-xl py-1.5 z-50 text-xs text-slate-700 animate-in fade-in zoom-in-95 duration-100"
                onClick={() => setShowTemplatesMenu(false)}
              >
                <div className="px-3 py-1 text-[11px] font-semibold text-slate-400 uppercase tracking-wider">
                  Prebuilt Models & Sheets
                </div>
                <button
                  onClick={() => onSelectTemplate('executive_sales')}
                  className="w-full text-left px-3 py-2 hover:bg-emerald-50 hover:text-emerald-900 flex flex-col gap-0.5"
                >
                  <span className="font-semibold text-slate-800">Executive Sales & Revenue</span>
                  <span className="text-[11px] text-slate-500">Multi-region performance, charts & KPIs</span>
                </button>
                <button
                  onClick={() => onSelectTemplate('financial_model')}
                  className="w-full text-left px-3 py-2 hover:bg-emerald-50 hover:text-emerald-900 flex flex-col gap-0.5"
                >
                  <span className="font-semibold text-slate-800">Financial Model & P&L</span>
                  <span className="text-[11px] text-slate-500">5-year forecast with EBIT, CAGR & margins</span>
                </button>
                <button
                  onClick={() => onSelectTemplate('project_tracker')}
                  className="w-full text-left px-3 py-2 hover:bg-emerald-50 hover:text-emerald-900 flex flex-col gap-0.5"
                >
                  <span className="font-semibold text-slate-800">Project Milestone Tracker</span>
                  <span className="text-[11px] text-slate-500">Sprint tasks, priorities & budget vs spent</span>
                </button>
                <button
                  onClick={() => onSelectTemplate('inventory_mgmt')}
                  className="w-full text-left px-3 py-2 hover:bg-emerald-50 hover:text-emerald-900 flex flex-col gap-0.5"
                >
                  <span className="font-semibold text-slate-800">Inventory & Valuation</span>
                  <span className="text-[11px] text-slate-500">Stock levels, reorder points & item share</span>
                </button>
                <div className="my-1 border-t border-slate-100"></div>
                <button
                  onClick={() => onSelectTemplate('blank')}
                  className="w-full text-left px-3 py-2 hover:bg-slate-50 font-medium text-slate-600"
                >
                  + Blank Workbook
                </button>
              </div>
            )}
          </div>

          {/* Import Button */}
          <button
            onClick={() => fileInputRef.current?.click()}
            className="flex items-center gap-1.5 px-2.5 py-1.5 text-xs font-medium text-slate-700 bg-white hover:bg-slate-100 border border-slate-200 rounded-md transition-colors"
            title="Open Excel (.xlsx) or CSV from local drive"
          >
            <FolderOpen className="w-3.5 h-3.5 text-slate-600" />
            <span>Open</span>
          </button>
          <input
            ref={fileInputRef}
            type="file"
            accept=".xlsx, .xls, .csv"
            onChange={handleFileChange}
            className="hidden"
          />

          {/* New Chart / Visual */}
          <button
            onClick={onOpenNewChartModal}
            className="flex items-center gap-1.5 px-2.5 py-1.5 text-xs font-medium text-slate-700 bg-white hover:bg-slate-100 border border-slate-200 rounded-md transition-colors"
            title="Create a new chart or visualization widget"
          >
            <Plus className="w-3.5 h-3.5 text-slate-600" />
            <span>Add Chart</span>
          </button>

          {/* Direct Local Excel Export Button */}
          <div className="relative">
            <div className="flex items-center">
              <button
                onClick={onExportExcel}
                className="flex items-center gap-1.5 px-3 py-1.5 text-xs font-semibold text-white bg-[#107c41] hover:bg-[#0c6233] active:bg-[#094c27] rounded-l-md shadow-xs transition-colors"
                title="Export directly to local drive in authentic Excel (.xlsx) format"
              >
                <Download className="w-3.5 h-3.5" />
                <span>Save to Excel (.xlsx)</span>
              </button>
              <button
                onClick={() => {
                  setShowExportMenu(!showExportMenu);
                  setShowTemplatesMenu(false);
                }}
                className="px-1.5 py-1.5 bg-[#107c41] hover:bg-[#0c6233] text-white border-l border-emerald-700 rounded-r-md transition-colors"
                title="More export options"
              >
                <ChevronDown className="w-3.5 h-3.5" />
              </button>
            </div>

            {showExportMenu && (
              <div 
                className="absolute right-0 mt-1 w-52 bg-white border border-slate-200 rounded-lg shadow-xl py-1 z-50 text-xs text-slate-700 animate-in fade-in zoom-in-95 duration-100"
                onClick={() => setShowExportMenu(false)}
              >
                <button
                  onClick={onExportExcel}
                  className="w-full text-left px-3 py-2 hover:bg-emerald-50 hover:text-emerald-900 flex items-center gap-2"
                >
                  <FileSpreadsheet className="w-4 h-4 text-[#107c41]" />
                  <div>
                    <div className="font-semibold text-slate-800">Microsoft Excel (.xlsx)</div>
                    <div className="text-[10px] text-slate-500">All sheets, formulas & styles</div>
                  </div>
                </button>
                <button
                  onClick={onExportCSV}
                  className="w-full text-left px-3 py-2 hover:bg-slate-50 flex items-center gap-2"
                >
                  <FileDown className="w-4 h-4 text-slate-600" />
                  <div>
                    <div className="font-medium text-slate-800">Comma Separated (.csv)</div>
                    <div className="text-[10px] text-slate-500">Active sheet only</div>
                  </div>
                </button>
                <button
                  onClick={() => window.print()}
                  className="w-full text-left px-3 py-2 hover:bg-slate-50 flex items-center gap-2 border-t border-slate-100"
                >
                  <Download className="w-4 h-4 text-slate-600" />
                  <div>
                    <div className="font-medium text-slate-800">Print / Save as PDF</div>
                    <div className="text-[10px] text-slate-500">Browser print preview</div>
                  </div>
                </button>
              </div>
            )}
          </div>
        </div>
      </div>
    </header>
  );
};
