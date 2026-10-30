/**
 * @license
 * SPDX-License-Identifier: Apache-2.0
 */

import React, { useState } from 'react';
import { SheetData } from '../types/excel';
import { Plus, ChevronLeft, ChevronRight, MoreHorizontal, Trash2, Copy, Edit2, Palette } from 'lucide-react';

interface SheetTabsProps {
  sheets: SheetData[];
  activeSheetId: string;
  onSelectSheet: (id: string) => void;
  onAddSheet: () => void;
  onRenameSheet: (id: string, newName: string) => void;
  onDeleteSheet: (id: string) => void;
  onDuplicateSheet: (id: string) => void;
  onChangeTabColor: (id: string, color: string) => void;
}

export const SheetTabs: React.FC<SheetTabsProps> = ({
  sheets,
  activeSheetId,
  onSelectSheet,
  onAddSheet,
  onRenameSheet,
  onDeleteSheet,
  onDuplicateSheet,
  onChangeTabColor,
}) => {
  const [editingSheetId, setEditingSheetId] = useState<string | null>(null);
  const [renameInput, setRenameInput] = useState('');
  const [contextMenu, setContextMenu] = useState<{ id: string; x: number; y: number } | null>(null);

  const startRename = (sheet: SheetData) => {
    setEditingSheetId(sheet.id);
    setRenameInput(sheet.name);
    setContextMenu(null);
  };

  const commitRename = (id: string) => {
    if (renameInput.trim()) {
      onRenameSheet(id, renameInput.trim());
    }
    setEditingSheetId(null);
  };

  const handleTabContextMenu = (e: React.MouseEvent, sheetId: string) => {
    e.preventDefault();
    setContextMenu({ id: sheetId, x: e.clientX, y: e.clientY });
  };

  const tabColors = ['#107c41', '#0284c7', '#8b5cf6', '#ea580c', '#e11d48', '#d97706', '#475569'];

  return (
    <div
      onClick={() => setContextMenu(null)}
      className="bg-slate-100 border-t border-slate-300 px-2 py-1 flex items-center justify-between select-none text-xs text-slate-700"
    >
      {/* Left: Navigation and Tabs List */}
      <div className="flex items-center gap-1 min-w-0 overflow-x-auto">
        <button
          onClick={onAddSheet}
          className="p-1 rounded hover:bg-slate-200 text-slate-700 transition-colors shrink-0"
          title="Add New Sheet"
        >
          <Plus className="w-4 h-4" />
        </button>

        <div className="h-4 w-[1px] bg-slate-300 mx-1 shrink-0" />

        {/* Tab Items */}
        <div className="flex items-center gap-1 overflow-x-auto">
          {sheets.map((sheet) => {
            const isActive = sheet.id === activeSheetId;
            const isEditing = sheet.id === editingSheetId;

            return (
              <div
                key={sheet.id}
                onClick={() => onSelectSheet(sheet.id)}
                onDoubleClick={() => startRename(sheet)}
                onContextMenu={(e) => handleTabContextMenu(e, sheet.id)}
                className={`group relative flex items-center gap-1.5 px-3 py-1 rounded-t-md cursor-pointer transition-all border-t-2 ${
                  isActive
                    ? 'bg-white text-slate-900 font-semibold shadow-xs border-[#107c41]'
                    : 'bg-slate-200/70 hover:bg-slate-200 text-slate-600 border-transparent'
                }`}
                style={{
                  borderBottom: sheet.tabColor ? `3px solid ${sheet.tabColor}` : undefined
                }}
              >
                {isEditing ? (
                  <input
                    type="text"
                    value={renameInput}
                    onChange={(e) => setRenameInput(e.target.value)}
                    onBlur={() => commitRename(sheet.id)}
                    onKeyDown={(e) => e.key === 'Enter' && commitRename(sheet.id)}
                    autoFocus
                    className="font-medium text-xs bg-white border border-emerald-500 rounded px-1 outline-none w-24"
                  />
                ) : (
                  <span className="truncate max-w-[130px]">{sheet.name}</span>
                )}
              </div>
            );
          })}
        </div>
      </div>

      {/* Tab Context Menu */}
      {contextMenu && (
        <div
          className="fixed bg-white border border-slate-200 rounded-lg shadow-xl py-1 z-50 text-xs text-slate-700 w-44"
          style={{ left: contextMenu.x, top: contextMenu.y - 140 }}
          onClick={(e) => e.stopPropagation()}
        >
          <button
            onClick={() => {
              const target = sheets.find(s => s.id === contextMenu.id);
              if (target) startRename(target);
            }}
            className="w-full text-left px-3 py-1.5 hover:bg-slate-100 flex items-center gap-2"
          >
            <Edit2 className="w-3.5 h-3.5 text-slate-500" /> Rename
          </button>
          <button
            onClick={() => {
              onDuplicateSheet(contextMenu.id);
              setContextMenu(null);
            }}
            className="w-full text-left px-3 py-1.5 hover:bg-slate-100 flex items-center gap-2"
          >
            <Copy className="w-3.5 h-3.5 text-slate-500" /> Duplicate Sheet
          </button>
          <button
            onClick={() => {
              if (sheets.length > 1) {
                onDeleteSheet(contextMenu.id);
              } else {
                alert('A workbook must have at least one visible worksheet.');
              }
              setContextMenu(null);
            }}
            className="w-full text-left px-3 py-1.5 hover:bg-red-50 text-rose-600 flex items-center gap-2"
          >
            <Trash2 className="w-3.5 h-3.5" /> Delete Sheet
          </button>
          <div className="my-1 border-t border-slate-100" />
          <div className="px-3 py-1 text-[10px] uppercase font-bold text-slate-400">Tab Color</div>
          <div className="px-3 py-1 flex items-center gap-1.5 flex-wrap">
            {tabColors.map((color) => (
              <button
                key={color}
                onClick={() => {
                  onChangeTabColor(contextMenu.id, color);
                  setContextMenu(null);
                }}
                className="w-4 h-4 rounded-full border border-slate-300 hover:scale-125 transition-transform"
                style={{ backgroundColor: color }}
              />
            ))}
          </div>
        </div>
      )}
    </div>
  );
};
