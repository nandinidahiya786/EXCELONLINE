/**
 * @license
 * SPDX-License-Identifier: Apache-2.0
 */

import React, { useState, useEffect } from 'react';
import { X, BarChart2, PieChart, LineChart, TrendingUp, Sparkles, Check } from 'lucide-react';
import { ChartConfig, ChartType } from '../types/charts';
import { SheetData } from '../types/excel';

interface CreateChartModalProps {
  isOpen: boolean;
  onClose: () => void;
  sheets: SheetData[];
  activeSheetId: string;
  defaultRange: string;
  initialChart?: ChartConfig | null;
  onSaveChart: (chart: ChartConfig) => void;
}

export const CreateChartModal: React.FC<CreateChartModalProps> = ({
  isOpen,
  onClose,
  sheets,
  activeSheetId,
  defaultRange,
  initialChart,
  onSaveChart,
}) => {
  if (!isOpen) return null;

  const [title, setTitle] = useState(initialChart?.title || 'Data Analytics Chart');
  const [subtitle, setSubtitle] = useState(initialChart?.subtitle || '');
  const [type, setType] = useState<ChartType>(initialChart?.type || 'column');
  const [sheetId, setSheetId] = useState(initialChart?.sheetId || activeSheetId);
  const [dataRange, setDataRange] = useState(initialChart?.dataRange || defaultRange || 'A1:E6');
  const [labelColumn, setLabelColumn] = useState(initialChart?.labelColumn || 'A');
  const [seriesColumnsStr, setSeriesColumnsStr] = useState(initialChart?.seriesColumns.join(', ') || 'B, C, D');
  const [colorTheme, setColorTheme] = useState('excel');
  const [showLegend, setShowLegend] = useState(initialChart?.showLegend ?? true);
  const [showGrid, setShowGrid] = useState(initialChart?.showGrid ?? true);

  // KPI specific
  const [kpiCell, setKpiCell] = useState(initialChart?.kpiDetails?.metricCell || 'B5');
  const [kpiFormat, setKpiFormat] = useState<'currency' | 'percent' | 'number'>(initialChart?.kpiDetails?.format || 'currency');
  const [kpiPeriod, setKpiPeriod] = useState(initialChart?.kpiDetails?.periodLabel || 'Annual Metric');

  useEffect(() => {
    if (initialChart) {
      setTitle(initialChart.title);
      setSubtitle(initialChart.subtitle || '');
      setType(initialChart.type);
      setSheetId(initialChart.sheetId);
      setDataRange(initialChart.dataRange);
      setLabelColumn(initialChart.labelColumn);
      setSeriesColumnsStr(initialChart.seriesColumns.join(', '));
      setShowLegend(initialChart.showLegend);
      setShowGrid(initialChart.showGrid);
      if (initialChart.kpiDetails) {
        setKpiCell(initialChart.kpiDetails.metricCell);
        setKpiFormat(initialChart.kpiDetails.format);
        setKpiPeriod(initialChart.kpiDetails.periodLabel || '');
      }
    } else {
      setDataRange(defaultRange || 'A1:E6');
    }
  }, [initialChart, defaultRange]);

  const colorThemes: Record<string, string[]> = {
    excel: ['#107c41', '#3b82f6', '#f59e0b', '#8b5cf6', '#ec4899', '#06b6d4'],
    modern: ['#0284c7', '#10b981', '#6366f1', '#f97316', '#a855f7', '#14b8a6'],
    executive: ['#0f172a', '#334155', '#475569', '#64748b', '#94a3b8', '#cbd5e1'],
    warm: ['#ea580c', '#d97706', '#ca8a04', '#e11d48', '#be123c', '#9f1239'],
  };

  const handleSubmit = (e: React.FormEvent) => {
    e.preventDefault();

    const seriesCols = seriesColumnsStr
      .split(',')
      .map(s => s.trim().toUpperCase())
      .filter(s => s.length > 0);

    const config: ChartConfig = {
      id: initialChart?.id || `chart-${Date.now()}`,
      title: title.trim() || 'Visual Chart',
      subtitle: subtitle.trim() || undefined,
      type,
      sheetId,
      dataRange: dataRange.trim().toUpperCase(),
      labelColumn: labelColumn.trim().toUpperCase() || 'A',
      seriesColumns: seriesCols.length > 0 ? seriesCols : ['B'],
      colors: colorThemes[colorTheme] || colorThemes.excel,
      showLegend,
      showGrid,
      showValues: true,
      isDashboardWidget: true,
      kpiDetails: type === 'kpi' ? {
        metricCell: kpiCell.trim().toUpperCase(),
        format: kpiFormat,
        periodLabel: kpiPeriod.trim() || undefined,
      } : undefined,
    };

    onSaveChart(config);
    onClose();
  };

  return (
    <div className="fixed inset-0 bg-slate-900/40 backdrop-blur-xs z-50 flex items-center justify-center p-4 animate-in fade-in duration-100 select-none">
      <div className="bg-white border border-slate-200 rounded-xl shadow-2xl w-full max-w-lg overflow-hidden flex flex-col">
        {/* Modal Header */}
        <div className="px-5 py-3 border-b border-slate-200 flex items-center justify-between bg-slate-50">
          <div className="flex items-center gap-2">
            <BarChart2 className="w-4 h-4 text-[#107c41]" />
            <h2 className="text-sm font-bold text-slate-800">
              {initialChart ? 'Edit Chart / Visualization' : 'Create Chart or Dashboard Widget'}
            </h2>
          </div>
          <button onClick={onClose} className="p-1 text-slate-400 hover:text-slate-600 rounded">
            <X className="w-4 h-4" />
          </button>
        </div>

        {/* Form Body */}
        <form onSubmit={handleSubmit} className="p-5 space-y-4 text-xs overflow-y-auto max-h-[75vh]">
          {/* Chart Type Selector */}
          <div>
            <label className="block font-semibold text-slate-700 mb-1.5">Chart Type</label>
            <div className="grid grid-cols-4 gap-2">
              {[
                { id: 'column', label: 'Column', icon: BarChart2 },
                { id: 'bar', label: 'Bar', icon: BarChart2, rotate: true },
                { id: 'line', label: 'Line', icon: LineChart },
                { id: 'area', label: 'Area', icon: TrendingUp },
                { id: 'pie', label: 'Pie', icon: PieChart },
                { id: 'donut', label: 'Donut', icon: PieChart },
                { id: 'kpi', label: 'KPI Card', icon: Sparkles },
              ].map((item) => {
                const Icon = item.icon;
                const isSelected = type === item.id;
                return (
                  <button
                    key={item.id}
                    type="button"
                    onClick={() => setType(item.id as ChartType)}
                    className={`flex flex-col items-center gap-1 p-2 rounded-lg border text-center transition-all ${
                      isSelected
                        ? 'border-[#107c41] bg-emerald-50 text-[#107c41] font-bold shadow-xs'
                        : 'border-slate-200 hover:bg-slate-50 text-slate-600'
                    }`}
                  >
                    <Icon className={`w-4 h-4 ${item.rotate ? 'rotate-90' : ''}`} />
                    <span className="text-[11px]">{item.label}</span>
                  </button>
                );
              })}
            </div>
          </div>

          {/* Titles */}
          <div className="grid grid-cols-1 sm:grid-cols-2 gap-3">
            <div>
              <label className="block font-semibold text-slate-700 mb-1">Chart Title</label>
              <input
                type="text"
                value={title}
                onChange={(e) => setTitle(e.target.value)}
                required
                className="w-full px-2.5 py-1.5 border border-slate-300 rounded-md outline-none focus:border-[#107c41]"
                placeholder="e.g. Quarterly Revenue Growth"
              />
            </div>
            <div>
              <label className="block font-semibold text-slate-700 mb-1">Subtitle (Optional)</label>
              <input
                type="text"
                value={subtitle}
                onChange={(e) => setSubtitle(e.target.value)}
                className="w-full px-2.5 py-1.5 border border-slate-300 rounded-md outline-none focus:border-[#107c41]"
                placeholder="e.g. Q1 - Q4 Comparison"
              />
            </div>
          </div>

          {/* KPI Configuration */}
          {type === 'kpi' ? (
            <div className="p-3 bg-emerald-50/50 border border-emerald-200 rounded-lg space-y-3">
              <span className="font-bold text-emerald-900 block text-[11px] uppercase tracking-wider">
                KPI Metric Cell Link
              </span>
              <div className="grid grid-cols-3 gap-2">
                <div>
                  <label className="block text-[11px] text-slate-600 mb-1">Target Cell (e.g. F10)</label>
                  <input
                    type="text"
                    value={kpiCell}
                    onChange={(e) => setKpiCell(e.target.value)}
                    required
                    className="w-full font-mono px-2 py-1 bg-white border border-slate-300 rounded outline-none"
                  />
                </div>
                <div>
                  <label className="block text-[11px] text-slate-600 mb-1">Format</label>
                  <select
                    value={kpiFormat}
                    onChange={(e) => setKpiFormat(e.target.value as any)}
                    className="w-full px-2 py-1 bg-white border border-slate-300 rounded outline-none"
                  >
                    <option value="currency">Currency ($)</option>
                    <option value="percent">Percentage (%)</option>
                    <option value="number">Number</option>
                  </select>
                </div>
                <div>
                  <label className="block text-[11px] text-slate-600 mb-1">Period Label</label>
                  <input
                    type="text"
                    value={kpiPeriod}
                    onChange={(e) => setKpiPeriod(e.target.value)}
                    placeholder="e.g. vs last quarter"
                    className="w-full px-2 py-1 bg-white border border-slate-300 rounded outline-none"
                  />
                </div>
              </div>
            </div>
          ) : (
            /* Multi-Series Chart Configuration */
            <div className="space-y-3">
              <div className="grid grid-cols-1 sm:grid-cols-3 gap-2">
                <div>
                  <label className="block font-semibold text-slate-700 mb-1">Data Range</label>
                  <input
                    type="text"
                    value={dataRange}
                    onChange={(e) => setDataRange(e.target.value)}
                    required
                    className="w-full font-mono px-2.5 py-1.5 border border-slate-300 rounded-md outline-none focus:border-[#107c41]"
                    placeholder="e.g. A4:E9"
                  />
                </div>
                <div>
                  <label className="block font-semibold text-slate-700 mb-1">Category Label Col</label>
                  <input
                    type="text"
                    value={labelColumn}
                    onChange={(e) => setLabelColumn(e.target.value)}
                    required
                    className="w-full font-mono px-2.5 py-1.5 border border-slate-300 rounded-md outline-none focus:border-[#107c41]"
                    placeholder="e.g. A"
                  />
                </div>
                <div>
                  <label className="block font-semibold text-slate-700 mb-1">Series Columns</label>
                  <input
                    type="text"
                    value={seriesColumnsStr}
                    onChange={(e) => setSeriesColumnsStr(e.target.value)}
                    required
                    className="w-full font-mono px-2.5 py-1.5 border border-slate-300 rounded-md outline-none focus:border-[#107c41]"
                    placeholder="e.g. B, C, D"
                  />
                </div>
              </div>

              {/* Color Theme & Toggles */}
              <div className="flex items-center justify-between pt-2">
                <div className="flex items-center gap-2">
                  <span className="font-semibold text-slate-700">Color Palette:</span>
                  <select
                    value={colorTheme}
                    onChange={(e) => setColorTheme(e.target.value)}
                    className="px-2 py-1 border border-slate-300 rounded outline-none"
                  >
                    <option value="excel">Classic Excel (Green & Blue)</option>
                    <option value="modern">Modern Emerald & Sky</option>
                    <option value="executive">Executive Monolith (Slate)</option>
                    <option value="warm">Vibrant Sunset</option>
                  </select>
                </div>

                <div className="flex items-center gap-3">
                  <label className="flex items-center gap-1.5 cursor-pointer">
                    <input
                      type="checkbox"
                      checked={showLegend}
                      onChange={(e) => setShowLegend(e.target.checked)}
                      className="rounded text-[#107c41]"
                    />
                    <span>Legend</span>
                  </label>
                  <label className="flex items-center gap-1.5 cursor-pointer">
                    <input
                      type="checkbox"
                      checked={showGrid}
                      onChange={(e) => setShowGrid(e.target.checked)}
                      className="rounded text-[#107c41]"
                    />
                    <span>Gridlines</span>
                  </label>
                </div>
              </div>
            </div>
          )}

          {/* Target Sheet */}
          <div>
            <label className="block font-semibold text-slate-700 mb-1">Source Sheet</label>
            <select
              value={sheetId}
              onChange={(e) => setSheetId(e.target.value)}
              className="w-full px-2.5 py-1.5 border border-slate-300 rounded-md outline-none"
            >
              {sheets.map((s) => (
                <option key={s.id} value={s.id}>{s.name}</option>
              ))}
            </select>
          </div>

          {/* Action Buttons */}
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
              <span>{initialChart ? 'Update Chart' : 'Add to Dashboard'}</span>
            </button>
          </div>
        </form>
      </div>
    </div>
  );
};
