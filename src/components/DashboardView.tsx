/**
 * @license
 * SPDX-License-Identifier: Apache-2.0
 */

import React, { useState } from 'react';
import { ChartConfig } from '../types/charts';
import { SheetData, WorkbookData } from '../types/excel';
import { ChartWidget } from './ChartWidget';
import { 
  Plus, 
  Sparkles, 
  Maximize2, 
  Minimize2, 
  Download, 
  BarChart3, 
  Table, 
  ArrowLeft,
  Share2
} from 'lucide-react';

interface DashboardViewProps {
  workbook: WorkbookData;
  charts: ChartConfig[];
  onOpenNewChart: (type?: any) => void;
  onEditChart: (chart: ChartConfig) => void;
  onDeleteChart: (chartId: string) => void;
  onReturnToGrid: () => void;
}

export const DashboardView: React.FC<DashboardViewProps> = ({
  workbook,
  charts,
  onOpenNewChart,
  onEditChart,
  onDeleteChart,
  onReturnToGrid,
}) => {
  const [isFullscreen, setIsFullscreen] = useState(false);
  const [activeFilterCategory, setActiveFilterCategory] = useState<string>('all');

  const toggleFullscreen = () => {
    if (!document.fullscreenElement) {
      document.documentElement.requestFullscreen().then(() => setIsFullscreen(true)).catch(() => {});
    } else {
      document.exitFullscreen().then(() => setIsFullscreen(false)).catch(() => {});
    }
  };

  // Separate KPI scorecards from full charts
  const kpiWidgets = charts.filter(c => c.type === 'kpi');
  const visualCharts = charts.filter(c => c.type !== 'kpi');

  return (
    <div className={`flex-1 overflow-y-auto bg-slate-50/70 p-6 ${isFullscreen ? 'fixed inset-0 z-50 bg-white p-8' : ''}`}>
      {/* Dashboard Top Command Header */}
      <div className="max-w-7xl mx-auto mb-6 flex flex-col md:flex-row md:items-center justify-between gap-4 border-b border-slate-200 pb-4">
        <div>
          <div className="flex items-center gap-2 mb-1">
            <button
              onClick={onReturnToGrid}
              className="flex items-center gap-1 text-xs font-semibold text-slate-500 hover:text-slate-800 transition-colors"
            >
              <ArrowLeft className="w-3.5 h-3.5" /> Back to Sheet
            </button>
            <span className="text-slate-300">/</span>
            <span className="text-xs font-semibold text-[#107c41]">Live Analytics Studio</span>
          </div>
          <h1 className="text-xl font-bold text-slate-900 tracking-tight flex items-center gap-2">
            <span>Executive Dashboard</span>
            <span className="text-xs font-normal text-slate-500 bg-emerald-50 text-[#107c41] px-2 py-0.5 rounded-full border border-emerald-200">
              Auto-Synced
            </span>
          </h1>
          <p className="text-xs text-slate-500 mt-0.5">
            Visualizations automatically update when worksheet cell values change
          </p>
        </div>

        {/* Dashboard Actions */}
        <div className="flex items-center gap-2.5">
          <button
            onClick={() => onOpenNewChart()}
            className="flex items-center gap-1.5 px-3 py-1.5 text-xs font-semibold text-white bg-[#107c41] hover:bg-[#0c6233] rounded-lg shadow-xs transition-colors"
          >
            <Plus className="w-3.5 h-3.5" />
            <span>Add Visualization</span>
          </button>

          <button
            onClick={() => window.print()}
            className="flex items-center gap-1.5 px-3 py-1.5 text-xs font-medium text-slate-700 bg-white hover:bg-slate-100 border border-slate-200 rounded-lg shadow-xs transition-colors"
            title="Print or export dashboard as PDF"
          >
            <Download className="w-3.5 h-3.5 text-slate-500" />
            <span>Export View</span>
          </button>

          <button
            onClick={toggleFullscreen}
            className="p-2 text-slate-600 bg-white hover:bg-slate-100 border border-slate-200 rounded-lg shadow-xs transition-colors"
            title={isFullscreen ? 'Exit Fullscreen' : 'Enter Fullscreen Presentation Mode'}
          >
            {isFullscreen ? <Minimize2 className="w-4 h-4" /> : <Maximize2 className="w-4 h-4" />}
          </button>
        </div>
      </div>

      <div className="max-w-7xl mx-auto space-y-6">
        {/* KPI Scorecards Strip */}
        {kpiWidgets.length > 0 && (
          <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-4 gap-4">
            {kpiWidgets.map((kpi) => (
              <ChartWidget
                key={kpi.id}
                config={kpi}
                sheets={workbook.sheets}
                onEdit={() => onEditChart(kpi)}
                onDelete={() => onDeleteChart(kpi.id)}
              />
            ))}
          </div>
        )}

        {/* Visual Charts Grid */}
        {visualCharts.length > 0 ? (
          <div className="grid grid-cols-1 lg:grid-cols-2 gap-6">
            {visualCharts.map((chart) => (
              <ChartWidget
                key={chart.id}
                config={chart}
                sheets={workbook.sheets}
                onEdit={() => onEditChart(chart)}
                onDelete={() => onDeleteChart(chart.id)}
              />
            ))}
          </div>
        ) : (
          /* Empty State if no charts */
          <div className="bg-white border border-dashed border-slate-300 rounded-2xl p-12 text-center max-w-lg mx-auto my-12">
            <div className="w-14 h-14 bg-emerald-50 text-[#107c41] rounded-2xl flex items-center justify-center mx-auto mb-4 shadow-xs">
              <BarChart3 className="w-7 h-7" />
            </div>
            <h3 className="text-base font-bold text-slate-800 mb-1">No Visualizations Created Yet</h3>
            <p className="text-xs text-slate-500 mb-6 leading-relaxed">
              Connect chart widgets directly to cell ranges on your spreadsheet to turn raw data into executive charts and KPIs.
            </p>
            <div className="flex items-center justify-center gap-3">
              <button
                onClick={() => onOpenNewChart('column')}
                className="px-4 py-2 text-xs font-semibold text-white bg-[#107c41] hover:bg-[#0c6233] rounded-lg shadow-sm transition-colors"
              >
                + Create First Chart
              </button>
              <button
                onClick={onReturnToGrid}
                className="px-4 py-2 text-xs font-medium text-slate-700 bg-white hover:bg-slate-50 border border-slate-200 rounded-lg transition-colors"
              >
                Go to Spreadsheet
              </button>
            </div>
          </div>
        )}
      </div>
    </div>
  );
};
