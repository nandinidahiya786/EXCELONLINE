/**
 * @license
 * SPDX-License-Identifier: Apache-2.0
 */

import React, { useMemo } from 'react';
import { ChartConfig } from '../types/charts';
import { SheetData } from '../types/excel';
import { parseRange, coordsToCellId } from '../utils/cellUtils';
import { formatCellValue } from '../utils/numberFormat';
import { MoreVertical, Trash2, Edit3, ArrowUpRight, ArrowDownRight } from 'lucide-react';

interface ChartWidgetProps {
  config: ChartConfig;
  sheets: SheetData[];
  onEdit: () => void;
  onDelete: () => void;
}

export const ChartWidget: React.FC<ChartWidgetProps> = ({
  config,
  sheets,
  onEdit,
  onDelete,
}) => {
  const sheet = sheets.find(s => s.id === config.sheetId) || sheets[0];

  // Extract data from range
  const chartData = useMemo(() => {
    if (!sheet) return { labels: [], datasets: [] };

    if (config.type === 'kpi' && config.kpiDetails) {
      const metricCell = sheet.cells?.[config.kpiDetails.metricCell];
      const comparisonCell = config.kpiDetails.comparisonCell ? sheet.cells?.[config.kpiDetails.comparisonCell] : null;
      return {
        metricValue: metricCell?.value !== undefined ? metricCell.value : metricCell?.raw ?? 0,
        comparisonValue: comparisonCell?.value !== undefined ? comparisonCell.value : comparisonCell?.raw ?? null,
        labels: [],
        datasets: [],
      };
    }

    const range = parseRange(config.dataRange);
    if (!range) return { labels: [], datasets: [] };

    const labels: string[] = [];
    const datasets: { label: string; values: number[]; color: string }[] = [];

    // Header row or category labels
    const hasHeaderRow = true;
    const startDataRow = hasHeaderRow ? range.startRow + 1 : range.startRow;

    // Series labels from header row
    config.seriesColumns.forEach((colLetter, idx) => {
      const colIdx = colLetter.toUpperCase().charCodeAt(0) - 65;
      const headerCellId = coordsToCellId(range.startRow, colIdx);
      const headerVal = sheet.cells?.[headerCellId]?.value ?? sheet.cells?.[headerCellId]?.raw ?? `Series ${idx + 1}`;
      datasets.push({
        label: String(headerVal),
        values: [],
        color: config.colors[idx % config.colors.length] || '#107c41',
      });
    });

    // Populate row labels and numeric values
    const labelColIdx = config.labelColumn.toUpperCase().charCodeAt(0) - 65;
    for (let r = startDataRow; r <= range.endRow; r++) {
      const labelCellId = coordsToCellId(r, labelColIdx);
      const labelVal = sheet.cells?.[labelCellId]?.value ?? sheet.cells?.[labelCellId]?.raw ?? `Row ${r + 1}`;
      labels.push(String(labelVal));

      config.seriesColumns.forEach((colLetter, cIdx) => {
        const colIdx = colLetter.toUpperCase().charCodeAt(0) - 65;
        const cellId = coordsToCellId(r, colIdx);
        const cell = sheet.cells?.[cellId];
        const num = Number(cell?.value !== undefined ? cell.value : cell?.raw);
        datasets[cIdx].values.push(isNaN(num) ? 0 : num);
      });
    }

    return { labels, datasets };
  }, [config, sheet]);

  // KPI Scorecard Rendering
  if (config.type === 'kpi' && config.kpiDetails) {
    const formattedMetric = formatCellValue(chartData.metricValue, {
      numberFormat: config.kpiDetails.format,
      currencySymbol: '$',
      decimals: config.kpiDetails.format === 'percent' ? 1 : 0,
    });

    return (
      <div className="bg-white border border-slate-200 rounded-xl p-5 shadow-xs hover:shadow-md transition-shadow relative group">
        <div className="flex items-center justify-between mb-2">
          <span className="text-xs font-semibold text-slate-500 uppercase tracking-wider">{config.title}</span>
          <div className="opacity-0 group-hover:opacity-100 transition-opacity flex items-center gap-1">
            <button onClick={onEdit} className="p-1 hover:bg-slate-100 rounded text-slate-400 hover:text-slate-700">
              <Edit3 className="w-3.5 h-3.5" />
            </button>
            <button onClick={onDelete} className="p-1 hover:bg-red-50 rounded text-slate-400 hover:text-red-600">
              <Trash2 className="w-3.5 h-3.5" />
            </button>
          </div>
        </div>

        <div className="text-2xl font-bold font-mono text-slate-900 tabular-nums">
          {formattedMetric}
        </div>

        {config.kpiDetails.periodLabel && (
          <div className="flex items-center gap-1.5 mt-2 text-xs text-slate-500">
            <span className="text-emerald-600 font-semibold flex items-center gap-0.5">
              <ArrowUpRight className="w-3.5 h-3.5" /> Live Linked
            </span>
            <span>· {config.kpiDetails.periodLabel}</span>
          </div>
        )}
      </div>
    );
  }

  // Calculate SVG scales
  const width = 460;
  const height = 240;
  const padding = { top: 30, right: 20, bottom: 40, left: 55 };
  const chartW = width - padding.left - padding.right;
  const chartH = height - padding.top - padding.bottom;

  // Max value calculation
  let maxVal = 0;
  chartData.datasets.forEach(ds => {
    ds.values.forEach(v => {
      if (v > maxVal) maxVal = v;
    });
  });
  if (maxVal === 0) maxVal = 100;
  // Nice round max
  const niceMax = Math.ceil(maxVal * 1.15);

  return (
    <div className="bg-white border border-slate-200 rounded-xl p-4 shadow-xs hover:shadow-md transition-shadow flex flex-col relative group">
      {/* Chart Header */}
      <div className="flex items-start justify-between mb-2">
        <div>
          <h3 className="text-sm font-bold text-slate-900">{config.title}</h3>
          {config.subtitle && (
            <p className="text-xs text-slate-500">{config.subtitle}</p>
          )}
        </div>

        <div className="opacity-0 group-hover:opacity-100 transition-opacity flex items-center gap-1">
          <button onClick={onEdit} className="p-1 hover:bg-slate-100 rounded text-slate-400 hover:text-slate-700" title="Edit Chart Settings">
            <Edit3 className="w-3.5 h-3.5" />
          </button>
          <button onClick={onDelete} className="p-1 hover:bg-red-50 rounded text-slate-400 hover:text-red-600" title="Delete Chart">
            <Trash2 className="w-3.5 h-3.5" />
          </button>
        </div>
      </div>

      {/* SVG Chart Rendering */}
      <div className="flex-1 min-h-[220px] flex items-center justify-center relative">
        {/* ================= COLUMN / BAR CHART ================= */}
        {config.type === 'column' && (
          <svg viewBox={`0 0 ${width} ${height}`} className="w-full h-full overflow-visible">
            {/* Gridlines */}
            {config.showGrid && [0, 0.25, 0.5, 0.75, 1].map((pct, idx) => {
              const y = padding.top + chartH * (1 - pct);
              const val = Math.round(niceMax * pct);
              return (
                <g key={idx}>
                  <line x1={padding.left} y1={y} x2={width - padding.right} y2={y} stroke="#f1f5f9" strokeDasharray="3 3" />
                  <text x={padding.left - 8} y={y + 3} textAnchor="end" className="text-[10px] fill-slate-400 font-mono">
                    {val >= 1000 ? `${(val / 1000).toFixed(0)}k` : val}
                  </text>
                </g>
              );
            })}

            {/* Bars */}
            {chartData.labels.map((lbl, lIdx) => {
              const groupW = chartW / chartData.labels.length;
              const barCount = chartData.datasets.length;
              const barW = Math.max(8, (groupW * 0.7) / barCount);
              const groupStart = padding.left + lIdx * groupW + (groupW * 0.15);

              return (
                <g key={lIdx}>
                  {chartData.datasets.map((ds, dIdx) => {
                    const val = ds.values[lIdx] || 0;
                    const barH = (val / niceMax) * chartH;
                    const x = groupStart + dIdx * barW;
                    const y = padding.top + chartH - barH;

                    return (
                      <g key={dIdx} className="group/bar cursor-pointer">
                        <rect
                          x={x}
                          y={y}
                          width={barW - 2}
                          height={Math.max(2, barH)}
                          fill={ds.color}
                          rx={3}
                          className="transition-all duration-300 hover:brightness-110"
                        />
                        {/* Tooltip hover */}
                        <title>{`${lbl} — ${ds.label}: ${val.toLocaleString()}`}</title>
                      </g>
                    );
                  })}
                  {/* Category Label */}
                  <text
                    x={groupStart + (barCount * barW) / 2}
                    y={height - 12}
                    textAnchor="middle"
                    className="text-[10px] fill-slate-600 font-medium"
                  >
                    {lbl.length > 10 ? `${lbl.substring(0, 9)}…` : lbl}
                  </text>
                </g>
              );
            })}
          </svg>
        )}

        {/* ================= HORIZONTAL BAR CHART ================= */}
        {config.type === 'bar' && (
          <svg viewBox={`0 0 ${width} ${height}`} className="w-full h-full overflow-visible">
            {chartData.labels.map((lbl, lIdx) => {
              const rowH = chartH / chartData.labels.length;
              const ds = chartData.datasets[0] || { values: [], color: '#107c41' };
              const val = ds.values[lIdx] || 0;
              const barW = (val / niceMax) * (chartW - 40);
              const y = padding.top + lIdx * rowH + 4;

              return (
                <g key={lIdx} className="group/bar cursor-pointer">
                  <text
                    x={padding.left + 50}
                    y={y + rowH / 2 - 2}
                    textAnchor="end"
                    className="text-[10px] fill-slate-700 font-medium"
                  >
                    {lbl.length > 14 ? `${lbl.substring(0, 13)}…` : lbl}
                  </text>
                  <rect
                    x={padding.left + 60}
                    y={y}
                    width={Math.max(4, barW)}
                    height={Math.max(12, rowH - 8)}
                    fill={ds.color}
                    rx={3}
                    className="transition-all duration-300 hover:brightness-110"
                  />
                  <text
                    x={padding.left + 65 + barW}
                    y={y + (rowH - 8) / 2 + 3}
                    className="text-[10px] fill-slate-500 font-mono font-semibold"
                  >
                    {val >= 1000 ? `${(val / 1000).toFixed(0)}k` : val}
                  </text>
                  <title>{`${lbl}: ${val.toLocaleString()}`}</title>
                </g>
              );
            })}
          </svg>
        )}

        {/* ================= LINE & AREA CHART ================= */}
        {(config.type === 'line' || config.type === 'area') && (
          <svg viewBox={`0 0 ${width} ${height}`} className="w-full h-full overflow-visible">
            {/* Gridlines */}
            {config.showGrid && [0, 0.25, 0.5, 0.75, 1].map((pct, idx) => {
              const y = padding.top + chartH * (1 - pct);
              return (
                <line key={idx} x1={padding.left} y1={y} x2={width - padding.right} y2={y} stroke="#f1f5f9" strokeDasharray="3 3" />
              );
            })}

            {/* Datasets */}
            {chartData.datasets.map((ds, dIdx) => {
              const points = ds.values.map((v, idx) => {
                const step = chartW / Math.max(1, chartData.labels.length - 1);
                const x = padding.left + idx * step;
                const y = padding.top + chartH - (v / niceMax) * chartH;
                return { x, y, v };
              });

              const pathData = points.reduce((acc, p, idx) => `${acc} ${idx === 0 ? 'M' : 'L'} ${p.x} ${p.y}`, '');
              const areaPath = points.length > 0
                ? `${pathData} L ${points[points.length - 1].x} ${padding.top + chartH} L ${points[0].x} ${padding.top + chartH} Z`
                : '';

              return (
                <g key={dIdx}>
                  {config.type === 'area' && (
                    <path d={areaPath} fill={ds.color} fillOpacity={0.15} />
                  )}
                  <path
                    d={pathData}
                    fill="none"
                    stroke={ds.color}
                    strokeWidth={2.5}
                    strokeLinecap="round"
                    strokeLinejoin="round"
                  />
                  {points.map((p, pIdx) => (
                    <circle
                      key={pIdx}
                      cx={p.x}
                      cy={p.y}
                      r={3.5}
                      fill="#ffffff"
                      stroke={ds.color}
                      strokeWidth={2}
                      className="cursor-pointer hover:r-5 transition-all"
                    >
                      <title>{`${chartData.labels[pIdx]} — ${ds.label}: ${p.v.toLocaleString()}`}</title>
                    </circle>
                  ))}
                </g>
              );
            })}

            {/* X Axis Labels */}
            {chartData.labels.map((lbl, idx) => {
              const step = chartW / Math.max(1, chartData.labels.length - 1);
              const x = padding.left + idx * step;
              return (
                <text key={idx} x={x} y={height - 12} textAnchor="middle" className="text-[10px] fill-slate-500 font-medium">
                  {lbl.length > 8 ? `${lbl.substring(0, 7)}…` : lbl}
                </text>
              );
            })}
          </svg>
        )}

        {/* ================= PIE & DONUT CHART ================= */}
        {(config.type === 'pie' || config.type === 'donut') && (
          <div className="flex items-center justify-around w-full">
            <svg viewBox="0 0 200 200" className="w-44 h-44 overflow-visible">
              {(() => {
                const values = chartData.datasets[0]?.values || [];
                const total = values.reduce((a, b) => a + b, 0) || 1;
                let currentAngle = 0;
                const isDonut = config.type === 'donut';
                const outerR = 80;
                const innerR = isDonut ? 48 : 0;
                const cx = 100;
                const cy = 100;

                return values.map((val, idx) => {
                  const sliceAngle = (val / total) * 2 * Math.PI;
                  const startAngle = currentAngle;
                  const endAngle = currentAngle + sliceAngle;
                  currentAngle += sliceAngle;

                  const x1 = cx + outerR * Math.cos(startAngle);
                  const y1 = cy + outerR * Math.sin(startAngle);
                  const x2 = cx + outerR * Math.cos(endAngle);
                  const y2 = cy + outerR * Math.sin(endAngle);

                  const x3 = cx + innerR * Math.cos(endAngle);
                  const y3 = cy + innerR * Math.sin(endAngle);
                  const x4 = cx + innerR * Math.cos(startAngle);
                  const y4 = cy + innerR * Math.sin(startAngle);

                  const largeArc = sliceAngle > Math.PI ? 1 : 0;
                  const path = isDonut
                    ? `M ${x1} ${y1} A ${outerR} ${outerR} 0 ${largeArc} 1 ${x2} ${y2} L ${x3} ${y3} A ${innerR} ${innerR} 0 ${largeArc} 0 ${x4} ${y4} Z`
                    : `M ${cx} ${cy} L ${x1} ${y1} A ${outerR} ${outerR} 0 ${largeArc} 1 ${x2} ${y2} Z`;

                  const color = config.colors[idx % config.colors.length] || '#107c41';
                  const pct = ((val / total) * 100).toFixed(1);

                  return (
                    <g key={idx} className="cursor-pointer group/slice">
                      <path
                        d={path}
                        fill={color}
                        stroke="#ffffff"
                        strokeWidth={2}
                        className="transition-all hover:opacity-90"
                      />
                      <title>{`${chartData.labels[idx]}: ${val.toLocaleString()} (${pct}%)`}</title>
                    </g>
                  );
                });
              })()}
            </svg>

            {/* Donut Legend */}
            <div className="flex flex-col gap-1.5 max-h-40 overflow-y-auto pr-2">
              {chartData.labels.map((lbl, idx) => {
                const color = config.colors[idx % config.colors.length] || '#107c41';
                const val = chartData.datasets[0]?.values[idx] || 0;
                return (
                  <div key={idx} className="flex items-center gap-2 text-xs">
                    <div className="w-2.5 h-2.5 rounded-full shrink-0" style={{ backgroundColor: color }} />
                    <span className="text-slate-700 truncate max-w-[90px]">{lbl}</span>
                    <span className="font-mono text-slate-400 tabular-nums ml-auto text-[11px]">{val.toLocaleString()}</span>
                  </div>
                );
              })}
            </div>
          </div>
        )}
      </div>

      {/* Legend below for multi-series */}
      {config.showLegend && (config.type === 'column' || config.type === 'line' || config.type === 'area') && (
        <div className="flex items-center justify-center gap-4 mt-2 pt-2 border-t border-slate-100 flex-wrap">
          {chartData.datasets.map((ds, idx) => (
            <div key={idx} className="flex items-center gap-1.5 text-xs text-slate-600">
              <div className="w-3 h-3 rounded-xs" style={{ backgroundColor: ds.color }} />
              <span>{ds.label}</span>
            </div>
          ))}
        </div>
      )}
    </div>
  );
};
