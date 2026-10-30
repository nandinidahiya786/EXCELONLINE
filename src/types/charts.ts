/**
 * @license
 * SPDX-License-Identifier: Apache-2.0
 */

export type ChartType = 
  | 'column'
  | 'bar'
  | 'line'
  | 'area'
  | 'pie'
  | 'donut'
  | 'kpi';

export interface ChartConfig {
  id: string;
  title: string;
  subtitle?: string;
  type: ChartType;
  sheetId: string;
  dataRange: string; // e.g. "A1:E6"
  labelColumn: string; // e.g. "A" (row headers / categories)
  seriesColumns: string[]; // e.g. ["B", "C", "D"]
  colors: string[];
  showLegend: boolean;
  showGrid: boolean;
  showValues: boolean;
  isDashboardWidget: boolean;
  kpiDetails?: {
    metricCell: string; // e.g. "C12"
    comparisonCell?: string; // e.g. "C13"
    targetCell?: string;
    format: 'currency' | 'percent' | 'number';
    periodLabel?: string; // e.g. "vs last quarter"
  };
}
