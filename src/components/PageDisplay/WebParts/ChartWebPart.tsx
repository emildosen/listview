import { useState, useEffect, useCallback, useMemo } from 'react';
import {
  makeStyles,
  tokens,
  Text,
  Spinner,
  mergeClasses,
} from '@fluentui/react-components';
import {
  DonutChart,
  VerticalBarChart,
  LineChart,
  HorizontalBarChartWithAxis,
  AreaChart,
  GaugeChart,
  HeatMapChart,
  ScatterChart,
  GanttChart,
} from '@fluentui/react-charts';
import type {
  ChartProps,
  VerticalBarChartDataPoint,
  LineChartPoints,
  HorizontalBarChartWithAxisDataPoint,
  GaugeChartSegment,
  HeatMapChartData,
  GanttChartDataPoint,
} from '@fluentui/react-charts';
import { DataPieRegular } from '@fluentui/react-icons';
import { useMsal } from '@azure/msal-react';
import { useTheme } from '../../../contexts/ThemeContext';
import type { ChartWebPartConfig, AnyWebPartConfig } from '../../../types/page';
import {
  fetchChartWebPartData,
  fetchHeatmapData,
  fetchGanttData,
  CHART_COLORS,
  type ChartDataPoint,
  type HeatmapDataPoint,
  type GanttDataPoint,
} from '../../../services/webPartData';
import WebPartHeader from './WebPartHeader';
import WebPartSettingsDrawer from './WebPartSettingsDrawer';

const useStyles = makeStyles({
  container: {
    backgroundColor: tokens.colorNeutralBackground1,
    borderRadius: '2px',
    border: '1px solid transparent',
    borderImage: 'linear-gradient(135deg, rgba(0,0,0,0.05) 0%, rgba(0,0,0,0.15) 100%) 1',
    boxShadow: '0 1px 2px rgba(0, 0, 0, 0.08), 0 2px 4px rgba(0, 0, 0, 0.04)',
    overflow: 'hidden',
    display: 'flex',
    flexDirection: 'column',
    height: '100%',
  },
  containerDark: {
    backgroundColor: '#1a1a1a',
    borderImage: 'none',
    border: '1px solid #333333',
  },
  chartWrapper: {
    display: 'flex',
    justifyContent: 'center',
    alignItems: 'center',
    flex: 1,
    minHeight: 0,
    padding: '16px',
  },
  emptyState: {
    display: 'flex',
    flexDirection: 'column',
    alignItems: 'center',
    justifyContent: 'center',
    padding: '48px 24px',
    gap: '12px',
  },
  emptyIcon: {
    color: tokens.colorNeutralForeground3,
    fontSize: '32px',
  },
  emptyText: {
    color: tokens.colorNeutralForeground3,
    textAlign: 'center',
    fontSize: tokens.fontSizeBase200,
  },
  loadingContainer: {
    display: 'flex',
    alignItems: 'center',
    justifyContent: 'center',
    padding: '48px 24px',
  },
  errorText: {
    color: tokens.colorPaletteRedForeground1,
    padding: '16px',
    textAlign: 'center',
  },
});

interface ChartWebPartProps {
  config: ChartWebPartConfig;
  onConfigChange?: (config: AnyWebPartConfig) => void;
}

export default function ChartWebPart({ config, onConfigChange }: ChartWebPartProps) {
  const { theme } = useTheme();
  const styles = useStyles();
  const { instance, accounts } = useMsal();
  const account = accounts[0];

  const [dataPoints, setDataPoints] = useState<ChartDataPoint[]>([]);
  const [heatmapData, setHeatmapData] = useState<HeatmapDataPoint[]>([]);
  const [ganttData, setGanttData] = useState<GanttDataPoint[]>([]);
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState<string | null>(null);
  const [settingsOpen, setSettingsOpen] = useState(false);

  const chartType = config.chartType || 'donut';

  // Different chart types have different configuration requirements
  const isConfigured = Boolean(
    config.dataSource?.siteId &&
    config.dataSource?.listId &&
    (chartType === 'gantt'
      ? config.ganttStartColumn && config.ganttEndColumn
      : chartType === 'heatmap'
        ? config.groupByColumn && config.secondaryGroupByColumn
        : config.groupByColumn)
  );

  // Load data when config changes
  useEffect(() => {
    async function loadData() {
      if (!isConfigured || !account) {
        setDataPoints([]);
        setHeatmapData([]);
        setGanttData([]);
        return;
      }

      setLoading(true);
      setError(null);

      try {
        if (chartType === 'heatmap') {
          const data = await fetchHeatmapData(instance, account, config);
          setHeatmapData(data);
          setDataPoints([]);
          setGanttData([]);
        } else if (chartType === 'gantt') {
          const data = await fetchGanttData(instance, account, config);
          setGanttData(data);
          setDataPoints([]);
          setHeatmapData([]);
        } else {
          const points = await fetchChartWebPartData(instance, account, config);
          setDataPoints(points);
          setHeatmapData([]);
          setGanttData([]);
        }
      } catch (err) {
        console.error('Failed to load chart data:', err);
        setError(err instanceof Error ? err.message : 'Failed to load data');
      } finally {
        setLoading(false);
      }
    }

    loadData();
  }, [config, instance, account, isConfigured, chartType]);

  const handleSettingsClick = useCallback(() => {
    setSettingsOpen(true);
  }, []);

  const handleSettingsSave = useCallback(
    (updatedConfig: AnyWebPartConfig) => {
      onConfigChange?.(updatedConfig);
      setSettingsOpen(false);
    },
    [onConfigChange]
  );

  // Convert data points to chart format
  const donutChartData: ChartProps = useMemo(() => ({
    chartTitle: config.title || 'Chart',
    chartData: dataPoints.map((dp) => ({
      legend: dp.legend,
      data: dp.data,
      color: dp.color,
    })),
  }), [dataPoints, config.title]);

  // Generate default legend label based on aggregation and value column
  const defaultLegendLabel = useMemo(() => {
    const agg = config.aggregation || 'count';
    const valueCol = config.valueColumn;
    if (agg === 'count') return 'Count';
    if (valueCol) {
      const aggLabel = agg.charAt(0).toUpperCase() + agg.slice(1);
      return `${aggLabel} of ${valueCol}`;
    }
    return 'Value';
  }, [config.aggregation, config.valueColumn]);

  const chartLegendLabel = config.legendLabel || defaultLegendLabel;

  // Single color for bar charts (first color from palette)
  const singleBarColor = dataPoints[0]?.color || '#56C596';

  const barChartData: VerticalBarChartDataPoint[] = useMemo(() =>
    dataPoints.map((dp) => ({
      x: dp.legend,
      y: dp.data,
      color: singleBarColor,
      legend: chartLegendLabel, // Single legend for all bars
    })),
  [dataPoints, chartLegendLabel, singleBarColor]);

  // For line charts, use Date objects for proper date axis, falling back to index for non-dates
  const lineChartData: LineChartPoints[] = useMemo(() => [{
    legend: chartLegendLabel,
    data: dataPoints.map((dp, index) => ({
      x: dp.sortKey !== undefined ? new Date(dp.sortKey) : index,
      y: dp.data,
      xAxisCalloutData: dp.legend, // Show formatted date in tooltip
    })),
    color: singleBarColor,
  }], [dataPoints, chartLegendLabel, singleBarColor]);

  // Custom x-axis tick values for line/area charts when using dates
  const xAxisTickValues = useMemo(() => {
    if (dataPoints.length === 0 || dataPoints[0].sortKey === undefined) return undefined;
    // Convert timestamps to Date objects for the chart
    return dataPoints.map((dp) => new Date(dp.sortKey!));
  }, [dataPoints]);

  // D3 date format string for x-axis labels (e.g., "Dec 18")
  const xAxisTickFormat = useMemo(() => {
    if (dataPoints.length === 0 || dataPoints[0].sortKey === undefined) return undefined;
    return '%b %d'; // Short month + day format
  }, [dataPoints]);

  const horizontalBarChartData: HorizontalBarChartWithAxisDataPoint[] = useMemo(() =>
    dataPoints.map((dp) => ({
      x: dp.data,
      y: dp.legend,
      color: singleBarColor,
      legend: chartLegendLabel, // Single legend for all bars
    })),
  [dataPoints, chartLegendLabel, singleBarColor]);

  // Calculate total for donut center and gauge
  const total = useMemo(() =>
    dataPoints.reduce((sum, dp) => sum + dp.data, 0),
  [dataPoints]);

  // Area chart data (same format as line chart, uses Date objects for dates)
  const areaChartData: ChartProps = useMemo(() => ({
    chartTitle: config.title || 'Chart',
    chartData: dataPoints.map((dp) => ({
      legend: dp.legend,
      data: dp.data,
      color: dp.color,
    })),
    lineChartData: [{
      legend: chartLegendLabel,
      data: dataPoints.map((dp, index) => ({
        x: dp.sortKey !== undefined ? new Date(dp.sortKey) : index,
        y: dp.data,
        xAxisCalloutData: dp.legend,
      })),
      color: singleBarColor,
    }],
  }), [dataPoints, config.title, chartLegendLabel, singleBarColor]);

  // Gauge chart segments - distribute segments based on data points or show single value
  const gaugeSegments: GaugeChartSegment[] = useMemo(() => {
    const maxValue = config.gaugeMaxValue ?? 100;
    // If we have data points, use them as segments; otherwise create a simple progress gauge
    if (dataPoints.length > 0) {
      return dataPoints.map((dp) => ({
        legend: dp.legend,
        size: dp.data,
        color: dp.color,
      }));
    }
    // Default empty segments
    return [{ legend: 'Value', size: maxValue, color: CHART_COLORS.data[0] }];
  }, [dataPoints, config.gaugeMaxValue]);

  const gaugeValue = total;
  const gaugeMinValue = config.gaugeMinValue ?? 0;
  const gaugeMaxValue = config.gaugeMaxValue ?? Math.max(total * 1.2, 100);

  // Heatmap chart data transformation
  const heatmapChartData: HeatMapChartData[] = useMemo(() => {
    if (heatmapData.length === 0) return [];

    // Group by Y value (legend)
    const groupedByY = new Map<string, typeof heatmapData>();
    for (const dp of heatmapData) {
      if (!groupedByY.has(dp.y)) {
        groupedByY.set(dp.y, []);
      }
      groupedByY.get(dp.y)!.push(dp);
    }

    // Convert to HeatMapChartData format
    const result: HeatMapChartData[] = [];
    for (const [yValue, points] of groupedByY.entries()) {
      result.push({
        legend: yValue,
        data: points.map((p) => ({
          x: p.x,
          y: yValue,
          value: p.value,
          rectText: p.rectText,
        })),
        value: points.reduce((sum, p) => sum + p.value, 0) / points.length, // avg for color scale
      });
    }
    return result;
  }, [heatmapData]);

  // Heatmap color scale
  const heatmapDomain = useMemo(() => {
    if (heatmapData.length === 0) return [0, 50, 100];
    const values = heatmapData.map((d) => d.value);
    const min = Math.min(...values);
    const max = Math.max(...values);
    const mid = (min + max) / 2;
    return [min, mid, max];
  }, [heatmapData]);

  const heatmapColors = [CHART_COLORS.minimum, CHART_COLORS.center, CHART_COLORS.maximum];

  // Scatter chart data (uses same structure as line/area)
  const scatterChartData: ChartProps = useMemo(() => ({
    chartTitle: config.title || 'Chart',
    chartData: dataPoints.map((dp, index) => ({
      legend: dp.legend,
      data: dp.data,
      color: dp.color,
      xAxisCalloutData: dp.legend,
      yAxisCalloutData: String(dp.data),
      // For scatter, x is index, y is value
      x: index,
      y: dp.data,
    })),
  }), [dataPoints, config.title]);

  // Gantt chart data transformation
  const ganttChartData: GanttChartDataPoint[] = useMemo(() => {
    return ganttData.map((dp) => ({
      x: {
        start: dp.start,
        end: dp.end,
      },
      y: dp.label,
      legend: dp.label,
      color: dp.color,
    }));
  }, [ganttData]);

  // Check if we have any data to display
  const hasData = chartType === 'heatmap'
    ? heatmapData.length > 0
    : chartType === 'gantt'
      ? ganttData.length > 0
      : dataPoints.length > 0;

  const renderChart = () => {
    if (!hasData) return null;

    const chartType = config.chartType || 'donut';
    const legendPosition = config.legendPosition || 'on';
    const hideLegend = legendPosition === 'off';
    const useAngledLabels = config.xAxisLabelStyle === 'angled';
    // Use negative bottom margin to reclaim space from rotated labels
    const barChartMargins = useAngledLabels
      ? { top: 20, bottom: 20, left: 20, right: 20 }
      : { top: 20, bottom: 25, left: 20, right: 20 };
    // H-Bar needs extra left padding for Y-axis labels
    const hBarChartMargins = { top: 20, bottom: 20, left: 90, right: 20 };

    switch (chartType) {
      case 'bar':
        return (
          <VerticalBarChart
            key={`bar-${legendPosition}-${useAngledLabels}`}
            data={barChartData}
            height={250}
            width={350}
            hideLegend={hideLegend}
            rotateXAxisLables={useAngledLabels}
            margins={barChartMargins}
          />
        );
      case 'horizontal-bar':
        return (
          <HorizontalBarChartWithAxis
            key={`hbar-${legendPosition}`}
            data={horizontalBarChartData}
            height={250}
            width={350}
            hideLegend={hideLegend}
            margins={hBarChartMargins}
          />
        );
      case 'line':
        return (
          <LineChart
            key={`line-${legendPosition}`}
            data={{ lineChartData: lineChartData }}
            height={250}
            width={350}
            hideLegend={hideLegend}
            tickValues={xAxisTickValues}
            tickFormat={xAxisTickFormat}
          />
        );
      case 'area':
        return (
          <AreaChart
            key={`area-${legendPosition}`}
            data={areaChartData}
            height={250}
            width={350}
            hideLegend={hideLegend}
            tickValues={xAxisTickValues}
            tickFormat={xAxisTickFormat}
          />
        );
      case 'gauge':
        return (
          <GaugeChart
            key={`gauge-${legendPosition}`}
            segments={gaugeSegments}
            chartValue={gaugeValue}
            minValue={gaugeMinValue}
            maxValue={gaugeMaxValue}
            height={250}
            width={350}
            chartTitle={config.title}
          />
        );
      case 'heatmap':
        return (
          <HeatMapChart
            key={`heatmap-${legendPosition}`}
            data={heatmapChartData}
            domainValuesForColorScale={heatmapDomain}
            rangeValuesForColorScale={heatmapColors}
            height={250}
            width={350}
            hideLegend={hideLegend}
          />
        );
      case 'scatter':
        return (
          <ScatterChart
            key={`scatter-${legendPosition}`}
            data={scatterChartData}
            height={250}
            width={350}
            hideLegend={hideLegend}
          />
        );
      case 'gantt':
        return (
          <GanttChart
            key={`gantt-${legendPosition}`}
            data={ganttChartData}
            height={250}
            width={350}
            hideLegend={hideLegend}
          />
        );
      case 'donut':
      default:
        return (
          <DonutChart
            key={`donut-${legendPosition}`}
            data={donutChartData}
            innerRadius={55}
            height={250}
            width={300}
            valueInsideDonut={String(Math.round(total))}
            hideLegend={hideLegend}
          />
        );
    }
  };

  return (
    <div className={mergeClasses(styles.container, theme === 'dark' && styles.containerDark)}>
      <WebPartHeader
        title={config.title}
        isConfigured={isConfigured}
        onSettingsClick={handleSettingsClick}
      />

      {/* Loading state */}
      {loading && (
        <div className={styles.loadingContainer}>
          <Spinner size="small" label="Loading chart data..." />
        </div>
      )}

      {/* Error state */}
      {error && !loading && <Text className={styles.errorText}>{error}</Text>}

      {/* Empty/Not configured state */}
      {!loading && !error && !isConfigured && (
        <div className={styles.emptyState}>
          <DataPieRegular className={styles.emptyIcon} />
          <Text className={styles.emptyText}>
            Click the settings icon to configure this chart
          </Text>
        </div>
      )}

      {/* No data state */}
      {!loading && !error && isConfigured && !hasData && (
        <div className={styles.emptyState}>
          <Text className={styles.emptyText}>No data available for this chart</Text>
        </div>
      )}

      {/* Chart */}
      {!loading && !error && isConfigured && hasData && (
        <div className={styles.chartWrapper}>
          {renderChart()}
        </div>
      )}

      {/* Settings Drawer */}
      <WebPartSettingsDrawer
        webPart={config}
        open={settingsOpen}
        onClose={() => setSettingsOpen(false)}
        onSave={handleSettingsSave}
      />
    </div>
  );
}
