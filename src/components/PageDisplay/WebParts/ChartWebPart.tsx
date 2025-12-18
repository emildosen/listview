import { useState, useEffect, useCallback, useMemo } from 'react';
import {
  makeStyles,
  tokens,
  Text,
  Spinner,
  mergeClasses,
} from '@fluentui/react-components';
import { DonutChart, VerticalBarChart, LineChart } from '@fluentui/react-charts';
import type { ChartProps, VerticalBarChartDataPoint, LineChartPoints } from '@fluentui/react-charts';
import { DataPieRegular } from '@fluentui/react-icons';
import { useMsal } from '@azure/msal-react';
import { useTheme } from '../../../contexts/ThemeContext';
import type { ChartWebPartConfig, AnyWebPartConfig } from '../../../types/page';
import { fetchChartWebPartData, type ChartDataPoint } from '../../../services/webPartData';
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
    minHeight: '250px',
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
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState<string | null>(null);
  const [settingsOpen, setSettingsOpen] = useState(false);

  const isConfigured = Boolean(
    config.dataSource?.siteId &&
    config.dataSource?.listId &&
    config.groupByColumn
  );

  // Load data when config changes
  useEffect(() => {
    async function loadData() {
      if (!isConfigured || !account) {
        setDataPoints([]);
        return;
      }

      setLoading(true);
      setError(null);

      try {
        const points = await fetchChartWebPartData(instance, account, config);
        setDataPoints(points);
      } catch (err) {
        console.error('Failed to load chart data:', err);
        setError(err instanceof Error ? err.message : 'Failed to load data');
      } finally {
        setLoading(false);
      }
    }

    loadData();
  }, [config, instance, account, isConfigured]);

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

  const barChartData: VerticalBarChartDataPoint[] = useMemo(() =>
    dataPoints.map((dp) => ({
      x: dp.legend,
      y: dp.data,
      color: dp.color,
    })),
  [dataPoints]);

  const lineChartData: LineChartPoints[] = useMemo(() => [{
    legend: config.title || 'Data',
    data: dataPoints.map((dp, index) => ({
      x: index,
      y: dp.data,
    })),
    color: dataPoints[0]?.color || '#0078d4',
  }], [dataPoints, config.title]);

  // Calculate total for donut center
  const total = useMemo(() =>
    dataPoints.reduce((sum, dp) => sum + dp.data, 0),
  [dataPoints]);

  const renderChart = () => {
    if (dataPoints.length === 0) return null;

    const chartType = config.chartType || 'donut';

    switch (chartType) {
      case 'bar':
        return (
          <VerticalBarChart
            data={barChartData}
            height={250}
            width={350}
            hideLegend={!config.showLegend}
          />
        );
      case 'horizontal-bar':
        // Using vertical bar as fallback since HorizontalBarChart has different API
        return (
          <VerticalBarChart
            data={barChartData}
            height={250}
            width={350}
            hideLegend={!config.showLegend}
          />
        );
      case 'line':
        return (
          <LineChart
            data={{ lineChartData: lineChartData }}
            height={250}
            width={350}
            hideLegend={!config.showLegend}
          />
        );
      case 'donut':
      default:
        return (
          <DonutChart
            data={donutChartData}
            innerRadius={55}
            height={250}
            width={300}
            valueInsideDonut={String(Math.round(total))}
            hideLegend={!config.showLegend}
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
      {!loading && !error && isConfigured && dataPoints.length === 0 && (
        <div className={styles.emptyState}>
          <Text className={styles.emptyText}>No data available for this chart</Text>
        </div>
      )}

      {/* Chart */}
      {!loading && !error && isConfigured && dataPoints.length > 0 && (
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
