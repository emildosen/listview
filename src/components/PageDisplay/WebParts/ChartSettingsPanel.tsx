import {
  makeStyles,
  tokens,
  Dropdown,
  Option,
  Field,
  Switch,
  SpinButton,
  ToggleButton,
} from '@fluentui/react-components';
import {
  DataBarVerticalRegular,
  DataPieRegular,
  DataLineRegular,
  DataBarHorizontalRegular,
} from '@fluentui/react-icons';
import type { ChartAggregation } from '../../../types/page';
import type { GraphListColumn } from '../../../auth/graphClient';

const useStyles = makeStyles({
  container: {
    display: 'flex',
    flexDirection: 'column',
    gap: '16px',
  },
  chartTypeGroup: {
    display: 'flex',
    gap: '8px',
    flexWrap: 'wrap',
  },
  chartTypeButton: {
    display: 'flex',
    flexDirection: 'column',
    alignItems: 'center',
    gap: '4px',
    padding: '12px 16px',
    minWidth: '70px',
  },
  chartTypeIcon: {
    fontSize: '24px',
  },
  chartTypeLabel: {
    fontSize: tokens.fontSizeBase100,
  },
  fieldRow: {
    display: 'flex',
    gap: '16px',
    flexWrap: 'wrap',
  },
  fieldHalf: {
    flex: 1,
    minWidth: '120px',
  },
  displayOptionsRow: {
    display: 'flex',
    gap: '16px',
    flexWrap: 'wrap',
    alignItems: 'flex-end',
  },
  spinButtonField: {
    maxWidth: '120px',
  },
});

const CHART_TYPES = [
  { value: 'bar' as const, label: 'Bar', icon: <DataBarVerticalRegular /> },
  { value: 'donut' as const, label: 'Donut', icon: <DataPieRegular /> },
  { value: 'line' as const, label: 'Line', icon: <DataLineRegular /> },
  { value: 'horizontal-bar' as const, label: 'H-Bar', icon: <DataBarHorizontalRegular /> },
];

const AGGREGATION_OPTIONS: { value: ChartAggregation; label: string }[] = [
  { value: 'count', label: 'Count' },
  { value: 'sum', label: 'Sum' },
  { value: 'average', label: 'Average' },
  { value: 'min', label: 'Minimum' },
  { value: 'max', label: 'Maximum' },
];

function isNumericColumn(column: GraphListColumn): boolean {
  return !!column.number;
}

interface ChartSettingsPanelProps {
  chartType: 'bar' | 'donut' | 'line' | 'horizontal-bar';
  groupByColumn?: string;
  valueColumn?: string;
  aggregation: ChartAggregation;
  showLegend: boolean;
  maxGroups: number;
  columns: GraphListColumn[];
  onChartTypeChange: (type: 'bar' | 'donut' | 'line' | 'horizontal-bar') => void;
  onGroupByColumnChange: (column: string | undefined) => void;
  onValueColumnChange: (column: string | undefined) => void;
  onAggregationChange: (agg: ChartAggregation) => void;
  onShowLegendChange: (show: boolean) => void;
  onMaxGroupsChange: (max: number) => void;
}

export default function ChartSettingsPanel({
  chartType,
  groupByColumn,
  valueColumn,
  aggregation,
  showLegend,
  maxGroups,
  columns,
  onChartTypeChange,
  onGroupByColumnChange,
  onValueColumnChange,
  onAggregationChange,
  onShowLegendChange,
  onMaxGroupsChange,
}: ChartSettingsPanelProps) {
  const styles = useStyles();

  // Get columns suitable for grouping (choice, text, lookup, boolean)
  const groupableColumns = columns.filter(
    (col) => col.choice || col.text || col.lookup || col.boolean
  );

  // Get numeric columns for value aggregation
  const numericColumns = columns.filter(isNumericColumn);

  // For count aggregation, we don't need a value column
  const needsValueColumn = aggregation !== 'count';

  const selectedGroupColumn = columns.find((c) => c.name === groupByColumn);
  const selectedValueColumn = columns.find((c) => c.name === valueColumn);

  return (
    <div className={styles.container}>
      {/* Chart Type */}
      <Field label="Chart Type">
        <div className={styles.chartTypeGroup}>
          {CHART_TYPES.map((type) => (
            <ToggleButton
              key={type.value}
              className={styles.chartTypeButton}
              appearance="subtle"
              checked={chartType === type.value}
              onClick={() => onChartTypeChange(type.value)}
            >
              <span className={styles.chartTypeIcon}>{type.icon}</span>
              <span className={styles.chartTypeLabel}>{type.label}</span>
            </ToggleButton>
          ))}
        </div>
      </Field>

      {/* Group By Column */}
      <div className={styles.fieldRow}>
        <Field label="Group by" className={styles.fieldHalf}>
          <Dropdown
            placeholder="Select column"
            value={selectedGroupColumn?.displayName || ''}
            selectedOptions={groupByColumn ? [groupByColumn] : []}
            onOptionSelect={(_, data) => onGroupByColumnChange(data.optionValue as string)}
          >
            {groupableColumns.map((col) => (
              <Option key={col.name} value={col.name}>
                {col.displayName}
              </Option>
            ))}
          </Dropdown>
        </Field>

        {/* Aggregation */}
        <Field label="Aggregation" className={styles.fieldHalf}>
          <Dropdown
            value={AGGREGATION_OPTIONS.find((a) => a.value === aggregation)?.label || ''}
            selectedOptions={[aggregation]}
            onOptionSelect={(_, data) => onAggregationChange(data.optionValue as ChartAggregation)}
          >
            {AGGREGATION_OPTIONS.map((opt) => (
              <Option key={opt.value} value={opt.value}>
                {opt.label}
              </Option>
            ))}
          </Dropdown>
        </Field>
      </div>

      {/* Value Column (only for non-count aggregations) */}
      {needsValueColumn && (
        <Field label="Value column">
          <Dropdown
            placeholder="Select numeric column"
            value={selectedValueColumn?.displayName || ''}
            selectedOptions={valueColumn ? [valueColumn] : []}
            onOptionSelect={(_, data) => onValueColumnChange(data.optionValue as string)}
            disabled={numericColumns.length === 0}
          >
            {numericColumns.length === 0 ? (
              <Option value="" disabled>
                No numeric columns available
              </Option>
            ) : (
              numericColumns.map((col) => (
                <Option key={col.name} value={col.name}>
                  {col.displayName}
                </Option>
              ))
            )}
          </Dropdown>
        </Field>
      )}

      {/* Display Options */}
      <div className={styles.displayOptionsRow}>
        <Field label="Max groups" className={styles.spinButtonField}>
          <SpinButton
            value={maxGroups}
            min={1}
            max={50}
            onChange={(_, data) => onMaxGroupsChange(data.value || 10)}
          />
        </Field>
        <Field label="Show legend">
          <Switch checked={showLegend} onChange={(_, data) => onShowLegendChange(data.checked)} />
        </Field>
      </div>
    </div>
  );
}
