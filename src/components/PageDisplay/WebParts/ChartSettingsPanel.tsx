import {
  makeStyles,
  tokens,
  Dropdown,
  Option,
  Field,
  SpinButton,
  ToggleButton,
  Switch,
  Input,
} from '@fluentui/react-components';
import {
  DataBarVerticalRegular,
  DataPieRegular,
  DataLineRegular,
  DataBarHorizontalRegular,
  DataAreaRegular,
  GaugeRegular,
  GridDotsRegular,
  DataScatterRegular,
  GanttChartRegular,
} from '@fluentui/react-icons';
import type { ChartAggregation, ChartType, LegendPosition, XAxisLabelStyle } from '../../../types/page';
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

const CHART_TYPES: { value: ChartType; label: string; icon: React.ReactNode }[] = [
  { value: 'bar', label: 'Bar', icon: <DataBarVerticalRegular /> },
  { value: 'donut', label: 'Donut', icon: <DataPieRegular /> },
  { value: 'line', label: 'Line', icon: <DataLineRegular /> },
  { value: 'horizontal-bar', label: 'H-Bar', icon: <DataBarHorizontalRegular /> },
  { value: 'area', label: 'Area', icon: <DataAreaRegular /> },
  { value: 'gauge', label: 'Gauge', icon: <GaugeRegular /> },
  { value: 'heatmap', label: 'Heatmap', icon: <GridDotsRegular /> },
  { value: 'scatter', label: 'Scatter', icon: <DataScatterRegular /> },
  { value: 'gantt', label: 'Gantt', icon: <GanttChartRegular /> },
];

const AGGREGATION_OPTIONS: { value: ChartAggregation; label: string }[] = [
  { value: 'count', label: 'Count' },
  { value: 'sum', label: 'Sum' },
  { value: 'average', label: 'Average' },
  { value: 'min', label: 'Minimum' },
  { value: 'max', label: 'Maximum' },
];

function isNumericColumn(column: GraphListColumn): boolean {
  // Boolean columns can be aggregated as 1/0 (Yes=1, No=0)
  return !!column.number || !!column.boolean;
}

interface ChartSettingsPanelProps {
  chartType: ChartType;
  groupByColumn?: string;
  valueColumn?: string;
  aggregation: ChartAggregation;
  legendPosition: LegendPosition;
  legendLabel?: string;
  maxGroups: number;
  showOther: boolean;
  includeNull: boolean;
  xAxisLabelStyle: XAxisLabelStyle;
  sortBy: 'label' | 'value';
  sortDirection: 'asc' | 'desc';
  columns: GraphListColumn[];
  // Gauge chart options
  gaugeMinValue?: number;
  gaugeMaxValue?: number;
  // Heatmap chart options
  secondaryGroupByColumn?: string;
  // Gantt chart options
  ganttStartColumn?: string;
  ganttEndColumn?: string;
  ganttLabelColumn?: string;
  // Callbacks
  onChartTypeChange: (type: ChartType) => void;
  onGroupByColumnChange: (column: string | undefined) => void;
  onValueColumnChange: (column: string | undefined) => void;
  onAggregationChange: (agg: ChartAggregation) => void;
  onLegendPositionChange: (position: LegendPosition) => void;
  onLegendLabelChange: (label: string) => void;
  onMaxGroupsChange: (max: number) => void;
  onShowOtherChange: (show: boolean) => void;
  onIncludeNullChange: (include: boolean) => void;
  onXAxisLabelStyleChange: (style: XAxisLabelStyle) => void;
  onSortByChange: (sortBy: 'label' | 'value') => void;
  onSortDirectionChange: (direction: 'asc' | 'desc') => void;
  onGaugeMinValueChange?: (value: number) => void;
  onGaugeMaxValueChange?: (value: number) => void;
  onSecondaryGroupByColumnChange?: (column: string | undefined) => void;
  onGanttStartColumnChange?: (column: string | undefined) => void;
  onGanttEndColumnChange?: (column: string | undefined) => void;
  onGanttLabelColumnChange?: (column: string | undefined) => void;
}

export default function ChartSettingsPanel({
  chartType,
  groupByColumn,
  valueColumn,
  aggregation,
  legendPosition,
  legendLabel,
  maxGroups,
  showOther,
  includeNull,
  xAxisLabelStyle,
  sortBy,
  sortDirection,
  columns,
  // Gauge options
  gaugeMinValue,
  gaugeMaxValue,
  // Heatmap options
  secondaryGroupByColumn,
  // Gantt options
  ganttStartColumn,
  ganttEndColumn,
  ganttLabelColumn,
  // Callbacks
  onChartTypeChange,
  onGroupByColumnChange,
  onValueColumnChange,
  onAggregationChange,
  onLegendPositionChange,
  onLegendLabelChange,
  onMaxGroupsChange,
  onShowOtherChange,
  onIncludeNullChange,
  onXAxisLabelStyleChange,
  onSortByChange,
  onSortDirectionChange,
  onGaugeMinValueChange,
  onGaugeMaxValueChange,
  onSecondaryGroupByColumnChange,
  onGanttStartColumnChange,
  onGanttEndColumnChange,
  onGanttLabelColumnChange,
}: ChartSettingsPanelProps) {
  const styles = useStyles();

  // Get columns suitable for grouping (choice, text, lookup, boolean, dateTime)
  const groupableColumns = columns.filter(
    (col) => col.choice || col.text || col.lookup || col.boolean || col.dateTime
  );

  // Get numeric columns for value aggregation
  const numericColumns = columns.filter(isNumericColumn);

  // Get date columns for Gantt chart
  const dateColumns = columns.filter((col) => col.dateTime);

  // For count aggregation, we don't need a value column
  const needsValueColumn = aggregation !== 'count';

  // Gauge chart shows aggregated total, not individual groups
  const isGaugeChart = chartType === 'gauge';
  // Heatmap needs two grouping dimensions
  const isHeatmapChart = chartType === 'heatmap';
  // Gantt chart needs date columns
  const isGanttChart = chartType === 'gantt';

  const selectedGroupColumn = columns.find((c) => c.name === groupByColumn);
  const selectedValueColumn = columns.find((c) => c.name === valueColumn);
  const selectedSecondaryGroupColumn = columns.find((c) => c.name === secondaryGroupByColumn);
  const selectedGanttStartColumn = columns.find((c) => c.name === ganttStartColumn);
  const selectedGanttEndColumn = columns.find((c) => c.name === ganttEndColumn);
  const selectedGanttLabelColumn = columns.find((c) => c.name === ganttLabelColumn);

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

      {/* X-Axis Label Style (only for vertical bar charts) */}
      {chartType === 'bar' && (
        <Field label="X-Axis labels">
          <Dropdown
            value={xAxisLabelStyle === 'angled' ? 'Angled (show all)' : 'Normal'}
            selectedOptions={[xAxisLabelStyle]}
            onOptionSelect={(_, data) => onXAxisLabelStyleChange(data.optionValue as XAxisLabelStyle)}
          >
            <Option value="normal">Normal</Option>
            <Option value="angled">Angled (show all)</Option>
          </Dropdown>
        </Field>
      )}

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

      {/* Value Column (only for non-count aggregations, not for Gantt) */}
      {needsValueColumn && !isGanttChart && (
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

      {/* Gauge Chart Settings */}
      {isGaugeChart && (
        <div className={styles.fieldRow}>
          <Field label="Min value" className={styles.fieldHalf}>
            <SpinButton
              value={gaugeMinValue ?? 0}
              min={0}
              onChange={(_, data) => onGaugeMinValueChange?.(data.value ?? 0)}
            />
          </Field>
          <Field label="Max value" className={styles.fieldHalf}>
            <SpinButton
              value={gaugeMaxValue ?? 100}
              min={1}
              onChange={(_, data) => onGaugeMaxValueChange?.(data.value ?? 100)}
            />
          </Field>
        </div>
      )}

      {/* Heatmap Chart Settings */}
      {isHeatmapChart && (
        <Field label="Y-Axis (secondary group)">
          <Dropdown
            placeholder="Select column"
            value={selectedSecondaryGroupColumn?.displayName || ''}
            selectedOptions={secondaryGroupByColumn ? [secondaryGroupByColumn] : []}
            onOptionSelect={(_, data) => onSecondaryGroupByColumnChange?.(data.optionValue as string)}
          >
            {groupableColumns.map((col) => (
              <Option key={col.name} value={col.name}>
                {col.displayName}
              </Option>
            ))}
          </Dropdown>
        </Field>
      )}

      {/* Gantt Chart Settings */}
      {isGanttChart && (
        <>
          <div className={styles.fieldRow}>
            <Field label="Start date" className={styles.fieldHalf}>
              <Dropdown
                placeholder="Select date column"
                value={selectedGanttStartColumn?.displayName || ''}
                selectedOptions={ganttStartColumn ? [ganttStartColumn] : []}
                onOptionSelect={(_, data) => onGanttStartColumnChange?.(data.optionValue as string)}
                disabled={dateColumns.length === 0}
              >
                {dateColumns.length === 0 ? (
                  <Option value="" disabled>
                    No date columns available
                  </Option>
                ) : (
                  dateColumns.map((col) => (
                    <Option key={col.name} value={col.name}>
                      {col.displayName}
                    </Option>
                  ))
                )}
              </Dropdown>
            </Field>
            <Field label="End date" className={styles.fieldHalf}>
              <Dropdown
                placeholder="Select date column"
                value={selectedGanttEndColumn?.displayName || ''}
                selectedOptions={ganttEndColumn ? [ganttEndColumn] : []}
                onOptionSelect={(_, data) => onGanttEndColumnChange?.(data.optionValue as string)}
                disabled={dateColumns.length === 0}
              >
                {dateColumns.length === 0 ? (
                  <Option value="" disabled>
                    No date columns available
                  </Option>
                ) : (
                  dateColumns.map((col) => (
                    <Option key={col.name} value={col.name}>
                      {col.displayName}
                    </Option>
                  ))
                )}
              </Dropdown>
            </Field>
          </div>
          <Field label="Label column">
            <Dropdown
              placeholder="Select column for bar labels"
              value={selectedGanttLabelColumn?.displayName || ''}
              selectedOptions={ganttLabelColumn ? [ganttLabelColumn] : []}
              onOptionSelect={(_, data) => onGanttLabelColumnChange?.(data.optionValue as string)}
            >
              {groupableColumns.map((col) => (
                <Option key={col.name} value={col.name}>
                  {col.displayName}
                </Option>
              ))}
            </Dropdown>
          </Field>
        </>
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
        <Field label="Show 'Other'">
          <Switch
            checked={showOther}
            onChange={(_, data) => onShowOtherChange(data.checked)}
          />
        </Field>
        <Field label="Fill gaps">
          <Switch
            checked={includeNull}
            onChange={(_, data) => onIncludeNullChange(data.checked)}
          />
        </Field>
        <Field label="Legend">
          <Switch
            checked={legendPosition === 'on'}
            onChange={(_, data) => onLegendPositionChange(data.checked ? 'on' : 'off')}
          />
        </Field>
      </div>

      {/* Legend Label (when legend is on) */}
      {legendPosition === 'on' && (
        <Field label="Legend label">
          <Input
            placeholder="e.g., Attendance Count"
            value={legendLabel || ''}
            onChange={(_, data) => onLegendLabelChange(data.value)}
          />
        </Field>
      )}

      {/* Sort Options */}
      <div className={styles.fieldRow}>
        <Field label="Sort by" className={styles.fieldHalf}>
          <Dropdown
            value={sortBy === 'value' ? 'Value' : 'Label'}
            selectedOptions={[sortBy]}
            onOptionSelect={(_, data) => onSortByChange(data.optionValue as 'label' | 'value')}
          >
            <Option value="label">Label</Option>
            <Option value="value">Value</Option>
          </Dropdown>
        </Field>
        <Field label="Direction" className={styles.fieldHalf}>
          <Dropdown
            value={sortDirection === 'desc' ? 'Descending' : 'Ascending'}
            selectedOptions={[sortDirection]}
            onOptionSelect={(_, data) => onSortDirectionChange(data.optionValue as 'asc' | 'desc')}
          >
            <Option value="asc">Ascending</Option>
            <Option value="desc">Descending</Option>
          </Dropdown>
        </Field>
      </div>
    </div>
  );
}
