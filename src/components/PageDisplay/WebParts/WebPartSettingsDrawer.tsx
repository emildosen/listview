import { useState, useCallback, useEffect } from 'react';
import {
  makeStyles,
  tokens,
  Text,
  Button,
  Divider,
  DrawerBody,
  DrawerHeader,
  DrawerHeaderTitle,
  OverlayDrawer,
  Input,
  Field,
  Switch,
  SpinButton,
} from '@fluentui/react-components';
import { DismissRegular } from '@fluentui/react-icons';
import { useMsal } from '@azure/msal-react';
import type {
  AnyWebPartConfig,
  ListItemsWebPartConfig,
  ChartWebPartConfig,
  WebPartDataSource,
  WebPartDisplayColumn,
  WebPartFilter,
  WebPartJoin,
  WebPartSort,
  ChartAggregation,
} from '../../../types/page';
import type { GraphListColumn } from '../../../auth/graphClient';
import { getListColumns } from '../../../auth/graphClient';
import DataSourcePicker from './DataSourcePicker';
import ColumnSelector from './ColumnSelector';
import FilterBuilder from './FilterBuilder';
import ChartSettingsPanel from './ChartSettingsPanel';
import JoinBuilder from './JoinBuilder';

const useStyles = makeStyles({
  body: {
    display: 'flex',
    flexDirection: 'column',
    height: '100%',
  },
  content: {
    flex: 1,
    overflowY: 'auto',
    paddingBottom: '16px',
  },
  section: {
    display: 'flex',
    flexDirection: 'column',
    gap: '12px',
    marginBottom: '16px',
  },
  sectionTitle: {
    fontWeight: tokens.fontWeightSemibold,
    fontSize: tokens.fontSizeBase200,
    textTransform: 'uppercase',
    letterSpacing: '0.02em',
    color: tokens.colorNeutralForeground2,
    marginBottom: '4px',
  },
  footer: {
    display: 'flex',
    justifyContent: 'flex-end',
    gap: '8px',
    padding: '16px 0',
    borderTop: `1px solid ${tokens.colorNeutralStroke1}`,
    marginTop: 'auto',
  },
  displayOptionsRow: {
    display: 'flex',
    gap: '16px',
    flexWrap: 'wrap',
  },
  spinButtonField: {
    maxWidth: '120px',
  },
});

interface WebPartSettingsDrawerProps {
  webPart: AnyWebPartConfig;
  open: boolean;
  onClose: () => void;
  onSave: (config: AnyWebPartConfig) => void;
}

export default function WebPartSettingsDrawer({
  webPart,
  open,
  onClose,
  onSave,
}: WebPartSettingsDrawerProps) {
  const styles = useStyles();
  const { instance, accounts } = useMsal();
  const account = accounts[0];

  // Local state for editing
  const [title, setTitle] = useState(webPart.title || '');
  const [dataSource, setDataSource] = useState<WebPartDataSource | undefined>(
    webPart.type === 'list-items'
      ? (webPart as ListItemsWebPartConfig).dataSource
      : (webPart as ChartWebPartConfig).dataSource
  );
  const [columns, setColumns] = useState<GraphListColumn[]>([]);
  const [loadingColumns, setLoadingColumns] = useState(false);

  // List Items specific state
  const [displayColumns, setDisplayColumns] = useState<WebPartDisplayColumn[]>(
    webPart.type === 'list-items' ? (webPart as ListItemsWebPartConfig).displayColumns || [] : []
  );
  const [filters, setFilters] = useState<WebPartFilter[]>(
    webPart.type === 'list-items'
      ? (webPart as ListItemsWebPartConfig).filters || []
      : (webPart as ChartWebPartConfig).filters || []
  );
  const [sort, setSort] = useState<WebPartSort | undefined>(
    webPart.type === 'list-items' ? (webPart as ListItemsWebPartConfig).sort : undefined
  );
  const [maxItems, setMaxItems] = useState(
    webPart.type === 'list-items' ? (webPart as ListItemsWebPartConfig).maxItems || 50 : 50
  );
  const [showSearch, setShowSearch] = useState(
    webPart.type === 'list-items' ? (webPart as ListItemsWebPartConfig).showSearch || false : false
  );
  const [joins, setJoins] = useState<WebPartJoin[]>(
    webPart.type === 'list-items'
      ? (webPart as ListItemsWebPartConfig).joins || []
      : (webPart as ChartWebPartConfig).joins || []
  );

  // Chart specific state
  const [chartType, setChartType] = useState<'bar' | 'donut' | 'line' | 'horizontal-bar'>(
    webPart.type === 'chart' ? (webPart as ChartWebPartConfig).chartType || 'bar' : 'bar'
  );
  const [groupByColumn, setGroupByColumn] = useState<string | undefined>(
    webPart.type === 'chart' ? (webPart as ChartWebPartConfig).groupByColumn : undefined
  );
  const [valueColumn, setValueColumn] = useState<string | undefined>(
    webPart.type === 'chart' ? (webPart as ChartWebPartConfig).valueColumn : undefined
  );
  const [aggregation, setAggregation] = useState<ChartAggregation>(
    webPart.type === 'chart' ? (webPart as ChartWebPartConfig).aggregation || 'count' : 'count'
  );
  const [showLegend, setShowLegend] = useState(
    webPart.type === 'chart' ? (webPart as ChartWebPartConfig).showLegend !== false : true
  );
  const [maxGroups, setMaxGroups] = useState(
    webPart.type === 'chart' ? (webPart as ChartWebPartConfig).maxGroups || 10 : 10
  );

  const [saving, setSaving] = useState(false);

  // Load columns when data source changes
  useEffect(() => {
    async function loadColumns() {
      if (!dataSource?.siteId || !dataSource?.listId || !account) {
        setColumns([]);
        return;
      }

      setLoadingColumns(true);
      try {
        const cols = await getListColumns(instance, account, dataSource.siteId, dataSource.listId);
        setColumns(cols);
      } catch (err) {
        console.error('Failed to load columns:', err);
        setColumns([]);
      } finally {
        setLoadingColumns(false);
      }
    }

    loadColumns();
  }, [dataSource?.siteId, dataSource?.listId, instance, account]);

  // Reset local state when webPart changes
  useEffect(() => {
    setTitle(webPart.title || '');
    if (webPart.type === 'list-items') {
      const config = webPart as ListItemsWebPartConfig;
      setDataSource(config.dataSource);
      setDisplayColumns(config.displayColumns || []);
      setFilters(config.filters || []);
      setJoins(config.joins || []);
      setSort(config.sort);
      setMaxItems(config.maxItems || 50);
      setShowSearch(config.showSearch || false);
    } else if (webPart.type === 'chart') {
      const config = webPart as ChartWebPartConfig;
      setDataSource(config.dataSource);
      setFilters(config.filters || []);
      setJoins(config.joins || []);
      setChartType(config.chartType || 'bar');
      setGroupByColumn(config.groupByColumn);
      setValueColumn(config.valueColumn);
      setAggregation(config.aggregation || 'count');
      setShowLegend(config.showLegend !== false);
      setMaxGroups(config.maxGroups || 10);
    }
  }, [webPart]);

  const handleSave = useCallback(async () => {
    setSaving(true);
    try {
      let updatedConfig: AnyWebPartConfig;

      if (webPart.type === 'list-items') {
        updatedConfig = {
          ...webPart,
          title,
          dataSource,
          displayColumns,
          filters,
          joins,
          sort,
          maxItems,
          showSearch,
        } as ListItemsWebPartConfig;
      } else {
        updatedConfig = {
          ...webPart,
          title,
          dataSource,
          filters,
          joins,
          chartType,
          groupByColumn,
          valueColumn,
          aggregation,
          showLegend,
          maxGroups,
        } as ChartWebPartConfig;
      }

      onSave(updatedConfig);
      onClose();
    } finally {
      setSaving(false);
    }
  }, [
    webPart,
    title,
    dataSource,
    displayColumns,
    filters,
    joins,
    sort,
    maxItems,
    showSearch,
    chartType,
    groupByColumn,
    valueColumn,
    aggregation,
    showLegend,
    maxGroups,
    onSave,
    onClose,
  ]);

  const drawerTitle =
    webPart.type === 'list-items' ? 'Configure List Items' : 'Configure Chart';

  return (
    <OverlayDrawer
      position="end"
      size="medium"
      open={open}
      onOpenChange={(_, { open: isOpen }) => {
        if (!isOpen) onClose();
      }}
    >
      <DrawerHeader>
        <DrawerHeaderTitle
          action={
            <Button
              appearance="subtle"
              aria-label="Close"
              icon={<DismissRegular />}
              onClick={onClose}
            />
          }
        >
          {drawerTitle}
        </DrawerHeaderTitle>
      </DrawerHeader>

      <DrawerBody className={styles.body}>
        <div className={styles.content}>
          {/* Title */}
          <div className={styles.section}>
            <Field label="Title">
              <Input
                value={title}
                onChange={(_, data) => setTitle(data.value)}
                placeholder="Enter title..."
              />
            </Field>
          </div>

          <Divider />

          {/* Data Source */}
          <div className={styles.section}>
            <Text className={styles.sectionTitle}>Data Source</Text>
            <DataSourcePicker value={dataSource} onChange={setDataSource} />
          </div>

          {dataSource && columns.length > 0 && (
            <>
              <Divider />

              {/* List Items specific settings */}
              {webPart.type === 'list-items' && (
                <>
                  {/* Columns */}
                  <div className={styles.section}>
                    <Text className={styles.sectionTitle}>Columns</Text>
                    <ColumnSelector
                      availableColumns={columns}
                      selectedColumns={displayColumns}
                      onChange={setDisplayColumns}
                      loading={loadingColumns}
                    />
                  </div>

                  <Divider />

                  {/* Filters */}
                  <div className={styles.section}>
                    <Text className={styles.sectionTitle}>Filters</Text>
                    <FilterBuilder
                      filters={filters}
                      columns={columns}
                      onChange={setFilters}
                    />
                  </div>

                  <Divider />

                  {/* Joins */}
                  <div className={styles.section}>
                    <Text className={styles.sectionTitle}>Joins</Text>
                    <JoinBuilder
                      joins={joins}
                      primaryColumns={columns}
                      onChange={setJoins}
                    />
                  </div>

                  <Divider />

                  {/* Display Options */}
                  <div className={styles.section}>
                    <Text className={styles.sectionTitle}>Display Options</Text>
                    <div className={styles.displayOptionsRow}>
                      <Field label="Max items" className={styles.spinButtonField}>
                        <SpinButton
                          value={maxItems}
                          min={1}
                          max={500}
                          onChange={(_, data) => setMaxItems(data.value || 50)}
                        />
                      </Field>
                      <Field label="Enable search">
                        <Switch
                          checked={showSearch}
                          onChange={(_, data) => setShowSearch(data.checked)}
                        />
                      </Field>
                    </div>
                  </div>
                </>
              )}

              {/* Chart specific settings */}
              {webPart.type === 'chart' && (
                <>
                  {/* Chart Settings */}
                  <div className={styles.section}>
                    <Text className={styles.sectionTitle}>Chart Settings</Text>
                    <ChartSettingsPanel
                      chartType={chartType}
                      groupByColumn={groupByColumn}
                      valueColumn={valueColumn}
                      aggregation={aggregation}
                      showLegend={showLegend}
                      maxGroups={maxGroups}
                      columns={columns}
                      onChartTypeChange={setChartType}
                      onGroupByColumnChange={setGroupByColumn}
                      onValueColumnChange={setValueColumn}
                      onAggregationChange={setAggregation}
                      onShowLegendChange={setShowLegend}
                      onMaxGroupsChange={setMaxGroups}
                    />
                  </div>

                  <Divider />

                  {/* Filters */}
                  <div className={styles.section}>
                    <Text className={styles.sectionTitle}>Filters</Text>
                    <FilterBuilder
                      filters={filters}
                      columns={columns}
                      onChange={setFilters}
                    />
                  </div>
                </>
              )}
            </>
          )}

          {dataSource && loadingColumns && (
            <Text style={{ color: tokens.colorNeutralForeground3 }}>
              Loading columns...
            </Text>
          )}
        </div>

        {/* Footer */}
        <div className={styles.footer}>
          <Button appearance="secondary" onClick={onClose} disabled={saving}>
            Cancel
          </Button>
          <Button appearance="primary" onClick={handleSave} disabled={saving}>
            {saving ? 'Saving...' : 'Save'}
          </Button>
        </div>
      </DrawerBody>
    </OverlayDrawer>
  );
}
